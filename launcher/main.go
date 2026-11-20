// Command calcplus is the local launcher for the calcplus browser app.
//
// It serves the static build on the loopback interface only and opens the
// default browser. It is a distribution shell: no browser engine, no native
// menus and no outbound calls. The contract is
// `.scratch/calcplus/spec/distribution.md` §2 and §8.
package main

import (
	"context"
	"embed"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/fs"
	"net"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"os/signal"
	"path"
	"runtime"
	"strconv"
	"strings"
	"time"
)

const (
	// defaultPort is the contract's stable preferred port. Keeping one origin
	// lets IndexedDB, the recent list and remembered file handles survive a
	// restart (distribution.md §2 and acceptance criterion 10).
	defaultPort = 47823
	// loopbackHost is the literal IPv4 loopback address. The server never
	// binds to a LAN interface (distribution.md acceptance criterion 2).
	loopbackHost = "127.0.0.1"
	// healthPath is how a second launch recognises the running instance.
	healthPath = "/__calcplus/health"
)

// version is stamped by scripts/build-launcher.mjs from package.json.
var version = "0.1.0"

// buildID is stamped from the embedded app's index.html hash, so two builds
// of the same version are distinguishable (a stale instance must never be
// silently reused).
var buildID = "dev"

// bundled carries the staged static build. scripts/build-launcher.mjs copies
// the app's dist/ into webapp/dist before `go build`; a plain source checkout
// only has webapp/README.md and builds a binary that explains the missing
// build.
//
//go:embed all:webapp
var bundled embed.FS

var errNoBuild = errors.New(
	"this launcher has no app build embedded; run `node scripts/build-launcher.mjs` " +
		"(it builds from dist/, so run `npm run build` first)",
)

func main() {
	err := run(os.Args[1:], os.Stdout)
	if err == nil || errors.Is(err, flag.ErrHelp) {
		return
	}
	fmt.Fprintf(os.Stderr, "calcplus: %v\n", err)
	os.Exit(1)
}

func run(args []string, out io.Writer) error {
	flags := flag.NewFlagSet("calcplus", flag.ContinueOnError)
	flags.SetOutput(out)
	port := flags.Int("port", defaultPort, "loopback port to use")
	distDir := flags.String("dist", "", "serve a build from this directory instead of the embedded build (development)")
	noOpen := flags.Bool("no-open", false, "do not open the default browser")
	showVersion := flags.Bool("version", false, "print the version and exit")
	if err := flags.Parse(args); err != nil {
		return err
	}
	if *showVersion {
		fmt.Fprintf(out, "calcplus %s\n", version)
		return nil
	}
	if *port < 1 || *port > 65535 {
		return fmt.Errorf("--port %d is not usable; pass a number from 1 to 65535", *port)
	}

	root, err := staticRoot(*distDir)
	if err != nil {
		return err
	}

	address := net.JoinHostPort(loopbackHost, strconv.Itoa(*port))
	appURL := "http://" + address + "/"

	// A second launch reuses a running instance of the same build rather than
	// failing or picking a random port, which would change the browser origin.
	// A different build on the port is refused: reusing it would silently open
	// a stale app.
	if runningBuild, ok := runningInstance(*port); ok {
		if runningBuild != buildID {
			return fmt.Errorf(
				"a different build of calcplus is already running at %s\n"+
					"Close it first (pkill -x calcplus) and start again, or run with --port N.\n"+
					"A different port has its own browser origin and storage area",
				appURL,
			)
		}
		fmt.Fprintf(out, "calcplus is already running at %s\n", appURL)
		if !*noOpen {
			openBrowser(appURL, out)
		}
		return nil
	}

	listener, err := net.Listen("tcp4", address)
	if err != nil {
		return fmt.Errorf(
			"cannot listen on %s: %v\n"+
				"Another program may be using this port. Close it, or start again with --port N. "+
				"A different port has its own browser origin and storage area",
			address, err,
		)
	}

	server := &http.Server{
		Handler:           newHandler(root),
		ReadHeaderTimeout: 5 * time.Second,
		IdleTimeout:       60 * time.Second,
	}

	fmt.Fprintf(out, "calcplus %s\n", version)
	fmt.Fprintf(out, "Serving the local build at %s\n", appURL)
	fmt.Fprintln(out, "Local only: reachable from this computer only, and no data leaves it.")
	if *port != defaultPort {
		fmt.Fprintf(
			out,
			"Note: port %d has its own browser origin, so recent files, file handles and "+
				"the recovery copy from port %d are not available here.\n",
			*port, defaultPort,
		)
	}
	fmt.Fprintln(out, "Close this window, or press Ctrl+C, to stop calcplus.")
	if !*noOpen {
		openBrowser(appURL, out)
	}

	stop := make(chan os.Signal, 1)
	signal.Notify(stop, os.Interrupt)
	defer signal.Stop(stop)

	served := make(chan error, 1)
	go func() { served <- server.Serve(listener) }()

	select {
	case err := <-served:
		if errors.Is(err, http.ErrServerClosed) {
			return nil
		}
		return err
	case <-stop:
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cancel()
		if err := server.Shutdown(ctx); err != nil {
			return err
		}
		fmt.Fprintln(out, "calcplus stopped.")
		return nil
	}
}

// staticRoot picks the embedded build, or a directory for development and
// tests. It fails early with an actionable message when the binary carries no
// app, rather than serving an empty site.
func staticRoot(distDir string) (fs.FS, error) {
	if distDir != "" {
		info, err := os.Stat(distDir)
		if err != nil || !info.IsDir() {
			return nil, fmt.Errorf("cannot serve %q: it is not a readable directory", distDir)
		}
		return os.DirFS(distDir), nil
	}
	build, err := fs.Sub(bundled, "webapp/dist")
	if err != nil {
		return nil, errNoBuild
	}
	if _, err := fs.Stat(build, "index.html"); err != nil {
		return nil, errNoBuild
	}
	return build, nil
}

// runningInstance reports the build id of a calcplus instance on the port.
func runningInstance(port int) (string, bool) {
	client := http.Client{Timeout: 750 * time.Millisecond}
	response, err := client.Get(
		"http://" + net.JoinHostPort(loopbackHost, strconv.Itoa(port)) + healthPath,
	)
	if err != nil {
		return "", false
	}
	defer response.Body.Close()
	var payload struct {
		App   string `json:"app"`
		Build string `json:"build"`
	}
	if err := json.NewDecoder(response.Body).Decode(&payload); err != nil {
		return "", false
	}
	if payload.App != "calcplus" {
		return "", false
	}
	return payload.Build, true
}

// newHandler serves the build and the instance health probe. Directory
// listings are never produced: a directory without index.html is a 404.
func newHandler(root fs.FS) http.Handler {
	files := http.FileServer(http.FS(root))
	static := http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("X-Content-Type-Options", "nosniff")
		w.Header().Set("Referrer-Policy", "no-referrer")
		w.Header().Set("X-Frame-Options", "SAMEORIGIN")
		if r.Method != http.MethodGet && r.Method != http.MethodHead {
			w.Header().Set("Allow", "GET, HEAD")
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		name, ok := resolveFile(root, r.URL.Path)
		if !ok {
			http.NotFound(w, r)
			return
		}
		if strings.HasSuffix(name, "/") || strings.HasSuffix(name, ".html") {
			// A replaced launcher must not serve a stale page from cache.
			w.Header().Set("Cache-Control", "no-cache")
		}
		request := r.Clone(r.Context())
		request.URL = &url.URL{Path: name}
		files.ServeHTTP(w, request)
	})

	mux := http.NewServeMux()
	mux.Handle(healthPath, healthHandler())
	mux.Handle("/", static)
	return mux
}

func healthHandler() http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet && r.Method != http.MethodHead {
			w.Header().Set("Allow", "GET, HEAD")
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		w.Header().Set("Cache-Control", "no-store")
		if r.Method == http.MethodHead {
			return
		}
		_ = json.NewEncoder(w).Encode(map[string]string{
			"app":     "calcplus",
			"version": version,
			"build":   buildID,
		})
	})
}

// resolveFile maps a request path to a servable path inside root. A directory
// resolves only when it has an index.html; the directory path is returned so
// the file server serves that index without a redirect and never a listing.
func resolveFile(root fs.FS, requestPath string) (string, bool) {
	name := strings.TrimPrefix(path.Clean("/"+requestPath), "/")
	if name == "" {
		name = "index.html"
	}
	if !fs.ValidPath(name) {
		return "", false
	}
	info, err := fs.Stat(root, name)
	if err != nil {
		return "", false
	}
	if info.IsDir() {
		index := path.Join(name, "index.html")
		if _, err := fs.Stat(root, index); err != nil {
			return "", false
		}
		return "/" + name + "/", true
	}
	if name == "index.html" {
		return "/", true
	}
	return "/" + name, true
}

// openBrowser starts the default browser. A failure is not fatal: the URL is
// already printed so the user can open it by hand.
func openBrowser(target string, out io.Writer) {
	var command string
	var args []string
	switch runtime.GOOS {
	case "windows":
		command, args = "rundll32.exe", []string{"url.dll,FileProtocolHandler", target}
	case "darwin":
		command, args = "open", []string{target}
	default:
		command, args = "xdg-open", []string{target}
	}
	if err := exec.Command(command, args...).Start(); err != nil {
		fmt.Fprintf(out, "Could not open the browser automatically; open %s yourself.\n", target)
	}
}
