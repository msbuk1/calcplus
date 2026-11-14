export class WorkbookOpenError extends Error {
  readonly code:
    'encrypted' | 'unsupported-format' | 'corrupt' | 'too-large' | 'strict-ooxml' | 'unsafe-xml';

  constructor(code: WorkbookOpenError['code'], message: string) {
    super(message);
    this.name = 'WorkbookOpenError';
    this.code = code;
  }
}

export class WorkbookSaveError extends Error {
  readonly code: 'unsafe-part' | 'unsupported-edit';

  constructor(code: WorkbookSaveError['code'], message: string) {
    super(message);
    this.name = 'WorkbookSaveError';
    this.code = code;
  }
}
