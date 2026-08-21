export type DocSDKErrorCode =
  | 'INVALID_CONTAINER'
  | 'UNSUPPORTED_DOC_TYPE'
  | 'ADAPTER_DESTROYED'
  | 'ENGINE_NOT_READY'
  | 'FILE_LOAD_FAILED'
  | 'EXPORT_FAILED'
  | 'COLLAB_CONFIG_INVALID'
  | 'COLLAB_FAILED'
  | 'CALLBACK_FAILED';

export class DocSDKError extends Error {
  readonly code: DocSDKErrorCode;

  constructor(code: DocSDKErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'DocSDKError';
    this.code = code;
  }
}

export function asError(value: unknown): Error {
  if (value instanceof Error) return value;
  return new Error(typeof value === 'string' ? value : 'Unknown document SDK error');
}

export function wrapError(
  code: DocSDKErrorCode,
  message: string,
  cause: unknown,
): DocSDKError {
  if (cause instanceof DocSDKError) return cause;
  return new DocSDKError(code, message, { cause: asError(cause) });
}
