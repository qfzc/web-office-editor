import { DocSDKError } from './core/errors';
import type { DocSDKOptions, IDocAdapter } from './types';

export async function createDocEditor(
  container: HTMLElement,
  options: DocSDKOptions,
): Promise<IDocAdapter> {
  if (!container || typeof container.attachShadow !== 'function') {
    throw new DocSDKError('INVALID_CONTAINER', 'createDocEditor requires an HTMLElement.');
  }

  switch (options.type) {
    case 'docx': {
      const { DocxAdapter } = await import('./adapters/docx-adapter');
      return new DocxAdapter(container, options);
    }
    case 'xlsx': {
      const { XlsxAdapter } = await import('./adapters/xlsx-adapter');
      return new XlsxAdapter(container, options);
    }
    case 'pptx': {
      const { PptxAdapter } = await import('./adapters/pptx-adapter');
      return new PptxAdapter(container, options);
    }
    default:
      throw new DocSDKError(
        'UNSUPPORTED_DOC_TYPE',
        `Unsupported document type: ${String(options.type)}`,
      );
  }
}

export { DocSDKError } from './core/errors';
export type { DocSDKErrorCode } from './core/errors';
export type {
  AdapterState,
  CollabConfig,
  DocSDKOptions,
  DocType,
  EditorChromeOptions,
  DocxAdapterOptions,
  DocxCollabBinding,
  DocxCollabBindingContext,
  IDocAdapter,
  OpChannel,
  OpChannelContext,
  PptxAdapterOptions,
  XlsxAdapterOptions,
  YjsCollabConfig,
  YProviderLike,
} from './types';
