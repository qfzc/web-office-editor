import { DocxEditor, type DocxEditorRef } from '@docx-editor.dev/react';
import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import docxStyles from '@docx-editor.dev/core/styles/editor.css?inline';
import { createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { BaseAdapter } from '../core/base-adapter';
import { resolveNativeToolbarVisibility } from '../core/editor-chrome';
import { DocSDKError } from '../core/errors';
import { toArrayBuffer } from '../core/file';
import type { CollabConfig, DocSDKOptions } from '../types';

export class DocxAdapter extends BaseAdapter {
  private root: Root | null = null;
  private editor: Editor | null = null;
  private editorRef: DocxEditorRef | null = null;
  private editorErrorCleanup: (() => void) | null = null;
  private collabConfig: CollabConfig | null = null;
  private collabCleanup: (() => void) | null = null;

  constructor(container: HTMLElement, options: DocSDKOptions) {
    super('docx', container, options, docxStyles);
  }

  async loadFile(file: Blob | ArrayBuffer): Promise<void> {
    this.startLoading();
    this.cleanupEditor();
    this.mount.reset();

    try {
      const bytes = new Uint8Array(await toArrayBuffer(file));
      await this.renderDocument(bytes);
    } catch (cause) {
      this.cleanupEditor();
      throw this.fail('FILE_LOAD_FAILED', 'Unable to load the DOCX file.', cause);
    }
  }

  async exportFile(): Promise<Blob> {
    this.assertReady();
    try {
      const output = await this.editorRef?.save();
      if (!output) {
        throw new DocSDKError('EXPORT_FAILED', 'The DOCX engine returned no output.');
      }
      return new Blob([output], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    } catch (cause) {
      throw this.fail('EXPORT_FAILED', 'Unable to export the DOCX file.', cause);
    }
  }

  async enableCollab(config: CollabConfig): Promise<void> {
    this.assertAlive();
    if (!config.docId.trim()) {
      throw new DocSDKError('COLLAB_CONFIG_INVALID', 'DOCX collaboration requires docId.');
    }
    if (!config.docxBinding) {
      throw new DocSDKError(
        'COLLAB_CONFIG_INVALID',
        'DOCX collaboration requires a docxBinding that connects the editor facade to Yjs.',
      );
    }

    this.collabConfig = config;
    if (!this.editor) return;

    this.cleanupCollab();
    try {
      config.yProvider?.connect?.();
      const cleanup = await config.docxBinding.bind({
        docId: config.docId,
        editor: this.editor,
        provider: config.yProvider,
        readOnly: Boolean(this.options.readOnly),
      });
      this.collabCleanup = cleanup ?? null;
    } catch (cause) {
      config.yProvider?.disconnect?.();
      throw this.fail('COLLAB_FAILED', 'Unable to enable DOCX collaboration.', cause);
    }
  }

  destroy(): void {
    if (this.state === 'destroyed') return;
    this.cleanupEditor();
    this.markDestroyed();
  }

  private cleanupCollab(): void {
    try {
      this.collabCleanup?.();
    } finally {
      this.collabCleanup = null;
      this.collabConfig?.yProvider?.disconnect?.();
    }
  }

  private cleanupEditor(): void {
    this.cleanupCollab();
    this.editorErrorCleanup?.();
    this.editorErrorCleanup = null;
    this.root?.unmount();
    this.root = null;
    this.editor = null;
    this.editorRef = null;
  }

  private renderDocument(bytes: Uint8Array): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      let settled = false;
      const settle = (callback: () => void): void => {
        if (settled) return;
        settled = true;
        callback();
      };
      const rejectLoad = (error: Error): void => {
        if (settled) {
          this.reportError(error);
          return;
        }
        settled = true;
        reject(error);
      };

      this.root = createRoot(this.mount.mountPoint);
      this.root.render(
        createElement(DocxEditor, {
          document: bytes,
          title: this.options.fileName,
          author: this.options.docx?.author,
          mode: this.options.readOnly ? 'view' : 'edit',
          zoom: this.options.docx?.initialZoom,
          zoomMode: this.options.docx?.initialZoom === undefined ? 'auto' : { type: 'fixed' },
          chrome: resolveNativeToolbarVisibility(
            this.options.chrome,
            this.options.docx?.showToolbar,
          ),
          // The host workbench already owns file open/export; keep the native format toolbar
          // but remove its duplicate File/Format/Insert/Help menu row.
          menu: false,
          rulers: this.options.docx?.showRuler,
          ref: (editorRef: DocxEditorRef | null) => {
            this.editorRef = editorRef;
          },
          onChange: () => this.markChanged(),
          onFontError: (error: Error) => this.reportError(error),
          onReady: (editor: Editor) => {
            this.editor = editor;
            this.editorErrorCleanup = editor.on('error', rejectLoad);
            const collab = this.collabConfig ?? this.options.collab;
            Promise.resolve(collab ? this.enableCollab(collab) : undefined)
              .then(() => settle(() => {
                this.markReady();
                resolve();
              }))
              .catch((cause: unknown) => {
                rejectLoad(cause instanceof Error ? cause : new Error('DOCX collaboration failed.'));
              });
          },
        }),
      );
    });
  }
}
