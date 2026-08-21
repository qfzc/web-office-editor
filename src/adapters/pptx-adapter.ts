import {
  createPptxViewer,
  type CollaborationConfig as PptxCollaborationConfig,
  type PptxViewerInstance,
} from 'pptx-vanilla-viewer';
import pptxStyles from 'pptx-vanilla-viewer/styles.css?inline';
import { BaseAdapter } from '../core/base-adapter';
import { resolveNativeToolbarVisibility } from '../core/editor-chrome';
import { DocSDKError, asError } from '../core/errors';
import { toBlobPart } from '../core/file';
import type { CollabConfig, DocSDKOptions } from '../types';

const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const PPTX_STYLE_OVERRIDES = `
.pptxv-backstage[hidden] { display: none !important; }
/* The vanilla viewer renders a contenteditable copy over the source text. */
[data-doc-sdk-pptx-inline-source].pptxv-text,
[data-doc-sdk-pptx-inline-source] .pptxv-text,
[data-doc-sdk-pptx-inline-source].pptxv-wordart,
[data-doc-sdk-pptx-inline-source] .pptxv-wordart { visibility: hidden !important; }
/* The host workbench owns document-level navigation; retain the active command row. */
.pptxv-titlebar { display: none !important; }
.pptxv-ribbon-primary { display: none !important; }
.pptxv-ribbon-tabs { display: none !important; }
`;
const INLINE_TEXT_SOURCE_ATTRIBUTE = 'data-doc-sdk-pptx-inline-source';

interface PendingLoad {
  promise: Promise<void>;
  resolve(): void;
  reject(error: Error): void;
}

export class PptxAdapter extends BaseAdapter {
  private viewer: PptxViewerInstance | null = null;
  private collabConfig: CollabConfig | null = null;
  private pendingLoad: PendingLoad | null = null;
  private loadQueue: Promise<void> = Promise.resolve();
  private disposeInlineTextLayerWorkaround: (() => void) | null = null;

  constructor(container: HTMLElement, options: DocSDKOptions) {
    super('pptx', container, options, `${pptxStyles}\n${PPTX_STYLE_OVERRIDES}`);
  }

  loadFile(file: Blob | ArrayBuffer): Promise<void> {
    this.assertAlive();
    const load = this.loadQueue
      .catch(() => undefined)
      .then(() => this.performLoad(file));
    this.loadQueue = load;
    return load;
  }

  private async performLoad(file: Blob | ArrayBuffer): Promise<void> {
    this.startLoading();
    const completion = this.beginLoad();
    try {
      if (this.viewer) {
        await Promise.all([this.viewer.loadFile(file), completion.promise]);
      } else {
        this.ensureViewer(file);
        await completion.promise;
      }
      this.markReady();
      const collab = this.collabConfig ?? this.options.collab;
      if (collab) await this.enableCollab(collab);
    } catch (cause) {
      throw this.fail('FILE_LOAD_FAILED', 'Unable to load the PPTX file.', cause);
    } finally {
      if (this.pendingLoad === completion) this.pendingLoad = null;
    }
  }

  async exportFile(): Promise<Blob> {
    this.assertReady();
    try {
      const bytes = await this.viewer!.save();
      return new Blob([toBlobPart(bytes)], { type: PPTX_MIME });
    } catch (cause) {
      throw this.fail('EXPORT_FAILED', 'Unable to export the PPTX file.', cause);
    }
  }

  async enableCollab(config: CollabConfig): Promise<void> {
    this.assertAlive();
    if (!config.docId.trim()) {
      throw new DocSDKError('COLLAB_CONFIG_INVALID', 'PPTX collaboration requires docId.');
    }
    if (!config.yjs) {
      throw new DocSDKError(
        'COLLAB_CONFIG_INVALID',
        'PPTX collaboration requires a yjs configuration.',
      );
    }
    if (config.yjs.transport !== 'webrtc' && !config.yjs.serverUrl.trim()) {
      throw new DocSDKError(
        'COLLAB_CONFIG_INVALID',
        'PPTX websocket collaboration requires yjs.serverUrl.',
      );
    }

    const roomId = config.yjs.roomId ?? config.docId;
    if (!/^[A-Za-z0-9_-]+$/.test(roomId)) {
      throw new DocSDKError(
        'COLLAB_CONFIG_INVALID',
        'PPTX collaboration roomId may only contain letters, numbers, hyphens, and underscores.',
      );
    }

    this.collabConfig = config;
    if (!this.viewer) return;

    const pptxConfig: PptxCollaborationConfig = {
      roomId,
      serverUrl: config.yjs.serverUrl,
      transport: config.yjs.transport,
      signaling: config.yjs.signaling,
      userName: config.yjs.userName ?? 'Anonymous',
      userAvatar: config.yjs.userAvatar,
      userColor: config.yjs.userColor,
      authToken: config.yjs.authToken,
      role: config.yjs.role,
      onWriteBack: config.yjs.onWriteBack,
      writeBackDebounceMs: config.yjs.writeBackDebounceMs,
    };

    try {
      await this.viewer.startCollaboration(pptxConfig);
    } catch (cause) {
      throw this.fail('COLLAB_FAILED', 'Unable to enable PPTX collaboration.', cause);
    }
  }

  destroy(): void {
    if (this.state === 'destroyed') return;
    this.pendingLoad?.reject(
      new DocSDKError('ADAPTER_DESTROYED', 'pptx adapter has been destroyed.'),
    );
    this.pendingLoad = null;
    this.disposeInlineTextLayerWorkaround?.();
    this.disposeInlineTextLayerWorkaround = null;
    this.viewer?.stopCollaboration();
    this.viewer?.destroy();
    this.viewer = null;
    this.markDestroyed();
  }

  private ensureViewer(source: Blob | ArrayBuffer): PptxViewerInstance {
    if (this.viewer) return this.viewer;

    const ownerDocument = this.mount.mountPoint.ownerDocument;
    const globalStyleId = 'pptx-vanilla-viewer-styles';
    const existingGlobalStyles = ownerDocument.getElementById(globalStyleId);
    this.viewer = createPptxViewer(this.mount.mountPoint, {
      source,
      editable: !this.options.readOnly,
      readOnly: this.options.readOnly,
      fileName: this.options.fileName,
      locale: this.options.pptx?.locale,
      showToolbar: resolveNativeToolbarVisibility(
        this.options.chrome,
        this.options.pptx?.showToolbar,
      ),
      showThumbnails: this.options.pptx?.showThumbnails,
      showInspector: this.options.pptx?.showInspector,
      autosave: this.options.pptx?.autosave,
      autosaveFilePath: this.options.fileName,
      onChange: () => this.markChanged(),
      onLoad: () => {
        this.closeBackstage();
        const pending = this.pendingLoad;
        this.pendingLoad = null;
        pending?.resolve();
      },
      onError: (message, error) => {
        const loadError = new DocSDKError('FILE_LOAD_FAILED', message, {
          cause: asError(error),
        });
        const pending = this.pendingLoad;
        this.pendingLoad = null;
        if (pending) pending.reject(loadError);
        else this.reportError(loadError);
      },
    });
    this.installInlineTextLayerWorkaround();

    // The complete stylesheet already lives inside this adapter's ShadowRoot.
    if (!existingGlobalStyles) ownerDocument.getElementById(globalStyleId)?.remove();

    return this.viewer;
  }

  private installInlineTextLayerWorkaround(): void {
    this.disposeInlineTextLayerWorkaround?.();

    const mountPoint = this.mount.mountPoint;
    let restoreInlineText: (() => void) | null = null;

    const onDoubleClick = (event: MouseEvent) => {
      if (!(event.target instanceof Element)) return;

      const source = event.target.closest<HTMLElement>('[data-element-id]');
      if (!source || !mountPoint.contains(source)) return;

      queueMicrotask(() => {
        const inlineEditor = mountPoint.querySelector<HTMLElement>(
          '.pptxv-inline-text-editor[data-inline-editor]',
        );
        if (!inlineEditor || !mountPoint.contains(inlineEditor) || !mountPoint.contains(source)) {
          return;
        }

        restoreInlineText?.();
        source.setAttribute(INLINE_TEXT_SOURCE_ATTRIBUTE, '');

        let restored = false;
        const observer = new MutationObserver(() => {
          if (!mountPoint.contains(inlineEditor)) restore();
        });
        const restore = () => {
          if (restored) return;
          restored = true;
          observer.disconnect();
          source.removeAttribute(INLINE_TEXT_SOURCE_ATTRIBUTE);
          if (restoreInlineText === restore) restoreInlineText = null;
        };

        restoreInlineText = restore;
        inlineEditor.addEventListener('blur', restore, { once: true });
        observer.observe(mountPoint, { childList: true, subtree: true });
      });
    };

    mountPoint.addEventListener('dblclick', onDoubleClick, true);
    this.disposeInlineTextLayerWorkaround = () => {
      mountPoint.removeEventListener('dblclick', onDoubleClick, true);
      restoreInlineText?.();
      restoreInlineText = null;
    };
  }

  private beginLoad(): PendingLoad {
    let resolve!: () => void;
    let reject!: (error: Error) => void;
    const promise = new Promise<void>((resolvePromise, rejectPromise) => {
      resolve = resolvePromise;
      reject = rejectPromise;
    });
    const pending = { promise, resolve, reject };
    this.pendingLoad = pending;
    return pending;
  }

  private closeBackstage(): void {
    this.mount.mountPoint
      .querySelector<HTMLButtonElement>('.pptxv-backstage .pptxv-bs-back')
      ?.click();
  }
}
