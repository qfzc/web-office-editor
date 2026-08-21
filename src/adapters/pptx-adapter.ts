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
const INLINE_TEXT_STYLE_PROPERTIES = [
  'color',
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'font-variant',
  'font-stretch',
  'font-kerning',
  'font-feature-settings',
  'font-variation-settings',
  'line-height',
  'letter-spacing',
  'word-spacing',
  'text-align',
  'text-align-last',
  'text-indent',
  'text-transform',
  'text-decoration',
  'text-decoration-color',
  'text-decoration-line',
  'text-decoration-style',
  'text-decoration-thickness',
  'text-underline-offset',
  'text-shadow',
  'white-space',
  'overflow-wrap',
  'word-break',
  'hyphens',
  'direction',
  'writing-mode',
  'text-orientation',
  'vertical-align',
] as const;

interface PendingLoad {
  promise: Promise<void>;
  resolve(): void;
  reject(error: Error): void;
}

function getInlineEditorText(element: Element): string {
  let text = '';
  const visit = (parent: Node) => {
    for (const child of Array.from(parent.childNodes)) {
      if (child.nodeType === 3) {
        text += child.nodeValue ?? '';
        continue;
      }
      if (!(child instanceof HTMLElement)) continue;
      if (child.tagName === 'BR') {
        text += '\n';
        continue;
      }
      if ((child.tagName === 'DIV' || child.tagName === 'P') && text.length > 0 && !text.endsWith('\n')) {
        text += '\n';
      }
      visit(child);
    }
  };

  visit(element);
  return text;
}

function copyTextAppearance(source: HTMLElement, inlineEditor: HTMLElement): void {
  const sourceStyle = source.ownerDocument.defaultView?.getComputedStyle(source);
  if (!sourceStyle) return;

  for (const property of INLINE_TEXT_STYLE_PROPERTIES) {
    const value = sourceStyle.getPropertyValue(property);
    if (value) inlineEditor.style.setProperty(property, value);
  }
}

function restoreInlineEditorCaret(inlineEditor: HTMLElement): void {
  const document = inlineEditor.ownerDocument;
  if (document.activeElement !== inlineEditor) return;

  const selection = document.defaultView?.getSelection();
  if (!selection) return;

  const range = document.createRange();
  range.selectNodeContents(inlineEditor);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function preserveInlineEditorSegments(
  sourceText: HTMLElement,
  inlineEditor: HTMLElement,
): void {
  const segments = Array.from(inlineEditor.querySelectorAll<HTMLElement>('[data-seg-idx]'));
  if (segments.length === 0) return;

  const sourceRuns = Array.from(sourceText.querySelectorAll<HTMLElement>('.pptxv-para > *')).filter(
    (run) =>
      !run.classList.contains('pptxv-bullet') &&
      !run.classList.contains('pptxv-bullet-image') &&
      run.textContent !== '',
  );
  let segmentIndex = 0;

  for (const run of sourceRuns) {
    while (segments[segmentIndex]?.textContent === '\n') segmentIndex += 1;
    if (segments[segmentIndex]?.textContent !== run.textContent) continue;
    run.dataset.segIdx = String(segmentIndex);
    segmentIndex += 1;
  }
}

function hydrateInlineEditorFromSource(source: HTMLElement, inlineEditor: HTMLElement): void {
  const sourceText = source.matches('.pptxv-text')
    ? source
    : source.querySelector<HTMLElement>('.pptxv-text');
  if (!sourceText) return;

  copyTextAppearance(sourceText, inlineEditor);
  if (getInlineEditorText(sourceText) !== getInlineEditorText(inlineEditor)) return;

  const clone = sourceText.cloneNode(true) as HTMLElement;
  // The editor should retain link styling without navigating away while typing.
  for (const link of clone.querySelectorAll('a')) {
    link.removeAttribute('href');
    link.removeAttribute('target');
    link.removeAttribute('rel');
  }
  preserveInlineEditorSegments(clone, inlineEditor);
  inlineEditor.replaceChildren(clone);
  restoreInlineEditorCaret(inlineEditor);
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
    let activeInlineEditor: HTMLElement | null = null;

    const getInlineEditor = () =>
      mountPoint.querySelector<HTMLElement>('.pptxv-inline-text-editor[data-inline-editor]');
    const isTextSource = (element: HTMLElement) =>
      element.classList.contains('pptxv-text') ||
      element.classList.contains('pptxv-wordart') ||
      element.querySelector('.pptxv-text, .pptxv-wordart') !== null;
    const findSourceText = (
      inlineEditor: HTMLElement,
      preferredSource: HTMLElement | null,
    ): HTMLElement | null => {
      if (preferredSource && mountPoint.contains(preferredSource) && isTextSource(preferredSource)) {
        return preferredSource;
      }

      const selectionBox = inlineEditor
        .closest('.pptxv-editor-overlay')
        ?.querySelector<HTMLElement>('.pptxv-sel-box:not([hidden])');
      const editorRect = (selectionBox ?? inlineEditor).getBoundingClientRect();
      if (editorRect.width === 0 || editorRect.height === 0) return null;

      let closestSource: HTMLElement | null = null;
      let closestDistance = Number.POSITIVE_INFINITY;
      for (const source of mountPoint.querySelectorAll<HTMLElement>('[data-element-id]')) {
        if (!isTextSource(source)) continue;
        const sourceRect = source.getBoundingClientRect();
        if (sourceRect.width === 0 || sourceRect.height === 0) continue;

        const distance =
          Math.abs(sourceRect.left - editorRect.left) +
          Math.abs(sourceRect.top - editorRect.top) +
          Math.abs(sourceRect.width - editorRect.width) +
          Math.abs(sourceRect.height - editorRect.height);
        if (distance < closestDistance) {
          closestDistance = distance;
          closestSource = source;
        }
      }

      return closestDistance <= 8 ? closestSource : null;
    };

    const hideSourceText = (inlineEditor: HTMLElement, preferredSource: HTMLElement | null) => {
      if (activeInlineEditor === inlineEditor) return;
      restoreInlineText?.();

      const source = findSourceText(inlineEditor, preferredSource);
      if (!source) return;

      hydrateInlineEditorFromSource(source, inlineEditor);
      source.setAttribute(INLINE_TEXT_SOURCE_ATTRIBUTE, '');
      activeInlineEditor = inlineEditor;
      let restored = false;
      const restore = () => {
        if (restored) return;
        restored = true;
        source.removeAttribute(INLINE_TEXT_SOURCE_ATTRIBUTE);
        activeInlineEditor = null;
        if (restoreInlineText === restore) restoreInlineText = null;
      };

      restoreInlineText = restore;
      inlineEditor.addEventListener('blur', restore, { once: true });
    };

    const hideSourceTextWhenInlineEditorOpens = (event: Event) => {
      if (!(event.target instanceof Element)) return;

      const source = event.target.closest<HTMLElement>('[data-element-id]') ?? null;

      queueMicrotask(() => {
        const inlineEditor = getInlineEditor();
        if (inlineEditor) hideSourceText(inlineEditor, source);
      });
    };

    const observer = new MutationObserver(() => {
      const inlineEditor = getInlineEditor();
      if (inlineEditor) hideSourceText(inlineEditor, null);
      else restoreInlineText?.();
    });

    // The viewer opens text editing on the second pointerdown, before the
    // browser dispatches dblclick. Keep dblclick as a fallback for viewers
    // that defer editor creation until that event.
    mountPoint.addEventListener('pointerdown', hideSourceTextWhenInlineEditorOpens, true);
    mountPoint.addEventListener('dblclick', hideSourceTextWhenInlineEditorOpens, true);
    observer.observe(mountPoint, { childList: true, subtree: true });
    this.disposeInlineTextLayerWorkaround = () => {
      mountPoint.removeEventListener('pointerdown', hideSourceTextWhenInlineEditorOpens, true);
      mountPoint.removeEventListener('dblclick', hideSourceTextWhenInlineEditorOpens, true);
      observer.disconnect();
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
