export type DocType = 'docx' | 'xlsx' | 'pptx';

export type AdapterState = 'idle' | 'loading' | 'ready' | 'error' | 'destroyed';

/**
 * Controls the boundary between the SDK engine and a host application's
 * document workbench. `native` preserves the vendor toolbar (the default).
 * `unified` hides it so the host can provide one consistent toolbar.
 */
export interface EditorChromeOptions {
  toolbar?: 'native' | 'unified';
}

export type MaybePromise<T> = T | Promise<T>;

export interface OpChannelContext {
  docId: string;
}

/** A host-owned transport used by the XLSX operation-log adapter. */
export interface OpChannel {
  send(ops: readonly unknown[], context: OpChannelContext): MaybePromise<void>;
  onReceive(
    callback: (ops: readonly unknown[], context?: OpChannelContext) => void,
  ): void | (() => void);
}

/** Minimal lifecycle shared by y-websocket and compatible providers. */
export interface YProviderLike {
  connect?(): void;
  disconnect?(): void;
  destroy?(): void;
}

export interface DocxCollabBindingContext {
  docId: string;
  readOnly: boolean;
  editor: unknown;
  provider?: YProviderLike;
}

/**
 * Bridges the DOCX editor facade to the host's Yjs document. This remains
 * injectable because the upstream editor does not expose a provider-only
 * collaboration API.
 */
export interface DocxCollabBinding {
  bind(context: DocxCollabBindingContext): MaybePromise<void | (() => void)>;
}

export interface YjsCollabConfig {
  serverUrl: string;
  roomId?: string;
  transport?: 'websocket' | 'webrtc';
  signaling?: string[];
  userName?: string;
  userAvatar?: string;
  userColor?: string;
  authToken?: string;
  role?: 'owner' | 'collaborator' | 'viewer';
  onWriteBack?: (bytes: Uint8Array) => void;
  writeBackDebounceMs?: number;
}

export interface CollabConfig {
  docId: string;
  /** Used by the PPTX adapter's built-in Yjs integration. */
  yjs?: YjsCollabConfig;
  /** Optional host-owned provider passed to a DOCX collaboration binding. */
  yProvider?: YProviderLike;
  /** Required for DOCX collaboration. */
  docxBinding?: DocxCollabBinding;
  /** Required for XLSX collaboration. */
  opChannel?: OpChannel;
}

export interface DocxAdapterOptions {
  author?: string;
  showToolbar?: boolean;
  showRuler?: boolean;
  initialZoom?: number;
}

export interface XlsxAdapterOptions {
  showToolbar?: boolean;
  showFormulaBar?: boolean;
  showSheetTabs?: boolean;
  locale?: string;
}

export interface PptxAdapterOptions {
  showToolbar?: boolean;
  showThumbnails?: boolean;
  showInspector?: boolean;
  locale?: string;
  autosave?: boolean;
}

export interface DocSDKOptions {
  type: DocType;
  readOnly?: boolean;
  fileName?: string;
  chrome?: EditorChromeOptions;
  onReady?: () => void;
  onChange?: () => void;
  onError?: (error: Error) => void;
  collab?: CollabConfig;
  docx?: DocxAdapterOptions;
  xlsx?: XlsxAdapterOptions;
  pptx?: PptxAdapterOptions;
}

export interface IDocAdapter {
  readonly type: DocType;
  readonly state: AdapterState;
  loadFile(file: Blob | ArrayBuffer): Promise<void>;
  exportFile(): Promise<Blob>;
  enableCollab?(config: CollabConfig): Promise<void>;
  destroy(): void;
}
