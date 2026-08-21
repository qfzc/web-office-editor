import { DocSDKError, wrapError } from './errors';
import { createShadowMount, type ShadowMount } from './shadow-mount';
import type {
  AdapterState,
  DocSDKOptions,
  DocType,
  IDocAdapter,
} from '../types';

export abstract class BaseAdapter implements IDocAdapter {
  readonly type: DocType;
  protected readonly options: DocSDKOptions;
  protected readonly mount: ShadowMount;
  private currentState: AdapterState = 'idle';

  protected constructor(
    type: DocType,
    container: HTMLElement,
    options: DocSDKOptions,
    css = '',
  ) {
    this.type = type;
    this.options = options;
    this.mount = createShadowMount(container, type, css);
  }

  get state(): AdapterState {
    return this.currentState;
  }

  abstract loadFile(file: Blob | ArrayBuffer): Promise<void>;
  abstract exportFile(): Promise<Blob>;
  abstract destroy(): void;

  protected startLoading(): void {
    this.assertAlive();
    this.currentState = 'loading';
  }

  protected markReady(): void {
    this.assertAlive();
    this.currentState = 'ready';
    this.invokeCallback('ready', this.options.onReady);
  }

  protected markChanged(): void {
    if (this.currentState === 'destroyed') return;
    this.invokeCallback('change', this.options.onChange);
  }

  protected markDestroyed(): void {
    this.currentState = 'destroyed';
    this.mount.shadowRoot.replaceChildren();
  }

  protected assertAlive(): void {
    if (this.currentState === 'destroyed') {
      throw new DocSDKError('ADAPTER_DESTROYED', `${this.type} adapter has been destroyed.`);
    }
  }

  protected assertReady(): void {
    this.assertAlive();
    if (this.currentState !== 'ready') {
      throw new DocSDKError(
        'ENGINE_NOT_READY',
        `${this.type} editor must finish loading before this operation.`,
      );
    }
  }

  protected fail(
    code: 'FILE_LOAD_FAILED' | 'EXPORT_FAILED' | 'COLLAB_FAILED',
    message: string,
    cause: unknown,
  ): DocSDKError {
    const error = wrapError(code, message, cause);
    if (code === 'FILE_LOAD_FAILED' && this.currentState !== 'destroyed') {
      this.currentState = 'error';
    }
    this.reportError(error);
    return error;
  }

  protected reportError(error: Error): void {
    try {
      this.options.onError?.(error);
    } catch {
      // Host callbacks must not corrupt adapter cleanup or error propagation.
    }
  }

  private invokeCallback(name: string, callback: (() => void) | undefined): void {
    try {
      callback?.();
    } catch (cause) {
      this.reportError(
        new DocSDKError('CALLBACK_FAILED', `The host ${name} callback failed.`, { cause }),
      );
    }
  }
}
