import { beforeEach, describe, expect, it, vi } from 'vitest';
import { PptxAdapter } from '../src/adapters/pptx-adapter';

interface ViewerOptionsMock {
  source?: Blob | ArrayBuffer;
  onLoad?: (info: { slideCount: number; canvasSize: { width: number; height: number } }) => void;
  onError?: (message: string, error: unknown) => void;
}

const mocks = vi.hoisted(() => ({
  createPptxViewer: vi.fn(),
}));

vi.mock('pptx-vanilla-viewer', () => ({
  createPptxViewer: mocks.createPptxViewer,
}));

function addBackstage(container: HTMLElement): HTMLButtonElement {
  const backstage = document.createElement('div');
  backstage.className = 'pptxv-backstage';
  const back = document.createElement('button');
  back.className = 'pptxv-bs-back';
  back.addEventListener('click', () => {
    backstage.hidden = true;
  });
  backstage.append(back);
  container.append(backstage);
  return back;
}

describe('PptxAdapter', () => {
  beforeEach(() => {
    document.body.replaceChildren();
    mocks.createPptxViewer.mockReset();
  });

  it('opens the first PPTX as the viewer source and closes the File backstage', async () => {
    let viewerOptions: ViewerOptionsMock | undefined;
    const backClick = vi.fn();
    mocks.createPptxViewer.mockImplementation(
      (mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        viewerOptions = options;
        const back = addBackstage(mountPoint);
        back.addEventListener('click', backClick);
        queueMicrotask(() =>
          options.onLoad?.({ slideCount: 1, canvasSize: { width: 1280, height: 720 } }),
        );
        return {
          loadFile: vi.fn(),
          save: vi.fn(),
          startCollaboration: vi.fn(),
          stopCollaboration: vi.fn(),
          destroy: vi.fn(),
        };
      },
    );

    const container = document.createElement('div');
    document.body.append(container);
    const adapter = new PptxAdapter(container, { type: 'pptx' });
    const file = new Blob(['pptx']);

    await adapter.loadFile(file);

    expect(viewerOptions?.source).toBe(file);
    expect(adapter.state).toBe('ready');
    expect(backClick).toHaveBeenCalledOnce();
    expect(container.shadowRoot?.querySelector<HTMLElement>('.pptxv-backstage')?.hidden).toBe(
      true,
    );
    expect(container.shadowRoot?.textContent).toContain(
      '.pptxv-backstage[hidden] { display: none !important; }',
    );
    expect(container.shadowRoot?.textContent).toContain(
      '.pptxv-titlebar { display: none !important; }',
    );
    expect(container.shadowRoot?.textContent).toContain(
      '.pptxv-ribbon-primary { display: none !important; }',
    );
    expect(container.shadowRoot?.textContent).toContain(
      '.pptxv-ribbon-tabs { display: none !important; }',
    );
  });

  it('hides the native toolbar when a host renders the unified chrome', async () => {
    let viewerOptions: ViewerOptionsMock | undefined;
    mocks.createPptxViewer.mockImplementation(
      (_mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        viewerOptions = options;
        queueMicrotask(() =>
          options.onLoad?.({ slideCount: 1, canvasSize: { width: 1280, height: 720 } }),
        );
        return {
          loadFile: vi.fn(),
          save: vi.fn(),
          startCollaboration: vi.fn(),
          stopCollaboration: vi.fn(),
          destroy: vi.fn(),
        };
      },
    );

    const container = document.createElement('div');
    const adapter = new PptxAdapter(container, {
      type: 'pptx',
      chrome: { toolbar: 'unified' },
    });

    await adapter.loadFile(new Blob(['pptx']));

    expect(viewerOptions).toMatchObject({ showToolbar: false });
  });

  it('rejects loadFile when the viewer reports a parsing error', async () => {
    mocks.createPptxViewer.mockImplementation(
      (_mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        queueMicrotask(() => options.onError?.('Invalid PPTX package', new Error('bad zip')));
        return {
          loadFile: vi.fn(),
          save: vi.fn(),
          startCollaboration: vi.fn(),
          stopCollaboration: vi.fn(),
          destroy: vi.fn(),
        };
      },
    );

    const container = document.createElement('div');
    const onError = vi.fn();
    const adapter = new PptxAdapter(container, { type: 'pptx', onError });

    await expect(adapter.loadFile(new Blob(['not a pptx']))).rejects.toMatchObject({
      code: 'FILE_LOAD_FAILED',
      message: 'Invalid PPTX package',
    });
    expect(adapter.state).toBe('error');
    expect(onError).toHaveBeenCalledOnce();
  });

  it('closes the backstage after a file is opened from the viewer UI', async () => {
    let viewerOptions: ViewerOptionsMock | undefined;
    mocks.createPptxViewer.mockImplementation(
      (_mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        viewerOptions = options;
        queueMicrotask(() =>
          options.onLoad?.({ slideCount: 1, canvasSize: { width: 1280, height: 720 } }),
        );
        return {
          loadFile: vi.fn(),
          save: vi.fn(),
          startCollaboration: vi.fn(),
          stopCollaboration: vi.fn(),
          destroy: vi.fn(),
        };
      },
    );

    const container = document.createElement('div');
    const adapter = new PptxAdapter(container, { type: 'pptx' });
    await adapter.loadFile(new Blob(['pptx']));

    const mountPoint = container.shadowRoot?.querySelector<HTMLElement>('[data-doc-sdk-mount]');
    expect(mountPoint).toBeTruthy();
    const back = addBackstage(mountPoint!);
    const backClick = vi.fn();
    back.addEventListener('click', backClick);

    viewerOptions?.onLoad?.({ slideCount: 2, canvasSize: { width: 1280, height: 720 } });

    expect(backClick).toHaveBeenCalledOnce();
    expect(mountPoint?.querySelector<HTMLElement>('.pptxv-backstage')?.hidden).toBe(true);
  });

  it('hides the source text while the vanilla viewer inline editor is open', async () => {
    mocks.createPptxViewer.mockImplementation(
      (mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        const stage = document.createElement('div');
        stage.className = 'pptxv-stage';
        const source = document.createElement('div');
        source.dataset.elementId = 'title';
        const text = document.createElement('div');
        text.className = 'pptxv-text';
        text.textContent = 'Changsha';
        source.append(text);
        stage.append(source);
        stage.addEventListener('dblclick', () => {
          const editor = document.createElement('div');
          editor.className = 'pptxv-inline-text-editor';
          editor.dataset.inlineEditor = '';
          mountPoint.append(editor);
        });
        mountPoint.append(stage);
        queueMicrotask(() =>
          options.onLoad?.({ slideCount: 1, canvasSize: { width: 1280, height: 720 } }),
        );
        return {
          loadFile: vi.fn(),
          save: vi.fn(),
          startCollaboration: vi.fn(),
          stopCollaboration: vi.fn(),
          destroy: vi.fn(),
        };
      },
    );

    const container = document.createElement('div');
    document.body.append(container);
    const adapter = new PptxAdapter(container, { type: 'pptx' });
    await adapter.loadFile(new Blob(['pptx']));

    const source = container.shadowRoot?.querySelector<HTMLElement>('[data-element-id="title"]');
    source?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    await Promise.resolve();

    const editor = container.shadowRoot?.querySelector<HTMLElement>('.pptxv-inline-text-editor');
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(true);
    expect(editor).toBeTruthy();

    editor?.dispatchEvent(new FocusEvent('blur'));
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(false);
  });
});
