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
        stage.addEventListener('pointerdown', () => {
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
    source?.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await Promise.resolve();

    const editor = container.shadowRoot?.querySelector<HTMLElement>('.pptxv-inline-text-editor');
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(true);
    expect(editor).toBeTruthy();

    editor?.dispatchEvent(new FocusEvent('blur'));
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(false);
  });

  it('uses the rendered slide text structure in the inline editor', async () => {
    mocks.createPptxViewer.mockImplementation(
      (mountPoint: HTMLElement, options: ViewerOptionsMock) => {
        const stage = document.createElement('div');
        const source = document.createElement('div');
        source.dataset.elementId = 'title';
        const text = document.createElement('div');
        text.className = 'pptxv-text';
        text.style.color = '#be5347';
        text.style.fontFamily = 'Georgia';
        text.style.fontSize = '44px';
        text.style.lineHeight = '1.15';
        const paragraph = document.createElement('p');
        paragraph.className = 'pptxv-para';
        paragraph.style.textAlign = 'center';
        const run = document.createElement('span');
        run.style.color = '#f3ead1';
        run.style.fontWeight = '700';
        run.textContent = 'Changsha: Mountain City';
        paragraph.append(run);
        text.append(paragraph);
        source.append(text);
        stage.append(source);
        stage.addEventListener('pointerdown', () => {
          const editor = document.createElement('div');
          editor.className = 'pptxv-inline-text-editor';
          editor.dataset.inlineEditor = '';
          editor.style.fontSize = '14px';
          const segment = document.createElement('span');
          segment.dataset.segIdx = '0';
          segment.textContent = 'Changsha: Mountain City';
          editor.append(segment);
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
    source?.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await Promise.resolve();

    const editor = container.shadowRoot?.querySelector<HTMLElement>('.pptxv-inline-text-editor');
    const clonedText = editor?.querySelector<HTMLElement>('.pptxv-text');
    const clonedParagraph = clonedText?.querySelector<HTMLElement>('.pptxv-para');
    const clonedRun = clonedParagraph?.querySelector<HTMLElement>('span');
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(true);
    expect(editor?.style.fontSize).toBe('44px');
    expect(clonedText).toBeTruthy();
    expect(clonedParagraph?.style.textAlign).toBe('center');
    expect(clonedRun).toMatchObject({ textContent: 'Changsha: Mountain City' });
    expect(clonedRun?.style.fontWeight).toBe('700');
    expect(clonedRun?.dataset.segIdx).toBe('0');
  });

  it('matches an inline editor to its source text when the event target is unavailable', async () => {
    let mountPoint: HTMLElement | undefined;
    mocks.createPptxViewer.mockImplementation(
      (mount: HTMLElement, options: ViewerOptionsMock) => {
        mountPoint = mount;
        const source = document.createElement('div');
        source.dataset.elementId = 'title';
        source.getBoundingClientRect = () =>
          ({ left: 120, top: 80, width: 360, height: 72 } as DOMRect);
        const text = document.createElement('div');
        text.className = 'pptxv-text';
        source.append(text);
        mount.append(source);
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

    const editor = document.createElement('div');
    editor.className = 'pptxv-inline-text-editor';
    editor.dataset.inlineEditor = '';
    editor.getBoundingClientRect = () =>
      ({ left: 120, top: 80, width: 360, height: 72 } as DOMRect);
    mountPoint?.append(editor);
    await Promise.resolve();

    const source = container.shadowRoot?.querySelector<HTMLElement>('[data-element-id="title"]');
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(true);

    editor.remove();
    await Promise.resolve();
    expect(source?.hasAttribute('data-doc-sdk-pptx-inline-source')).toBe(false);
  });
});
