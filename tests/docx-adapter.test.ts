import { beforeEach, describe, expect, it, vi } from 'vitest';
import { DocxAdapter } from '../src/adapters/docx-adapter';

const mocks = vi.hoisted(() => ({
  createRoot: vi.fn(),
  render: vi.fn(),
  unmount: vi.fn(),
  editor: {
    on: vi.fn(() => vi.fn()),
  },
  editorRef: {
    save: vi.fn(),
  },
}));

function initializeRenderedEditor(element: { props: { onReady: (editor: unknown) => void }; ref: unknown }): void {
  (element.ref as (value: unknown) => void)(mocks.editorRef);
  queueMicrotask(() => element.props.onReady(mocks.editor));
}

vi.mock('react-dom/client', () => ({
  createRoot: mocks.createRoot,
}));

vi.mock('@docx-editor.dev/react', () => ({
  DocxEditor: () => null,
}));

describe('DocxAdapter', () => {
  beforeEach(() => {
    document.body.replaceChildren();
    mocks.createRoot.mockReset();
    mocks.render.mockReset();
    mocks.unmount.mockReset();
    mocks.editor.on.mockReset().mockReturnValue(vi.fn());
    mocks.editorRef.save.mockReset();
    mocks.createRoot.mockReturnValue({ render: mocks.render, unmount: mocks.unmount });
  });

  it('mounts the current DOCX editor chrome and resolves when the engine is ready', async () => {
    mocks.render.mockImplementation((element) => {
      initializeRenderedEditor(element);
    });

    const container = document.createElement('div');
    const adapter = new DocxAdapter(container, {
      type: 'docx',
      fileName: 'report.docx',
      docx: { author: 'Ada', initialZoom: 1, showRuler: true },
    });

    await adapter.loadFile(new ArrayBuffer(4));

    const element = mocks.render.mock.calls[0]?.[0];
    expect(element.props).toMatchObject({
      title: 'report.docx',
      author: 'Ada',
      mode: 'edit',
      zoom: 1,
      zoomMode: { type: 'fixed' },
      chrome: true,
      menu: false,
      rulers: true,
    });
    expect(element.props.document).toBeInstanceOf(Uint8Array);
    expect(adapter.state).toBe('ready');
  });

  it('uses the surface-only DOCX editor when the host owns unified chrome', async () => {
    mocks.render.mockImplementation((element) => {
      initializeRenderedEditor(element);
    });

    const adapter = new DocxAdapter(document.createElement('div'), {
      type: 'docx',
      chrome: { toolbar: 'unified' },
      readOnly: true,
    });

    await adapter.loadFile(new ArrayBuffer(4));

    const element = mocks.render.mock.calls[0]?.[0];
    expect(element.props).toMatchObject({ mode: 'view', chrome: false });
  });

  it('passes the current editor facade to the configured collaboration binding', async () => {
    mocks.render.mockImplementation((element) => {
      initializeRenderedEditor(element);
    });
    const bind = vi.fn().mockResolvedValue(undefined);

    const adapter = new DocxAdapter(document.createElement('div'), {
      type: 'docx',
      collab: { docId: 'document-42', docxBinding: { bind } },
    });

    await adapter.loadFile(new ArrayBuffer(4));

    expect(bind).toHaveBeenCalledWith({
      docId: 'document-42',
      editor: mocks.editor,
      provider: undefined,
      readOnly: false,
    });
  });

  it('exports through the current editor ref and unmounts on destroy', async () => {
    mocks.render.mockImplementation((element) => {
      initializeRenderedEditor(element);
    });
    mocks.editorRef.save.mockResolvedValue(new ArrayBuffer(4));

    const adapter = new DocxAdapter(document.createElement('div'), { type: 'docx' });
    await adapter.loadFile(new ArrayBuffer(4));

    const output = await adapter.exportFile();
    adapter.destroy();

    expect(output.type).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    );
    expect(output.size).toBe(4);
    expect(mocks.unmount).toHaveBeenCalledOnce();
  });
});
