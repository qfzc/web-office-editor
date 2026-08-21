import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { AdapterState, DocType, IDocAdapter } from '../src/types';

class StubAdapter implements IDocAdapter {
  readonly state: AdapterState = 'idle';

  constructor(
    readonly type: DocType,
    readonly container: HTMLElement,
  ) {}

  async loadFile(): Promise<void> {}
  async exportFile(): Promise<Blob> {
    return new Blob();
  }
  destroy(): void {}
}

vi.mock('../src/adapters/docx-adapter', () => ({
  DocxAdapter: class extends StubAdapter {
    constructor(container: HTMLElement) {
      super('docx', container);
    }
  },
}));

vi.mock('../src/adapters/xlsx-adapter', () => ({
  XlsxAdapter: class extends StubAdapter {
    constructor(container: HTMLElement) {
      super('xlsx', container);
    }
  },
}));

vi.mock('../src/adapters/pptx-adapter', () => ({
  PptxAdapter: class extends StubAdapter {
    constructor(container: HTMLElement) {
      super('pptx', container);
    }
  },
}));

describe('createDocEditor', () => {
  beforeEach(() => {
    document.body.replaceChildren();
  });

  it.each(['docx', 'xlsx', 'pptx'] as const)('loads the %s adapter on demand', async (type) => {
    const { createDocEditor } = await import('../src');
    const container = document.createElement('div');
    const editor = await createDocEditor(container, { type });

    expect(editor.type).toBe(type);
  });

  it('returns a typed unsupported-format error', async () => {
    const { createDocEditor } = await import('../src');
    const container = document.createElement('div');

    await expect(
      createDocEditor(container, { type: 'pdf' as DocType }),
    ).rejects.toMatchObject({
      code: 'UNSUPPORTED_DOC_TYPE',
    });
  });
});
