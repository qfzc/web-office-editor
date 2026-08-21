import { describe, expect, it } from 'vitest';
import { DocSDKError } from '../src/core/errors';
import { createShadowMount } from '../src/core/shadow-mount';

describe('createShadowMount', () => {
  it('isolates styles and reuses an existing open shadow root', () => {
    const container = document.createElement('section');
    container.append(document.createElement('span'));

    const first = createShadowMount(container, 'docx', '.editor { color: red; }');
    const shadow = first.shadowRoot;

    expect(container.childNodes).toHaveLength(0);
    expect(first.mountPoint.dataset.docSdkMount).toBe('docx');
    expect(shadow.textContent).toContain('.editor { color: red; }');

    const second = createShadowMount(container, 'xlsx', '.sheet { color: blue; }');
    expect(second.shadowRoot).toBe(shadow);
    expect(shadow.textContent).not.toContain('.editor { color: red; }');
    expect(shadow.textContent).toContain('.sheet { color: blue; }');
  });

  it('rejects a container with a closed shadow root', () => {
    const container = document.createElement('div');
    container.attachShadow({ mode: 'closed' });

    expect(() => createShadowMount(container, 'docx')).toThrowError(DocSDKError);
  });
});
