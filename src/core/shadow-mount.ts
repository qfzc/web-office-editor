import { DocSDKError } from './errors';

const BASE_STYLES = `
:host {
  display: block;
  width: 100%;
  height: 100%;
  min-width: 0;
  min-height: 320px;
  color: #171717;
  background: #ffffff;
  font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
*, *::before, *::after { box-sizing: border-box; }
[data-doc-sdk-mount] {
  width: 100%;
  height: 100%;
  min-width: 0;
  min-height: 320px;
  overflow: hidden;
}
`;

export interface ShadowMount {
  shadowRoot: ShadowRoot;
  mountPoint: HTMLDivElement;
  addStyles(css: string): void;
  reset(): void;
}

export function createShadowMount(
  container: HTMLElement,
  adapterName: string,
  css = '',
): ShadowMount {
  container.replaceChildren();

  let shadowRoot = container.shadowRoot;
  if (!shadowRoot) {
    try {
      shadowRoot = container.attachShadow({ mode: 'open' });
    } catch (cause) {
      throw new DocSDKError(
        'INVALID_CONTAINER',
        'The editor container must not own an inaccessible closed ShadowRoot.',
        { cause },
      );
    }
  }

  const build = (): HTMLDivElement => {
    shadowRoot.replaceChildren();
    const style = container.ownerDocument.createElement('style');
    style.dataset.docSdkStyles = adapterName;
    style.textContent = `${BASE_STYLES}\n${css}`;

    const mountPoint = container.ownerDocument.createElement('div');
    mountPoint.dataset.docSdkMount = adapterName;
    shadowRoot.append(style, mountPoint);
    return mountPoint;
  };

  let mountPoint = build();

  return {
    shadowRoot,
    get mountPoint() {
      return mountPoint;
    },
    addStyles(extraCss: string) {
      if (!extraCss) return;
      const style = container.ownerDocument.createElement('style');
      style.dataset.docSdkStyles = `${adapterName}-engine`;
      style.textContent = extraCss;
      shadowRoot.insertBefore(style, mountPoint);
    },
    reset() {
      mountPoint = build();
    },
  };
}
