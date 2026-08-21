import { describe, expect, it } from 'vitest';
import { resolveNativeToolbarVisibility } from '../src/core/editor-chrome';

describe('resolveNativeToolbarVisibility', () => {
  it('keeps an adapter native by default and respects its local preference', () => {
    expect(resolveNativeToolbarVisibility(undefined, undefined)).toBe(true);
    expect(resolveNativeToolbarVisibility(undefined, false)).toBe(false);
    expect(resolveNativeToolbarVisibility(undefined, undefined, false)).toBe(false);
  });

  it('lets unified chrome take precedence over every vendor toolbar setting', () => {
    expect(resolveNativeToolbarVisibility({ toolbar: 'unified' }, true)).toBe(false);
    expect(resolveNativeToolbarVisibility({ toolbar: 'unified' }, false)).toBe(false);
  });
});
