import type { EditorChromeOptions } from '../types';

/**
 * Resolves a vendor toolbar flag without making each adapter own the shared
 * presentation policy. A unified toolbar belongs to the host application.
 */
export function resolveNativeToolbarVisibility(
  chrome: EditorChromeOptions | undefined,
  vendorPreference: boolean | undefined,
  defaultValue = true,
): boolean {
  if (chrome?.toolbar === 'unified') return false;
  return vendorPreference ?? defaultValue;
}
