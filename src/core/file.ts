export async function toArrayBuffer(file: Blob | ArrayBuffer): Promise<ArrayBuffer> {
  return file instanceof ArrayBuffer ? file : file.arrayBuffer();
}

export function toBlobPart(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}
