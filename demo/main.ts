import { createDocEditor, type DocType, type IDocAdapter } from '../src';
import './styles.css';

const typeSelect = document.querySelector<HTMLSelectElement>('#doc-type')!;
const fileInput = document.querySelector<HTMLInputElement>('#file')!;
const exportButton = document.querySelector<HTMLButtonElement>('#export')!;
const destroyButton = document.querySelector<HTMLButtonElement>('#destroy')!;
const status = document.querySelector<HTMLElement>('#status')!;
const statusGroup = document.querySelector<HTMLElement>('#document-status')!;
const documentTitle = document.querySelector<HTMLElement>('#document-title')!;
const formatBadge = document.querySelector<HTMLElement>('#format-badge')!;
const container = document.querySelector<HTMLElement>('#editor')!;

let editor: IDocAdapter | null = null;
let currentFileName = '';

type StatusTone = 'idle' | 'loading' | 'ready' | 'changed' | 'error';

function setStatus(message: string, tone: StatusTone = 'idle'): void {
  status.textContent = message;
  statusGroup.dataset.tone = tone;
}

function setActive(active: boolean): void {
  exportButton.disabled = !active;
  destroyButton.disabled = !active;
}

function extensionFor(type: DocType): string {
  return `.${type}`;
}

function labelFor(type: DocType): string {
  return type === 'pptx' ? 'PPTX Beta' : type.toUpperCase();
}

function setDocumentIdentity(fileName = '', type?: DocType): void {
  documentTitle.textContent = fileName || '未打开文档';
  formatBadge.textContent = type ? labelFor(type) : '工作台';
}

typeSelect.addEventListener('change', () => {
  fileInput.accept = extensionFor(typeSelect.value as DocType);
  fileInput.value = '';
});

fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) return;

  editor?.destroy();
  setActive(false);
  setStatus(`正在打开 ${file.name}`, 'loading');
  currentFileName = file.name;
  const type = typeSelect.value as DocType;
  setDocumentIdentity(file.name, type);

  try {
    editor = await createDocEditor(container, {
      type,
      fileName: file.name,
      // Keep the vendor command surface available for document-specific edits.
      chrome: { toolbar: 'native' },
      onChange: () => setStatus('有未导出更改', 'changed'),
      onError: (error) => setStatus(error.message, 'error'),
    });
    await editor.loadFile(file);
    setStatus('已打开', 'ready');
    setActive(true);
  } catch (error) {
    editor?.destroy();
    editor = null;
    setStatus(error instanceof Error ? error.message : '文件打开失败', 'error');
  }
});

exportButton.addEventListener('click', async () => {
  if (!editor) return;
  exportButton.disabled = true;
  setStatus('正在导出', 'loading');
  try {
    const blob = await editor.exportFile();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = currentFileName.replace(/(\.[^.]+)?$/, `-edited.${editor.type}`);
    link.click();
    URL.revokeObjectURL(url);
    setStatus('导出完成', 'ready');
  } catch (error) {
    setStatus(error instanceof Error ? error.message : '导出失败', 'error');
  } finally {
    exportButton.disabled = false;
  }
});

destroyButton.addEventListener('click', () => {
  editor?.destroy();
  editor = null;
  fileInput.value = '';
  setActive(false);
  currentFileName = '';
  setDocumentIdentity();
  setStatus('未加载文档');
});
