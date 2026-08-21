# 文档中台前端 SDK 第三方项目接入指南

本文面向需要在自身 Web 项目中打开、编辑和保存 DOCX、XLSX、PPTX 文件的接入方。SDK 运行在浏览器中，支持 React、Vue 和原生 HTML；宿主项目不需要安装或暴露 React。

## 1. 接入范围与前提

### 能力范围

| 格式 | 支持的基础流程 | 协同方式 | 上线建议 |
| --- | --- | --- | --- |
| DOCX | 加载、编辑、导出 | 宿主注入 Yjs 绑定 | 可用于生产，先用真实文件验证保真度 |
| XLSX | 加载、编辑、导出 | 宿主提供 Op 消息通道 | 可用于生产，复杂表格先验证 |
| PPTX | 加载、编辑、导出 | SDK 内置 Yjs WebSocket/WebRTC 接入 | Beta，仅建议低风险场景 |

接入页面需要满足以下条件：

- 使用支持 ES2022、`Blob`、`ArrayBuffer`、Shadow DOM 的现代浏览器。
- 构建和打包环境使用 Node.js `>= 20.19`。
- 编辑器只能在浏览器端创建。SSR 框架须在客户端挂载后再动态导入 SDK。
- 编辑器容器必须有明确高度，例如 `height: calc(100vh - 64px)`；没有高度时编辑区域无法正常显示。
- 文件下载地址必须允许浏览器访问。跨域下载需由文件服务配置 CORS；鉴权 Cookie 或 Token 由宿主负责。

SDK 会在传入容器中创建 open Shadow DOM，隔离编辑器样式。该容器应专供 SDK 使用，创建实例时其已有子节点会被清空。

## 2. 获取和安装

包名为 `@doc-platform/frontend-sdk`。交付方应发布到双方约定的 npm 私有源，接入方按以下方式安装：

```bash
npm config set @doc-platform:registry https://npm.example.com/
npm install @doc-platform/frontend-sdk
```

尚未接入私有源时，可由交付方提供 `.tgz` 安装包。交付方在 SDK 仓库执行：

```bash
npm ci
npm run check
npm pack
```

接入方安装交付的包文件：

```bash
npm install /path/to/doc-platform-frontend-sdk-0.1.0.tgz
```

不要把 SDK 源码目录直接复制进业务项目，也不要直接引用其 `src/` 目录；应通过版本化 npm 包使用 `dist/` 产物，才能保留动态分包、类型声明和依赖锁定。

## 3. 最小可用接入

页面提供一个专用容器：

```html
<div id="document-editor" style="height: calc(100vh - 64px)"></div>
```

以下代码完成“从文件服务读取、编辑、导出下载”的闭环：

```ts
import {
  createDocEditor,
  DocSDKError,
  type DocType,
  type IDocAdapter,
} from '@doc-platform/frontend-sdk';

const container = document.querySelector<HTMLElement>('#document-editor');
if (!container) throw new Error('Editor container is missing.');

function getDocType(fileName: string): DocType {
  const extension = fileName.split('.').pop()?.toLowerCase();
  if (extension === 'docx' || extension === 'xlsx' || extension === 'pptx') return extension;
  throw new Error(`Unsupported file type: ${fileName}`);
}

async function openDocument(fileUrl: string, fileName: string): Promise<IDocAdapter> {
  const response = await fetch(fileUrl, { credentials: 'include' });
  if (!response.ok) throw new Error(`Failed to download document: ${response.status}`);

  const editor = await createDocEditor(container, {
    type: getDocType(fileName),
    fileName,
    readOnly: false,
    onReady: () => console.log('Document is ready.'),
    onChange: () => console.log('Document has unsaved changes.'),
    onError: (error) => console.error('Document SDK error:', error),
  });

  try {
    await editor.loadFile(await response.blob());
    return editor;
  } catch (error) {
    editor.destroy();
    throw error;
  }
}

function downloadFile(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

let editor: IDocAdapter | null = null;

try {
  editor = await openDocument('/api/documents/42/content', 'contract.docx');
  const exported = await editor.exportFile();
  downloadFile(exported, 'contract-edited.docx');
} catch (error) {
  if (error instanceof DocSDKError) {
    console.error(error.code, error.message);
  } else {
    console.error(error);
  }
}

// 在路由离开、组件卸载或切换到另一份文件时调用。
editor?.destroy();
```

实际业务中不要在 `onChange` 中直接上传文件。该回调会随编辑操作频繁触发；应仅标记“有未保存修改”，再由保存按钮或经过防抖的自动保存任务调用 `exportFile()`。

## 4. 保存到业务后端

`exportFile()` 返回包含当前编辑结果的 `Blob`。以下示例使用 `FormData` 上传；URL、鉴权头、CSRF 防护和版本号由业务后端约定。

```ts
async function saveDocument(
  editor: IDocAdapter,
  documentId: string,
  fileName: string,
): Promise<void> {
  const blob = await editor.exportFile();
  const form = new FormData();
  form.append('file', blob, fileName);

  const response = await fetch(`/api/documents/${encodeURIComponent(documentId)}/content`, {
    method: 'PUT',
    body: form,
    credentials: 'include',
  });

  if (!response.ok) {
    throw new Error(`Failed to save document: ${response.status}`);
  }
}
```

建议后端在保存接口中校验文档 ID、当前用户权限、文件大小和 MIME 类型，并通过版本号或 ETag 防止两个用户的单机编辑互相覆盖。SDK 不负责文件存储、权限控制或文件病毒扫描。

## 5. 生命周期与界面选项

标准调用顺序如下：

```text
createDocEditor -> loadFile -> exportFile（可选，多次） -> destroy
```

`createDocEditor()` 只创建格式适配器，`loadFile()` 成功完成后才能调用 `exportFile()`。同一实例可以再次调用 `loadFile()` 以替换文件；`destroy()` 可重复调用，但销毁后的实例不可复用。

常用选项：

```ts
const editor = await createDocEditor(container, {
  type: 'xlsx',
  fileName: 'budget.xlsx',
  readOnly: false,
  // 默认 native：保留各文档引擎原生功能区。
  chrome: { toolbar: 'native' },
  xlsx: {
    showFormulaBar: true,
    showSheetTabs: true,
    locale: 'zh-CN',
  },
});
```

当业务系统需要自己提供统一的顶部工具栏时，使用 `chrome: { toolbar: 'unified' }` 隐藏第三方引擎的原生工具栏。此模式只负责统一外观，不会自动补齐字体、公式、动画等引擎特有命令；需要这些编辑能力时应保留 `native` 模式。

各格式的可选项如下：

| 格式 | 选项对象 | 可用配置 |
| --- | --- | --- |
| DOCX | `docx` | `author`、`showToolbar`、`showRuler`、`initialZoom` |
| XLSX | `xlsx` | `showToolbar`、`showFormulaBar`、`showSheetTabs`、`locale` |
| PPTX | `pptx` | `showToolbar`、`showThumbnails`、`showInspector`、`locale`、`autosave` |

当 `chrome.toolbar` 为 `unified` 时，它优先于各格式的 `showToolbar` 配置。

## 6. React、Vue 与 SSR

### React

在 `useEffect` 中创建实例，并在组件卸载时销毁：

```tsx
import { useEffect, useRef } from 'react';
import { createDocEditor, type IDocAdapter } from '@doc-platform/frontend-sdk';

export function DocumentEditor({ file }: { file: File }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<IDocAdapter | null>(null);

  useEffect(() => {
    let cancelled = false;

    async function mount(): Promise<void> {
      if (!containerRef.current) return;
      const type = file.name.split('.').pop()?.toLowerCase() as 'docx' | 'xlsx' | 'pptx';
      const editor = await createDocEditor(containerRef.current, { type, fileName: file.name });
      if (cancelled) {
        editor.destroy();
        return;
      }
      editorRef.current = editor;
      await editor.loadFile(file);
    }

    void mount();
    return () => {
      cancelled = true;
      editorRef.current?.destroy();
      editorRef.current = null;
    };
  }, [file]);

  return <div ref={containerRef} style={{ height: 'calc(100vh - 64px)' }} />;
}
```

### Vue 3

在 `onMounted` 中创建实例，在 `onBeforeUnmount` 中回收：

```vue
<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref } from 'vue';
import { createDocEditor, type IDocAdapter } from '@doc-platform/frontend-sdk';

const container = ref<HTMLElement>();
let editor: IDocAdapter | null = null;

onMounted(async () => {
  const file = await fetch('/api/documents/42/content').then((response) => response.blob());
  editor = await createDocEditor(container.value!, {
    type: 'docx',
    fileName: 'contract.docx',
  });
  await editor.loadFile(file);
});

onBeforeUnmount(() => editor?.destroy());
</script>

<template>
  <div ref="container" class="document-editor" />
</template>

<style scoped>
.document-editor { height: calc(100vh - 64px); }
</style>
```

### Next.js、Nuxt 等 SSR 框架

不要在服务端模块顶层导入 SDK。将编辑器组件声明为客户端组件，或在客户端生命周期中再动态导入：

```ts
const { createDocEditor } = await import('@doc-platform/frontend-sdk');
```

这样可避免服务端没有 `HTMLElement`、DOM 和 Shadow DOM 时的渲染错误。

## 7. 协同编辑（按需接入）

协同不是三种格式共用的同一协议。单机编辑无需提供协同配置。

| 格式 | 宿主需要提供的内容 |
| --- | --- |
| DOCX | `docId`、可选 Yjs provider、把 SDK 提供的 editor facade 绑定到 Yjs 的 `docxBinding` |
| XLSX | `docId` 和 Op Channel；通道按 `docId` 广播、回传并取消订阅操作日志 |
| PPTX | `docId` 及 `yjs` 配置；WebSocket 模式提供 `serverUrl`，WebRTC 模式提供信令服务 |

XLSX Op Channel 示例：

```ts
await editor.enableCollab?.({
  docId: 'document-42',
  opChannel: {
    send: (ops, context) => {
      socket.send(JSON.stringify({ type: 'xlsx-ops', docId: context.docId, ops }));
    },
    onReceive: (receive) => {
      const listener = (event: MessageEvent) => {
        const message = JSON.parse(String(event.data));
        if (message.type === 'xlsx-ops') receive(message.ops, { docId: message.docId });
      };
      socket.addEventListener('message', listener);
      return () => socket.removeEventListener('message', listener);
    },
  },
});
```

PPTX 协同示例：

```ts
await editor.enableCollab?.({
  docId: 'document-42',
  yjs: {
    serverUrl: 'wss://collab.example.com',
    roomId: 'document-42',
    transport: 'websocket',
    userName: 'Alice',
    role: 'collaborator',
  },
});
```

DOCX 的上游编辑器没有稳定的“只传入 WebSocket Provider 即自动绑定”的 API，因此必须实现业务侧 `docxBinding.bind(context)`。协同服务的身份认证、房间授权、离线恢复和最终文件落盘均由接入方后端负责。

## 8. 错误处理与上线检查

SDK 会抛出 `DocSDKError`，可通过 `error.code` 分类处理：

| 错误码 | 接入方处理建议 |
| --- | --- |
| `INVALID_CONTAINER` | 检查容器是否已挂载，且未使用 closed Shadow DOM |
| `UNSUPPORTED_DOC_TYPE` | 拒绝非 docx/xlsx/pptx 文件 |
| `ENGINE_NOT_READY` | 等待 `loadFile()` 成功后再导出 |
| `FILE_LOAD_FAILED` | 检查文件是否损坏、下载是否完整、格式是否匹配 |
| `EXPORT_FAILED` | 提示用户稍后重试并保留原始文件 |
| `COLLAB_CONFIG_INVALID` / `COLLAB_FAILED` | 检查协同配置、鉴权和消息路由 |
| `ADAPTER_DESTROYED` | 创建新实例，不要复用已销毁实例 |

正式上线前请至少完成：

- 用真实业务文件验证“下载 -> 编辑 -> 导出/保存 -> Office 客户端重新打开”的闭环。
- 分别覆盖只读、编辑、路由切换、重复打开不同文件和保存失败重试。
- 验证文件服务的 CORS、鉴权、权限校验、文件大小限制和病毒扫描。
- XLSX 验证公式、合并单元格、行列尺寸和关键格式；宏、图片、复杂条件格式、数据透视表及部分高级样式不保证写回。
- DOCX/PPTX 验证业务依赖的字体、图片、复杂母版、动画和嵌入对象；PPTX 按 Beta 能力评估后再推广。
- 如启用协同，至少进行双用户并发编辑、断线重连、房间隔离和最终文件落盘测试。

## 9. 接入支持所需信息

向 SDK 交付方反馈问题时，请提供 SDK 版本、浏览器版本、文档类型与脱敏样例、错误码/控制台日志、文件获取方式，以及是否开启协同。对于保真度问题，请同时提供原文件和导出文件的可复现样本。
