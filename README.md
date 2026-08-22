# 文档中台前端 SDK

统一封装浏览器端 DOCX、XLSX、PPTX 的加载、编辑、导出和销毁能力。宿主可以是 React、Vue 或原生 HTML，不需要提供 React 依赖；三种编辑器按文档类型动态加载，并挂载在 Shadow DOM 中。

源码仓库：[qfzc/web-office-editor](https://github.com/qfzc/web-office-editor)。外部接入请参阅[第三方项目接入指南](./THIRD_PARTY_INTEGRATION_GUIDE.md)。

## 当前实现

| 格式 | 引擎 | 加载/编辑/导出 | 协同 |
| --- | --- | --- | --- |
| DOCX | `@docx-editor.dev/react@2.6.0` | 已接通官方 `DocxEditor/save` API | 提供可注入的 editor facade-Yjs binding |
| XLSX | `@fortune-sheet/react@1.0.4` + SheetJS `0.20.3` | 已接通二进制转换与 `onOp/applyOp` | Op Channel 已实现 |
| PPTX Beta | `pptx-vanilla-viewer@1.22.3` | 已接通真实 `loadFile/save/destroy` API | 内置 Yjs WebSocket/WebRTC |

技术方案中的 PPTX 包名 `@pptx-viewer/core` 并不存在。实现使用同一项目实际发布的 UI 包 `pptx-vanilla-viewer`，其底层引擎为 `pptx-viewer-core`。

## 安装与构建

```bash
npm install
npm run check
```

本地演示：

```bash
npm run dev
```

发布产物位于 `dist/`，入口为 ESM `dist/index.js`，类型声明为 `dist/index.d.ts`。DOCX、XLSX、PPTX 适配器会生成独立动态 chunk。

## 基本接入

面向 React、Vue、原生 HTML 及 SSR 项目的完整落地说明，见[第三方项目接入指南](./THIRD_PARTY_INTEGRATION_GUIDE.md)。

```ts
import { createDocEditor } from '@qfzc/frontend-sdk';

const editor = await createDocEditor(document.querySelector('#editor')!, {
  type: 'docx',
  fileName: 'contract.docx',
  readOnly: false,
  onReady: () => console.log('ready'),
  onChange: () => console.log('changed'),
  onError: (error) => console.error(error),
});

await editor.loadFile(file);
const output = await editor.exportFile();
editor.destroy();
```

宿主必须给容器明确高度。容器可以重复用于不同格式；SDK 会复用自己创建的 open ShadowRoot 并清理旧挂载内容。

### 统一工作台外观

三种第三方编辑器的原生功能区无法共享组件或视觉语言。SDK 提供宿主级开关，允许宿主渲染自己的统一标题栏、文件操作和状态区，同时隐藏引擎原生工具栏：

```ts
const editor = await createDocEditor(container, {
  type: 'docx',
  fileName: 'contract.docx',
  chrome: { toolbar: 'unified' },
});
```

默认值为 `chrome.toolbar: 'native'`，完全保持原有行为。统一模式不会伪造字体、动画、公式等引擎专属操作；需要完整专属编辑能力时，宿主应保留 `native` 模式，或在自身工具栏中接入相应引擎 API。演示页使用统一的文件工作台顶栏，同时默认保留原生编辑工具栏。

## 协同编辑

XLSX 使用宿主提供的 Op 通道。通道应按 `context.docId` 路由消息，并在 `onReceive` 中返回取消订阅函数：

```ts
await editor.enableCollab?.({
  docId: 'sheet-42',
  opChannel: {
    send: (ops, context) => socket.send(JSON.stringify({ ...context, ops })),
    onReceive: (receive) => {
      const listener = (event: MessageEvent) => {
        const message = JSON.parse(String(event.data));
        receive(message.ops, { docId: message.docId });
      };
      socket.addEventListener('message', listener);
      return () => socket.removeEventListener('message', listener);
    },
  },
});
```

PPTX 使用引擎内置的 Yjs：

```ts
await editor.enableCollab?.({
  docId: 'slides-42',
  yjs: {
    serverUrl: 'wss://collab.example.com',
    userName: 'Alice',
    role: 'collaborator',
  },
});
```

DOCX 上游当前没有稳定的“只传 WebsocketProvider 即绑定文档”API，因此 SDK 不伪造绑定行为。宿主需提供 `docxBinding.bind(context)`，在其中把 `context.editor` 的 editor facade 接到 Yjs；返回值是解绑函数。`yProvider` 的连接和断开由 SDK 生命周期管理。

## 错误与生命周期

适配器状态为 `idle / loading / ready / error / destroyed`。在加载完成前导出会抛出 `DocSDKError`，常用 `code` 包括：

- `UNSUPPORTED_DOC_TYPE`
- `ENGINE_NOT_READY`
- `FILE_LOAD_FAILED`
- `EXPORT_FAILED`
- `COLLAB_CONFIG_INVALID`
- `ADAPTER_DESTROYED`

同一个 adapter 可以再次调用 `loadFile()` 替换文件。`destroy()` 幂等，销毁后不能再次加载或导出。

## 保真度边界

- XLSX 转换覆盖值、公式、工作表、合并单元格、行列尺寸和基础数字格式。宏、图片、复杂条件格式、数据透视表和部分高级样式不保证写回。
- DOCX/PPTX 的 OOXML 保真度取决于底层引擎。复杂母版、动画、嵌入对象和第三方字体必须使用真实业务文件做 POC。
- Shadow DOM 能隔离常规样式，但第三方浮层若 Portal 到 `document.body`，仍需逐项做交互验证。
- PPTX 仍按 Beta 管理，不建议直接进入高风险生产流程。

## 供应链状态

锁文件固定了全部版本。`npm audit` 当前结果为 0 漏洞；SheetJS 使用官方 `0.20.3` tarball，Fortune-sheet 的旧 `uuid` 通过兼容的 `11.1.1` override 修复。

DOCX 已迁移至官方 Live Demo 使用的 `@docx-editor.dev/react@2.6.0` 与 `@docx-editor.dev/core@2.6.0`。该版本通过 `DocxEditor` 组件提供完整原生 chrome，SDK 在 Shadow DOM 中挂载它，避免宿主 CSS 干扰。PPTX 使用官方当前 latest `pptx-vanilla-viewer@1.22.3`，并通过其公开静态样式入口加载主题。

DOCX 原生的 `File / Format / Insert / Help` 菜单栏默认关闭，避免与工作台的打开、导出操作重复；正文编辑和下方格式工具栏仍保持可用。

PPTX 原生的标题栏、快捷操作栏和功能标签导航默认关闭；当前功能区的编辑命令、幻灯片缩略图和属性侧栏仍保持可用。

PPTX 的右侧属性栏可由宿主控制初始状态，并在打开后随时切换，无需重新加载文件：

```ts
const editor = await createDocEditor(container, {
  type: 'pptx',
  pptx: { showInspector: false },
});

editor.setInspectorVisible?.(true);
```

## 验证命令

```bash
npm run lint
npm run typecheck
npm run test
npm run build
npm audit
```
