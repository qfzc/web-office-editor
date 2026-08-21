# 在线文档中台前端 SDK 技术方案

版本：v1.0　日期：2026-08-21

---

## 1. 背景与目标

将 docx / xlsx / pptx 三类文档的浏览器端读取、编辑、导出能力封装为一个统一的前端 SDK，由上层业务系统（React / Vue / 原生 HTML 等任意技术栈）以依赖方式接入，对外暴露一套**格式无关**的调用接口，内部按文档类型分发到不同的开源编辑引擎。

**明确的目标与非目标：**

| 目标 | 说明 |
| --- | --- |
| 统一接口 | `loadFile / exportFile / destroy` 等核心方法在三种文档类型上行为一致 |
| 按需加载 | 打开 docx 时不拉取 xlsx/pptx 相关代码 |
| 宿主无侵入 | 宿主项目不需要安装 React 作为前置依赖 |
| 样式隔离 | SDK 内部样式与 DOM 结构不污染宿主页面 |

| 非目标（本方案不承诺） | 说明 |
| --- | --- |
| 三种文档类型协同编辑走同一套协议 | 底层库的协同机制不同，见第 5 节 |
| pptx 编辑能力对标 docx/xlsx 的成熟度 | pptx-viewer 项目历史尚短，见第 7 节 |
| 完全字节级无损保真 | 依赖各底层库自身的 OOXML 处理能力 |

---

## 2. 技术选型总览

| 文档类型 | 选型组件 | 底层技术 | 许可证 | 成熟度 |
| --- | --- | --- | --- | --- |
| DOCX | `docx-editor`（`@docx-editor.dev/react`） | OOXML 解析/序列化与分页渲染 | Apache-2.0（核心）；追踪修订/评论为付费 Pro 插件 | 中，需持续验证供应链 |
| XLSX | `Fortune-sheet` | React + Canvas/DOM 混合渲染，公式引擎 | MIT | 高，社区活跃 |
| PPTX | `pptx-viewer`（ChristopherVR） | 框架无关核心 + 多框架适配器，Yjs 协同 | 需在引入前核实（仓库历史短） | 低，建议 Beta 阶段使用 |
| SDK 统一内核 | 自研调度器 | TypeScript + Vite | 自研 | — |

> 三个库均需在引入前锁定具体版本号、跑一次 `npm audit` / Socket.dev 供应链扫描，尤其 docx-editor 和 pptx-viewer 的生态成熟度还不高。

---

## 3. 整体架构

```
┌──────────────────────────────────────────────────────────────────┐
│                     上层宿主系统（任意技术栈）                       │
└──────────────────────────────┬─────────────────────────────────────┘
                                │ createDocEditor(container, options)
┌──────────────────────────────▼─────────────────────────────────────┐
│                   @qfzc/frontend-sdk                               │
│  ┌────────────────┐ ┌────────────────┐ ┌─────────────────────────┐ │
│  │ 生命周期管理器    │ │ 类型分发路由器   │ │ 双通道协同抽象层           │ │
│  └────────────────┘ └────────────────┘ └─────────────────────────┘ │
└───────┬──────────────────────┬──────────────────────┬──────────────┘
        │ 动态 import           │                      │
┌───────▼─────────┐  ┌─────────▼─────────┐  ┌──────────▼─────────┐
│   DocxAdapter    │  │    XlsxAdapter     │  │    PptxAdapter      │
│  docx-editor +   │  │ Fortune-sheet +    │  │  pptx-viewer        │
│  私有 React root │  │ 私有 React root    │  │  (框架无关核心)       │
│  + Shadow DOM    │  │ + Shadow DOM       │  │  + Shadow DOM        │
│  + Yjs Provider  │  │ + Op-sync Channel  │  │  + Yjs Provider      │
└──────────────────┘  └────────────────────┘  └──────────────────────┘
```

关键设计：DocxAdapter 与 XlsxAdapter 内部各自挂载**私有、隔离**的 React 运行时（不依赖宿主是否使用 React），PptxAdapter 直接使用其框架无关核心，三者统一通过 Shadow DOM 挂载以隔离样式。

---

## 4. 统一接口设计

```typescript
// src/types/index.ts
export type DocType = 'docx' | 'xlsx' | 'pptx';

export interface DocSDKOptions {
  type: DocType;
  readOnly?: boolean;
  onReady?: () => void;
  onChange?: () => void;
  onError?: (err: Error) => void;
  collab?: CollabConfig; // 可选，见第 5.2 节
}

export interface IDocAdapter {
  loadFile(file: Blob | ArrayBuffer): Promise<void>;
  exportFile(): Promise<Blob>;
  destroy(): void;
  /** 非所有 adapter 都实现协同，调用前需判空 */
  enableCollab?(config: CollabConfig): void;
}

// src/index.ts
export async function createDocEditor(
  container: HTMLElement,
  options: DocSDKOptions
): Promise<IDocAdapter> {
  container.innerHTML = '';
  switch (options.type) {
    case 'docx': {
      const { DocxAdapter } = await import('./adapters/docx-adapter');
      return new DocxAdapter(container, options);
    }
    case 'xlsx': {
      const { XlsxAdapter } = await import('./adapters/xlsx-adapter');
      return new XlsxAdapter(container, options);
    }
    case 'pptx': {
      const { PptxAdapter } = await import('./adapters/pptx-adapter');
      return new PptxAdapter(container, options);
    }
    default:
      throw new Error(`Unsupported doc type: ${options.type}`);
  }
}
```

---

## 5. 分模块适配器设计

### 5.1 DocxAdapter（docx-editor）

```typescript
// src/adapters/docx-adapter.ts
import { createRoot, type Root } from 'react-dom/client';
import { DocxEditor } from '@docx-editor.dev/react';
import '@docx-editor.dev/core/styles/editor.css';
import type { IDocAdapter, DocSDKOptions } from '../types';

export class DocxAdapter implements IDocAdapter {
  private root: Root | null = null;
  private shadow: ShadowRoot;
  private editorRef: any = null;
  private buffer: ArrayBuffer | null = null;

  constructor(private container: HTMLElement, private options: DocSDKOptions) {
    this.shadow = container.attachShadow({ mode: 'open' }); // 样式隔离
    const mountPoint = document.createElement('div');
    this.shadow.appendChild(mountPoint);
    this.root = createRoot(mountPoint); // 与宿主 React（如有）完全隔离的实例
  }

  async loadFile(file: Blob | ArrayBuffer): Promise<void> {
    this.buffer = file instanceof ArrayBuffer ? file : await file.arrayBuffer();
    this.root!.render(
      <DocxEditor
        documentBuffer={this.buffer}
        mode={this.options.readOnly ? 'viewing' : 'editing'}
        ref={(r: any) => (this.editorRef = r)}
        onChange={() => this.options.onChange?.()}
      />
    );
    this.options.onReady?.();
  }

  async exportFile(): Promise<Blob> {
    return this.editorRef.exportDocx(); // 具体 API 以官方文档为准，需接入时核实
  }

  enableCollab(config: CollabConfig): void {
    // docx-editor 走 Yjs，接入统一的 y-websocket provider
  }

  destroy(): void {
    this.root?.unmount();
    this.container.innerHTML = '';
  }
}
```

### 5.2 XlsxAdapter（Fortune-sheet）

Fortune-sheet 没有官方 Vue / vanilla 包，UI 层强绑定 React，因此必须走与 DocxAdapter 相同的"私有 React root"策略：

```typescript
// src/adapters/xlsx-adapter.ts
import { createRoot, type Root } from 'react-dom/client';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import type { IDocAdapter, DocSDKOptions } from '../types';

export class XlsxAdapter implements IDocAdapter {
  private root: Root | null = null;
  private shadow: ShadowRoot;
  private wbData: any[] = [];

  constructor(private container: HTMLElement, private options: DocSDKOptions) {
    this.shadow = container.attachShadow({ mode: 'open' });
    const mountPoint = document.createElement('div');
    mountPoint.style.height = '100%';
    this.shadow.appendChild(mountPoint);
    this.root = createRoot(mountPoint);
  }

  async loadFile(file: Blob | ArrayBuffer): Promise<void> {
    // xlsx 二进制 -> Fortune-sheet 数据结构，用 xlsx 库或自行转换
    this.wbData = await parseXlsxToFortuneData(file);
    this.root!.render(
      <Workbook
        data={this.wbData}
        onOp={(ops) => this.broadcastOp(ops)} // Op 日志，见协同章节
      />
    );
    this.options.onReady?.();
  }

  async exportFile(): Promise<Blob> {
    return exportFortuneDataToXlsx(this.wbData);
  }

  private broadcastOp(ops: unknown[]) {
    // 通过自定义 WebSocket 通道广播 Op，而非 Yjs
    this.options.collab?.opChannel?.send(ops);
  }

  destroy(): void {
    this.root?.unmount();
    this.container.innerHTML = '';
  }
}
```

### 5.3 PptxAdapter（pptx-viewer）

三者中唯一框架无关的一个，无需私有 React root：

```typescript
// src/adapters/pptx-adapter.ts
import { createEngine } from '@pptx-viewer/core'; // 包名以官方文档为准
import type { IDocAdapter, DocSDKOptions } from '../types';

export class PptxAdapter implements IDocAdapter {
  private engine: any;
  private shadow: ShadowRoot;

  constructor(private container: HTMLElement, private options: DocSDKOptions) {
    this.shadow = container.attachShadow({ mode: 'open' });
    const mountPoint = document.createElement('div');
    this.shadow.appendChild(mountPoint);
    this.engine = createEngine(mountPoint, {
      editable: !this.options.readOnly,
    });
  }

  async loadFile(file: Blob | ArrayBuffer): Promise<void> {
    await this.engine.open(file);
    this.options.onReady?.();
  }

  async exportFile(): Promise<Blob> {
    return this.engine.exportPptx();
  }

  enableCollab(config: CollabConfig): void {
    // pptx-viewer 原生走 Yjs，直接复用与 DocxAdapter 相同的 provider
  }

  destroy(): void {
    this.engine.destroy();
    this.container.innerHTML = '';
  }
}
```

---

## 6. 关键工程问题的处理方案

### 6.1 React 实例隔离

DocxAdapter、XlsxAdapter 内部各自 `createRoot`，不与宿主共享 React Context/Hooks。构建配置上二选一：

- **方案 A（推荐，默认）**：SDK 自带私有 React（打进产物），体积增加约 130KB+ gzip，但宿主零耦合，适合"上层系统把我们依赖进去使用"的场景。
- **方案 B**：将 React 声明为 `peerDependencies`，由宿主提供，体积更小但要求宿主 React 版本兼容——仅在明确知道宿主技术栈时启用。

Vite 配置示例（方案 A，默认自包含）：

```typescript
// vite.config.ts
export default defineConfig({
  build: {
    lib: {
      entry: {
        index: 'src/index.ts',
        'adapters/docx-adapter': 'src/adapters/docx-adapter.ts',
        'adapters/xlsx-adapter': 'src/adapters/xlsx-adapter.ts',
        'adapters/pptx-adapter': 'src/adapters/pptx-adapter.ts',
      },
      formats: ['es'],
    },
    rollupOptions: {
      // 不 external react，保证自包含；如需切换到方案 B 再改为 external
      output: { exports: 'named' },
    },
    chunkSizeWarningLimit: 3000,
  },
});
```

### 6.2 样式隔离

统一用 **Shadow DOM** 而非仅靠 CSS 命名空间——命名空间只能保护自己写的代码，保护不了三方库自带的 CSS。所有 adapter 的挂载点都封装在各自的 `attachShadow` 内。

### 6.3 协同编辑：双通道设计

底层机制不统一，SDK 层需要抽象出两条通道，而不是假设全部走同一协议：

| 文档类型 | 协同机制 | SDK 侧接入方式 |
| --- | --- | --- |
| docx / pptx | Yjs CRDT | 统一接入一个 `y-websocket` Provider，`enableCollab()` 直接绑定 |
| xlsx | 自定义 Op 日志（`onOp` 回调） | SDK 内建一个轻量 Op 广播/应用通道，基于普通 WebSocket，收到远端 Op 后用 `applyOp` 写回 Fortune-sheet |

```typescript
export interface CollabConfig {
  docId: string;
  yProvider?: WebsocketProvider; // 供 docx / pptx 使用
  opChannel?: { send(ops: unknown[]): void; onReceive(cb: (ops: unknown[]) => void): void }; // 供 xlsx 使用
}
```

后端需要同时起两套服务：一个 y-websocket 服务（docx/pptx 共用）、一个普通的 Op 广播服务（xlsx 专用），两者可以合并部署但协议层是分开的。

### 6.4 打包与懒加载

保留原方案的动态 `import()` 分发思路，三个 adapter 各自独立 chunk，宿主只在真正打开某类文档时才拉取对应代码与其依赖（含各自私有的 React、DOCX editor core 与 Yjs）。

---

## 7. 分阶段上线计划

| 阶段 | 范围 | 说明 |
| --- | --- | --- |
| Phase 1（GA） | DocxAdapter + XlsxAdapter | 两个底层库相对成熟，可进入生产 |
| Phase 2（Beta） | PptxAdapter | 标记为 `experimental`，仅内部/低风险场景启用，观察至少一个季度的稳定性反馈 |
| Phase 3 | 协同编辑全量开启 | 待 Phase 1/2 单机编辑稳定后再叠加协同复杂度，避免风险叠加 |

---

## 8. 风险登记表

| 风险项 | 等级 | 缓解措施 |
| --- | --- | --- |
| docx-editor 供应链成熟度不明（fork 生态混乱信号） | 中 | 锁版本、`npm audit`/Socket.dev 扫描、灰度观察期 |
| pptx-viewer 项目历史仅数周，功能描述与实际能力可能有落差 | 高 | POC 验证真实复杂 pptx 文件、Beta 阶段限制使用场景 |
| Fortune-sheet 无官方 Vue/vanilla 包，长期被迫捆绑私有 React | 中 | 已在架构中显式承担，写入技术选型文档避免后续误解 |
| 三种文档类型协同机制不统一，运维复杂度上升 | 中 | 双通道抽象层已在第 6.3 节设计，后端需明确两套服务的职责边界 |
| Shadow DOM 下部分第三方库的浮层组件（如颜色选择器）可能定位异常 | 低 | POC 阶段逐一验证浮层类交互，必要时用 `Portal` + 手动定位规避 |

---

## 9. POC 验证清单（正式接入前必须过一遍）

- [ ] 用带样式、图片、表格的真实 docx 文件跑通 加载 → 编辑 → 导出 → Word 打开校验
- [ ] 用带公式、条件格式、合并单元格的真实 xlsx 文件跑通同样闭环
- [ ] 用带动画、母版、多种版式的真实 pptx 文件跑通同样闭环，重点验证"无损写回"是否属实
- [ ] 在 React 宿主、Vue 宿主、纯 HTML 宿主三种环境下各跑一次集成，确认无 React 冲突、无样式泄漏
- [ ] docx/pptx 协同编辑双人实测；xlsx Op 广播双人实测
- [ ] 对 docx-editor、pptx-viewer 各自的 npm 包做一次供应链扫描并记录扫描报告

---

（完）
