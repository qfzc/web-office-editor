import { flushSync } from 'react-dom';
import { createRoot, type Root } from 'react-dom/client';
import { Workbook, type WorkbookInstance } from '@fortune-sheet/react';
import type { Op, Sheet } from '@fortune-sheet/core';
import fortuneStyles from '@fortune-sheet/react/dist/index.css?inline';
import { BaseAdapter } from '../core/base-adapter';
import { resolveNativeToolbarVisibility } from '../core/editor-chrome';
import { DocSDKError, asError } from '../core/errors';
import { toArrayBuffer } from '../core/file';
import type { CollabConfig, DocSDKOptions, OpChannelContext } from '../types';
import {
  exportFortuneDataToXlsx,
  parseXlsxToFortuneData,
} from './xlsx/converter';

function isFortuneOp(value: unknown): value is Op {
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as Partial<Op>;
  return (
    Array.isArray(candidate.path) &&
    ['replace', 'remove', 'add', 'insertRowCol', 'deleteRowCol', 'addSheet', 'deleteSheet'].includes(
      String(candidate.op),
    )
  );
}

export class XlsxAdapter extends BaseAdapter {
  private root: Root | null = null;
  private workbook: WorkbookInstance | null = null;
  private workbookData: Sheet[] = [];
  private collabConfig: CollabConfig | null = null;
  private unsubscribe: (() => void) | null = null;
  private applyingRemoteOps = false;

  constructor(container: HTMLElement, options: DocSDKOptions) {
    super('xlsx', container, options, fortuneStyles);
  }

  async loadFile(file: Blob | ArrayBuffer): Promise<void> {
    this.startLoading();
    this.cleanupWorkbook();
    this.mount.reset();

    try {
      this.workbookData = parseXlsxToFortuneData(await toArrayBuffer(file));
      this.root = createRoot(this.mount.mountPoint);

      flushSync(() => {
        this.root?.render(
          <Workbook
            ref={(instance) => {
              this.workbook = instance;
            }}
            data={this.workbookData}
            allowEdit={!this.options.readOnly}
            showToolbar={resolveNativeToolbarVisibility(
              this.options.chrome,
              this.options.xlsx?.showToolbar,
              !this.options.readOnly,
            )}
            showFormulaBar={this.options.xlsx?.showFormulaBar}
            showSheetTabs={this.options.xlsx?.showSheetTabs}
            lang={this.options.xlsx?.locale}
            onChange={(data) => {
              this.workbookData = data;
              this.markChanged();
            }}
            onOp={(ops) => this.broadcastOps(ops)}
          />,
        );
      });

      if (!this.workbook) throw new Error('Fortune-sheet did not expose its workbook instance.');
      this.markReady();
      const collab = this.collabConfig ?? this.options.collab;
      if (collab) await this.enableCollab(collab);
    } catch (cause) {
      this.cleanupWorkbook();
      throw this.fail('FILE_LOAD_FAILED', 'Unable to load the XLSX file.', cause);
    }
  }

  async exportFile(): Promise<Blob> {
    this.assertReady();
    try {
      const data = this.workbook?.getAllSheets() ?? this.workbookData;
      return exportFortuneDataToXlsx(data);
    } catch (cause) {
      throw this.fail('EXPORT_FAILED', 'Unable to export the XLSX file.', cause);
    }
  }

  async enableCollab(config: CollabConfig): Promise<void> {
    this.assertAlive();
    if (!config.docId.trim()) {
      throw new DocSDKError('COLLAB_CONFIG_INVALID', 'XLSX collaboration requires docId.');
    }
    if (!config.opChannel) {
      throw new DocSDKError(
        'COLLAB_CONFIG_INVALID',
        'XLSX collaboration requires an opChannel.',
      );
    }

    this.collabConfig = config;
    if (!this.workbook) return;

    this.unsubscribe?.();
    const maybeUnsubscribe = config.opChannel.onReceive((ops, context) => {
      if (context?.docId && context.docId !== config.docId) return;
      if (!this.workbook || !ops.every(isFortuneOp)) {
        this.reportError(
          new DocSDKError('COLLAB_FAILED', 'Received malformed XLSX collaboration operations.'),
        );
        return;
      }

      this.applyingRemoteOps = true;
      try {
        this.workbook.applyOp([...ops]);
      } catch (cause) {
        this.reportError(
          new DocSDKError('COLLAB_FAILED', 'Unable to apply remote XLSX operations.', {
            cause,
          }),
        );
      } finally {
        this.applyingRemoteOps = false;
      }
    });
    this.unsubscribe = maybeUnsubscribe ?? null;
  }

  destroy(): void {
    if (this.state === 'destroyed') return;
    this.cleanupWorkbook();
    this.markDestroyed();
  }

  private broadcastOps(ops: Op[]): void {
    if (this.applyingRemoteOps || !this.collabConfig?.opChannel) return;
    const context: OpChannelContext = { docId: this.collabConfig.docId };
    Promise.resolve(this.collabConfig.opChannel.send(ops, context)).catch((cause: unknown) => {
      this.reportError(
        new DocSDKError('COLLAB_FAILED', 'Unable to broadcast XLSX operations.', {
          cause: asError(cause),
        }),
      );
    });
  }

  private cleanupWorkbook(): void {
    this.unsubscribe?.();
    this.unsubscribe = null;
    this.root?.unmount();
    this.root = null;
    this.workbook = null;
  }
}
