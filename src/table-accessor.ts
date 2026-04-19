import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import { WorkbookTable } from "@microsoft/msgraph-sdk/models/index.js";
import { Cell } from "./cell.js";
import { convertRangeToCells } from "./range-converter.js";

/**
 * Excel テーブルを表すクラス
 */
export class TableAccessor {
  constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly table: WorkbookTable,
  ) {}

  /**
   * 生の WorkbookTable オブジェクト
   */
  public get tableObject(): WorkbookTable {
    return this.table;
  }

  /**
   * テーブル名
   */
  public get name(): string {
    return this.table.name ?? "";
  }

  /**
   * テーブルのデータ部分（ヘッダー行を除く）を Cell の2次元配列で返す
   */
  public async getDataBodyRange(): Promise<Cell[][]> {
    const range = await this.client.drives
      .byDriveId(this.driveId)
      .items.byDriveItemId(this.itemId)
      .workbook.tables.byWorkbookTableId(this.table.id!)
      .dataBodyRange.get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }

  /**
   * テーブルのヘッダー行を Cell の2次元配列で返す
   */
  public async getHeaderRowRange(): Promise<Cell[][]> {
    const range = await this.client.drives
      .byDriveId(this.driveId)
      .items.byDriveItemId(this.itemId)
      .workbook.tables.byWorkbookTableId(this.table.id!)
      .headerRowRange.get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }

  /**
   * テーブル全体（ヘッダー行を含む）を Cell の2次元配列で返す
   */
  public async getRange(): Promise<Cell[][]> {
    const range = await this.client.drives
      .byDriveId(this.driveId)
      .items.byDriveItemId(this.itemId)
      .workbook.tables.byWorkbookTableId(this.table.id!)
      .range.get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }
}
