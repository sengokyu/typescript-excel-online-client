import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import "@microsoft/msgraph-sdk-drives";
import { WorkbookTable } from "@microsoft/msgraph-sdk/models/index.js";
import { Cell } from "./cell.js";
import { convertRangeToCells } from "./range-converter.js";

/**
 * Excel テーブルを表すクラス
 */
export class Table {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly worksheetId: string,
    private readonly tableId: string,
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
   * Table インスタンスを生成する
   * @param client Graph API クライアント
   * @param driveId ドライブ ID
   * @param itemId ドライブアイテム ID
   * @param worksheetId ワークシート ID または名前
   * @param idOrName テーブル ID または名前
   */
  static async createInstance(
    client: GraphServiceClient,
    driveId: string,
    itemId: string,
    worksheetId: string,
    idOrName: string,
  ): Promise<Table> {
    const table = await client.drives
      .byDriveId(driveId)
      .items.byDriveItemId(itemId)
      .workbook.worksheets.byWorkbookWorksheetId(worksheetId)
      .tables.byWorkbookTableId(idOrName)
      .get();

    if (!table) {
      throw new Error(`Table '${idOrName}' not found.`);
    }

    const tableId = table.id ?? idOrName;
    return new Table(client, driveId, itemId, worksheetId, tableId, table);
  }

  /**
   * テーブルのデータ部分（ヘッダー行を除く）を Cell の2次元配列で返す
   */
  public async getDataBodyRange(): Promise<Cell[][]> {
    const range = await this.client.drives
      .byDriveId(this.driveId)
      .items.byDriveItemId(this.itemId)
      .workbook.worksheets.byWorkbookWorksheetId(this.worksheetId)
      .tables.byWorkbookTableId(this.tableId)
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
      .workbook.worksheets.byWorkbookWorksheetId(this.worksheetId)
      .tables.byWorkbookTableId(this.tableId)
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
      .workbook.worksheets.byWorkbookWorksheetId(this.worksheetId)
      .tables.byWorkbookTableId(this.tableId)
      .range.get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }
}
