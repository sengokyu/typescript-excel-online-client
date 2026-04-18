import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import { WorkbookWorksheet } from "@microsoft/msgraph-sdk/models/index.js";
import { Cell } from "./cell.js";
import { convertRangeToCells } from "./range-converter.js";
import { Table } from "./table.js";
/**
 * Excel ワークシートを表すクラス
 */
export class Worksheet {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly worksheetId: string,
    private readonly worksheet: WorkbookWorksheet,
  ) {}

  /**
   * Worksheet インスタンスを生成する
   * @param client Graph API クライアント
   * @param driveId ドライブ ID
   * @param itemId ドライブアイテム ID
   * @param idOrName ワークシート ID または名前
   */
  static async createInstance(
    client: GraphServiceClient,
    driveId: string,
    itemId: string,
    idOrName: string,
  ): Promise<Worksheet> {
    const worksheet = await client.drives
      .byDriveId(driveId)
      .items.byDriveItemId(itemId)
      .workbook.worksheets.byWorkbookWorksheetId(idOrName)
      .get({
        queryParameters: {
          select: ["id", "name", "position", "visibility", "charts", "tables"],
        },
      });

    if (!worksheet) {
      throw new Error(`Worksheet '${idOrName}' not found.`);
    }

    const id = worksheet.id ?? idOrName;

    return new Worksheet(client, driveId, itemId, id, worksheet);
  }

  /**
   * 生の WorkbookWorksheet オブジェクト
   */
  public get worksheetObject(): WorkbookWorksheet {
    return this.worksheet;
  }

  /**
   * 指定アドレスのセル範囲を Cell の2次元配列で返す
   * @param address セル範囲アドレス（例: "A1:X10"）
   */
  public async getRange(address: string): Promise<Cell[][]> {
    const range = await this.client.drives
      .byDriveId(this.driveId)
      .items.byDriveItemId(this.itemId)
      .workbook.worksheets.byWorkbookWorksheetId(this.worksheetId)
      .rangeWithAddress(address)
      .get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }

  /**
   * 指定したテーブルを取得する
   * @param idOrName テーブル ID または名前
   */
  public getTable(idOrName: string): Promise<Table> {
    return Table.createInstance(
      this.client,
      this.driveId,
      this.itemId,
      this.worksheetId,
      idOrName,
    );
  }
}
