import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import { WorkbookWorksheet } from "@microsoft/msgraph-sdk/models/index.js";
import { Cell } from "./cell.js";
import { convertRangeToCells } from "./range-converter.js";
/**
 * Excel ワークシートを表すクラス
 */
export class WorksheetAccessor {
  constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly worksheet: WorkbookWorksheet,
  ) {}

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
      .workbook.worksheets.byWorkbookWorksheetId(this.worksheet.id!)
      .rangeWithAddress(address)
      .get();

    if (!range) {
      return [];
    }

    return convertRangeToCells(range);
  }
}
