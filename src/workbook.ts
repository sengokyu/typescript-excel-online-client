import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import "@microsoft/msgraph-sdk-drives";
import {
  Workbook as GraphWorkbook,
  WorkbookTable,
  WorkbookWorksheet,
} from "@microsoft/msgraph-sdk/models/index.js";
import { WorksheetAccessor } from "./worksheet-accessor.js";
import { TableAccessor } from "./table-accessor.js";

/**
 *
 */
export class Workbook {
  private constructor(
    private readonly client: GraphServiceClient,
    public readonly driveId: string,
    public readonly itemId: string,
    private readonly workbook: GraphWorkbook,
  ) {}

  /**
   *
   * @param client
   * @param driveId
   * @param idOrName
   * @returns
   */
  static async createInstance(
    client: GraphServiceClient,
    driveId: string,
    idOrName: string,
  ): Promise<Workbook> {
    const workbook = await client.drives
      .byDriveId(driveId)
      .items.byDriveItemId(idOrName)
      .workbook.get({
        queryParameters: { select: ["id", "worksheets", "tables"] },
      });

    if (!workbook) {
      throw new Error(
        `Workbook not found for driveId='${driveId}', itemId='${idOrName}'.`,
      );
    }

    return new Workbook(client, driveId, idOrName, workbook);
  }

  public get workbookObject(): GraphWorkbook {
    return this.workbook;
  }

  /**
   *
   * @param idOrName
   * @returns
   */
  public worksheets(idOrName: string): WorksheetAccessor {
    const worksheet = this.findWorksheet(idOrName);

    if (!worksheet) {
      throw new Error(`Worksheet not found for '${idOrName}'.`);
    }

    return new WorksheetAccessor(
      this.client,
      this.driveId,
      this.itemId,
      worksheet,
    );
  }

  /**
   * 指定したテーブルを取得する
   * @param idOrName テーブル ID または名前
   */
  public tables(idOrName: string): TableAccessor {
    const table = this.findTable(idOrName);

    if (!table) {
      throw new Error(`Table not found for '${idOrName}'.`);
    }

    return new TableAccessor(this.client, this.driveId, this.itemId, table);
  }

  private findWorksheet(idOrName: string): WorkbookWorksheet | undefined {
    return this.workbook.worksheets!.find(
      (ws) => ws.id === idOrName || ws.name === idOrName,
    );
  }

  private findTable(idOrName: string): WorkbookTable | undefined {
    return this.workbook.tables!.find(
      (ws) => ws.id === idOrName || ws.name === idOrName,
    );
  }
}
