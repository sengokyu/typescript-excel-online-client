import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import "@microsoft/msgraph-sdk-drives";
import { Workbook as GraphWorkbook } from "@microsoft/msgraph-sdk/models/index.js";
import { Worksheet } from "./worksheet.js";

/**
 *
 */
export class Workbook {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly workbook: GraphWorkbook,
  ) {}

  /**
   *
   * @param client
   * @param driveId
   * @param itemId
   * @returns
   */
  static async createInstance(
    client: GraphServiceClient,
    driveId: string,
    itemId: string,
  ): Promise<Workbook> {
    const workbook = await client.drives
      .byDriveId(driveId)
      .items.byDriveItemId(itemId)
      .workbook.get();

    if (!workbook) {
      throw new Error(
        `Workbook not found for driveId='${driveId}', itemId='${itemId}'.`,
      );
    }

    return new Workbook(client, driveId, itemId, workbook);
  }

  public get workbookObject(): GraphWorkbook {
    return this.workbook;
  }

  /**
   *
   * @param idOrName
   * @returns
   */
  public getWorksheet(idOrName: string): Promise<Worksheet> {
    return Worksheet.createInstance(
      this.client,
      this.driveId,
      this.itemId,
      idOrName,
    );
  }
}
