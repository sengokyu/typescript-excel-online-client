import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import "@microsoft/msgraph-sdk-drives";
import { Cell } from "./cell.js";
import { convertRangeToCells } from "./range-converter.js";

/**
 *
 */
export class Worksheet {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly driveId: string,
    private readonly itemId: string,
    private readonly worksheetId: string,
  ) {}

  /**
   *
   * @param client
   * @param driveId
   * @param itemId
   * @param worksheetId
   */
  static async createInstance(
    client: GraphServiceClient,
    driveId: string,
    itemId: string,
    worksheetId: string,
  ): Promise<Worksheet> {
    const worksheet = await client.drives
      .byDriveId(driveId)
      .items.byDriveItemId(itemId)
      .workbook.worksheets.byWorkbookWorksheetId(worksheetId)
      .get();

    if (!worksheet) {
      throw new Error(`Worksheet '${worksheetId}' not found.`);
    }

    return new Worksheet(client, driveId, itemId, worksheetId);
  }

  /**
   * Get values of the range
   * @param address
   * @returns
   */
  async getRange(address: string): Promise<Cell[][]> {
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
}
