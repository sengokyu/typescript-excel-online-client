import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import "@microsoft/msgraph-sdk-drives";
import * as url from "url";
import { Worksheet } from "./worksheet";

/**
 *
 */
export class Workbook {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly workbookPath: string,
  ) {
    // TODO
  }

  /**
   *
   * @param client
   * @param workbookPath
   * @returns
   */
  static async createInstance(
    client: GraphServiceClient,
    workbookPath: string,
  ): Promise<Workbook> {
    // Check existence.

    // TODO

    return new Workbook(client, workbookPath);
  }

  /**
   *
   * @param idOrName
   * @returns
   */
  public getWorksheet(idOrName: string): Promise<Worksheet> {
    const api = url.resolve(
      this.workbookPath,
      `worksheets/${encodeURIComponent(idOrName)}`,
    );

    return Worksheet.init(this.client, api);
  }
}
