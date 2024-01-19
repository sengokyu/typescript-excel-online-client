import * as graph from "@microsoft/microsoft-graph-client";
import * as url from "url";
import { Worksheet } from "./worksheet";

/**
 *
 */
export class Workbook {
  private constructor(
    private readonly client: graph.Client,
    private readonly workbookPath: string
  ) {}

  /**
   *
   * @param client
   * @param workbookPath
   * @returns
   */
  static async init(
    client: graph.Client,
    workbookPath: string
  ): Promise<Workbook> {
    // Check existence.
    await client.api(workbookPath).select(["id"]).get();

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
      `worksheets/${encodeURIComponent(idOrName)}`
    );

    return Worksheet.init(this.client, api);
  }
}
