import * as graph from "@microsoft/microsoft-graph-client";
import * as url from "url";
import { Table } from "./table";
import { WorkbookRange } from "@microsoft/microsoft-graph-types";

/**
 *
 */
export class Worksheet {
  private constructor(
    private readonly client: graph.Client,
    private readonly worksheetPath: string
  ) {}

  /**
   *
   * @param client
   * @param worksheetPath
   */
  static async init(
    client: graph.Client,
    worksheetPath: string
  ): Promise<Worksheet> {
    // Check existence.
    await client.api(worksheetPath).select(["id"]).get();

    return new Worksheet(client, worksheetPath);
  }

  /**
   *
   * @param idOrName
   * @returns
   */
  async getTable(idOrName: string): Promise<Table> {
    const api = url.resolve(
      this.worksheetPath,
      `tables/${encodeURIComponent(idOrName)}`
    );

    return Table.init(this.client, api);
  }

  /**
   * Get values of the range
   * @param address
   * @returns
   */
  async getRangeValues(address: string): Promise<string[][]> {
    const api = url.resolve(this.worksheetPath, `range(address='${address}')`);
    const response = (await this.client
      .api(api)
      .select("values")
      .get()) as WorkbookRange;

    return response.values as string[][];
  }
}
