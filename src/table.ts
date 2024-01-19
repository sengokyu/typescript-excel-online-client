import * as graph from "@microsoft/microsoft-graph-client";
import {
  WorkbookTable,
  WorkbookTableRow,
} from "@microsoft/microsoft-graph-types";
import * as url from "url";

const FETCH_SIZE = 100; // 一度に取得する行数

/**
 *
 */
export class Table {
  private constructor(
    private readonly client: graph.Client,
    private readonly tablePath: string,
    public readonly columnNames: string[]
  ) {}

  /**
   *
   * @param client
   * @param tablePath
   * @returns
   */
  static async init(client: graph.Client, tablePath: string): Promise<Table> {
    // Get columns
    const workbookTable = (await client
      .api(tablePath)
      .select(["id"])
      .expand("columns")
      .get()) as WorkbookTable;

    const columnNames = workbookTable.columns!.map((column) => column.name!);

    return new Table(client, tablePath, columnNames);
  }

  /**
   * 行を返します。
   */
  async *getRowGenerator(): AsyncGenerator<Array<null | any>> {
    let skip = 0;

    const fetchNext = async (skip: number): Promise<WorkbookTableRow[]> => {
      const api = url.resolve(this.tablePath, "rows");
      const response = (await this.client
        .api(api)
        .skip(skip)
        .top(FETCH_SIZE)
        .get()) as WorkbookTableRow[];

      return response;
    };

    // FETCH_SIZE未満になるまで繰り返す
    let response: WorkbookTableRow[];
    do {
      response = await fetchNext(skip);

      for (const row of response) {
        yield row.values;
      }

      skip += FETCH_SIZE;
    } while (response.length >= FETCH_SIZE);
  }
}
