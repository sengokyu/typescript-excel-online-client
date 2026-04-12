import { type WorkbookRange } from "@microsoft/msgraph-sdk/models/index.js";
import {
  isUntypedArray,
  isUntypedNull,
  type UntypedNode,
} from "@microsoft/kiota-abstractions";
import { Cell, type ValueType } from "./cell.js";

function extractCellValue(
  node: UntypedNode,
): string | number | boolean | null {
  const raw = node.getValue();
  if (raw === null || raw === undefined) return null;
  if (
    typeof raw === "string" ||
    typeof raw === "number" ||
    typeof raw === "boolean"
  ) {
    return raw;
  }
  return null;
}

function extractValueType(node: UntypedNode): ValueType {
  if (isUntypedNull(node)) return "Empty";
  const raw = node.getValue();
  const validTypes: ValueType[] = [
    "Unknown",
    "Empty",
    "String",
    "Integer",
    "Double",
    "Boolean",
    "Error",
  ];
  if (typeof raw === "string" && validTypes.includes(raw as ValueType)) {
    return raw as ValueType;
  }
  return "Unknown";
}

function extractRows<T>(
  node: UntypedNode | null | undefined,
  extract: (cell: UntypedNode) => T,
  rowCount: number,
  colCount: number,
  fallback: T,
): T[][] {
  if (!node || !isUntypedArray(node)) {
    return Array.from({ length: rowCount }, () =>
      Array.from({ length: colCount }, () => fallback),
    );
  }
  return node.getValue().map((row) => {
    if (!isUntypedArray(row)) {
      return Array.from({ length: colCount }, () => fallback);
    }
    return row.getValue().map(extract);
  });
}

export function convertRangeToCells(range: WorkbookRange): Cell[][] {
  if (!range.values || !isUntypedArray(range.values)) {
    return [];
  }

  const valueRows = range.values.getValue();
  const rowCount = valueRows.length;
  const colCount =
    rowCount > 0 && isUntypedArray(valueRows[0])
      ? valueRows[0].getValue().length
      : 0;

  const values = extractRows(range.values, extractCellValue, rowCount, colCount, null);
  const valueTypes = extractRows(range.valueTypes, extractValueType, rowCount, colCount, "Unknown" as ValueType);
  const texts = extractRows(range.text, (node) => node.getValue(), rowCount, colCount, null);
  const numberFormats = extractRows(
    range.numberFormat,
    (node) => {
      const raw = node.getValue();
      return typeof raw === "string" ? raw : null;
    },
    rowCount,
    colCount,
    null,
  );

  return values.map((row, r) =>
    row.map(
      (value, c) =>
        new Cell({
          value,
          valueType: valueTypes[r][c],
          text: texts[r][c],
          numberFormat: numberFormats[r][c],
        }),
    ),
  );
}
