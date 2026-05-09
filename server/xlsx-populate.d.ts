declare module 'xlsx-populate' {
  const XlsxPopulate: {
    fromFileAsync(path: string): Promise<Workbook>;
    fromBlankAsync(): Promise<Workbook>;
  };

  interface Workbook {
    sheet(nameOrIndex: string | number): Sheet;
    sheets(): Sheet[];
    outputAsync(options?: { type?: string }): Promise<Buffer>;
    toFileAsync(path: string): Promise<void>;
  }

  interface Sheet {
    name(): string;
    cell(address: string): Cell;
    cell(row: number, col: number): Cell;
    range(address: string): Range;
  }

  interface Cell {
    value(): unknown;
    value(val: unknown): Cell;
    formula(): string | undefined;
    formula(f: string): Cell;
    style(name: string): unknown;
    style(name: string, val: unknown): Cell;
  }

  interface Range {
    value(): unknown[][];
    value(val: unknown[][]): Range;
  }

  export = XlsxPopulate;
}
