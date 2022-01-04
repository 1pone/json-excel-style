// secondary Packaging *XLSX* base on *json-as-xlsx*, style power by *xlsx-js-style*
// https://github.com/LuisEnMarroquin/json-as-xlsx
// [Core API] https://github.com/SheetJS/sheetjs
// [Style API] https://github.com/gitbrent/xlsx-js-style/
import XLSX, { CellObject, CellStyle, WritingOptions } from "xlsx-js-style";

export interface IHeader {
  label: string;
  value: string | ((value: IData) => string | number | boolean | Date | IData);
  option?: {
    type?: "text" | "hyperlink";
    render?: Hyperlink | ((value: IData) => Hyperlink);
    style?: CellStyle;
  };
}

export interface Hyperlink {
  Target: string;
  Tooltip?: string;
}

export interface IData {
  [key: string]: string | number | boolean | Date | IData | any;
}

export interface IJsonSheet {
  sheetName?: string;
  header?: IHeader[] | boolean;
  data: IData[];
}

export interface ISettings {
  extraLength?: number;
  fileName?: string;
  writeOptions?: WritingOptions;
}

function getContentProperty(content: IData, headerItemValue: string) {
  function accessContentProperties(content: IData, properties: string[]) {
    const value = content[properties[0]];
    if (properties.length === 1) return value || "";
    else return accessContentProperties(value as IData, properties.slice(1));
  }

  const properties = headerItemValue.split("."); // prepare to get deep props
  return accessContentProperties(content, properties);
}

function getWorksheetColumnWidths(
  worksheet: XLSX.WorkSheet,
  extraLength?: number
) {
  extraLength === undefined && (extraLength = 1);
  const ref = worksheet["!ref"];
  const columnRange = XLSX.utils.decode_range(ref || "");
  const columnLetters: string[] = [];

  for (let C = columnRange.s.c; C <= columnRange.e.c; C++) {
    const address = XLSX.utils.encode_col(C);
    columnLetters.push(address);
  }
  return columnLetters.map((column) => {
    const columnCells = Object.keys(worksheet).filter(
      (cell) => cell.charAt(0) === column || cell.slice(0, 2) === column
    );
    const maxWidthCell = columnCells.reduce(
      (previousCell, currentCell) =>
        worksheet[previousCell].v.length > worksheet[currentCell].v.length
          ? previousCell
          : currentCell,
      columnCells[0]
    );
    return { width: worksheet[maxWidthCell].v?.length * 2 || 10 + extraLength };
  });
}

function getWorksheet(jsonSheet: IJsonSheet, settings?: ISettings) {
  let jsonSheetRows: IData[] = jsonSheet.data;
  let skipHeader = false;

  let hasHyperlinksCol = false;
  let hasStyle = false;

  if (jsonSheet.header === undefined || jsonSheet.header === true) {
  } else if (jsonSheet.header === false) {
    skipHeader = true;
  } else {
    //  have IHeader[] config,
    if (jsonSheet.data.length) {
      jsonSheetRows = jsonSheet.data.map((dataItem) => {
        const jsonSheetRow: IData = {};
        (jsonSheet.header as IHeader[]).forEach((headerItem) => {
          headerItem?.option?.type === "hyperlink" && (hasHyperlinksCol = true); // mark hyperlink flag
          headerItem?.option?.style && (hasStyle = true); // mark hyperlink flag

          // convert data key with header label
          // get value by custom function
          if (typeof headerItem.value === "function")
            jsonSheetRow[headerItem.label] = headerItem.value(dataItem);
          else
            jsonSheetRow[headerItem.label] = getContentProperty(
              dataItem,
              headerItem.value
            );
        });
        return jsonSheetRow;
      });
    } else {
      // no data but have IHeader[] config
      // only show header row
      const headerRow = {};
      jsonSheet.header.forEach((headerItem) => {
        headerRow[headerItem.label] = "";
      });
      jsonSheetRows.push(headerRow);
    }
  }
  const worksheet = XLSX.utils.json_to_sheet(jsonSheetRows, { skipHeader });
  //  header options, process style and hyperlink
  if (hasHyperlinksCol || hasStyle) getHeaderOptions(jsonSheet, worksheet);

  // setting options
  settings?.extraLength &&
    (worksheet["!cols"] = getWorksheetColumnWidths(
      worksheet,
      settings.extraLength
    ));
  return worksheet;
}

function getHeaderOptions(jsonSheet: IJsonSheet, worksheet: XLSX.WorkSheet) {
  (jsonSheet.header as IHeader[]).forEach((headerItem, index) => {
    const encode_col = XLSX.utils.encode_col(index);

    for (let i = 0; i < jsonSheet.data.length; i++) {
      const cell: CellObject = worksheet[`${encode_col}${i + 2}`];
      //  process style
      if (headerItem?.option?.style) cell.s = headerItem.option.style;
      //  process hyperlink
      if (headerItem?.option?.type === "hyperlink") {
        if (cell.s?.font?.underline === undefined) {
          !Object.prototype.hasOwnProperty.call(cell, "s") && (cell.s = {});
          !Object.prototype.hasOwnProperty.call(cell.s, "font") &&
            (cell.s.font = {});
          !Object.prototype.hasOwnProperty.call(cell.s.font, "underline") &&
            (cell.s.font.underline = true);
        }
        if (typeof headerItem.option?.render === "function") {
          cell.l = headerItem.option?.render(jsonSheet.data[i]);
        } else if (
          typeof headerItem.option?.render === "object" &&
          headerItem.option?.render?.Target
        ) {
          cell.l = headerItem.option?.render;
        }
      }
    }
  });
}

function writeWorkbook(workbook?: XLSX.WorkBook, settings?: ISettings) {
  const filename = settings?.fileName
    ? settings?.fileName.match(/.xlsx$/)
      ? settings?.fileName
      : settings?.fileName + ".xlsx"
    : "export.xlsx";
  const writeOptions = settings?.writeOptions || {};

  return writeOptions.type === "buffer"
    ? XLSX.write(workbook, writeOptions)
    : XLSX.writeFile(workbook, filename, writeOptions);
}

function jsonExcel(
  data: IJsonSheet[],
  settings?: ISettings
): Buffer | undefined {
  if (!data.length) return;
  const _settings = settings || {};
  const workbook = XLSX.utils.book_new();

  data.forEach((actualSheet, actualIndex) => {
    const worksheet = getWorksheet(actualSheet, _settings);
    const worksheetName = actualSheet.sheetName || "Sheet " + (actualIndex + 1);
    XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName);
  });
  return writeWorkbook(workbook, _settings);
}

function jsonSheet(
  data: IData[],
  header?: IHeader[] | boolean,
  fileName?: string,
  settings?: ISettings
) {
  const _settings = settings || { fileName, extraLength: 6 };
  jsonExcel([{ data, header, sheetName: fileName }], _settings);
}

export default jsonExcel;
export { jsonExcel, jsonSheet };
