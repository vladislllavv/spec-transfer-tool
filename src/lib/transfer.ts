import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import { imageSize } from "image-size";

const SOURCE_HEADER_MARKER = "Наименование";
const SOURCE_PHOTO_COLUMN = "Фото товара / \n产品照片";
const MAX_HEADER_SCAN_ROWS = 15;

export interface TemplateProfile {
  /** Ключ профиля, используется в API как ?target=<id> */
  id: string;
  /** Название результата для пользователя (без расширения) */
  outputName: string;
  /** Путь к файлу шаблона на диске */
  templatePath: string;
  /** Текст маркера, по которому ищется строка заголовков в шаблоне */
  templateHeaderMarker: string;
  /** Соответствие: заголовок в источнике -> заголовок в шаблоне */
  columnMapping: Record<string, string>;
  /** Заголовок колонки для фото в шаблоне (если есть) */
  photoColumnName?: string;
}

export const SPECIFICATION_PROFILE: TemplateProfile = {
  id: "priority",
  outputName: "Приоритет",
  templatePath: path.join(process.cwd(), "data", "template.xlsx"),
  templateHeaderMarker: "ФОТО",
  photoColumnName: "ФОТО",
  columnMapping: {
    "Фото товара / \n产品照片": "ФОТО",
    "Наименование": "Наименование товара",
    "полный состав материалов / \n材料清单（完整清单）":
      "Технические характеристики товара (материал, тех.данные, конструктивные особенности и т.д.)",
    "Описание применения / \n描述、用途、使用场所": "Область применения",
    "Торговая марка \n品牌名称": "Торговая марка",
    "Артикул \n产品货号": "Артикул (если есть)",
    "Количество \n数量": "Кол-во шт.",
    "Вес брутто \n毛重 ": "Общий вес брутто (кг)",
    "Закупочная цена за штуку, юани \n每件购买价格": "Цена за ед., CNY",
    "Общая закупочная цена, юани\n总购买价格": "Сумма, CNY",
    "метры кубические груза / \n立方米货物": "Объем (м3)",
  },
};

export const BITRIX_PROFILE: TemplateProfile = {
  id: "bitrix",
  outputName: "Битрикс",
  templatePath: path.join(process.cwd(), "data", "bitrix-template.xlsx"),
  templateHeaderMarker: "Наименование",
  photoColumnName: "Фото товара",
  columnMapping: {
    "Наименование": "Наименование",
    "Фото товара / \n产品照片": "Фото товара",
    "Китайское наименование товара / \n中文产品名称": "Китайское наименование товара",
    "полный состав материалов / \n材料清单（完整清单）": "Материал",
    "Описание применения / \n描述、用途、使用场所": "Описание",
    "Ед. измерения    计量单位": "Ед. измерения",
    "Количество \n数量": "Количество",
    "Вес брутто \n毛重 ": "Вес брутто",
    "Торговая марка \n品牌名称": "Торговая марка",
    "Артикул \n产品货号": "Артикул",
    "Общая закупочная цена, юани\n总购买价格": "Общая закупочная цена $",
    "Закупочная цена за штуку, юани \n每件购买价格": "Закупочная цена за штуку $",
  },
};

export const TEMPLATE_PROFILES: TemplateProfile[] = [
  SPECIFICATION_PROFILE,
  BITRIX_PROFILE,
];

function normalize(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "object") {
    const v = value as { richText?: { text: string }[]; text?: string };
    if (Array.isArray(v.richText)) {
      return v.richText.map((t) => t.text).join("");
    }
    if (typeof v.text === "string") return v.text;
  }
  return String(value);
}

function findHeaderRow(
  worksheet: ExcelJS.Worksheet,
  markerText: string
): number {
  for (let r = 1; r <= MAX_HEADER_SCAN_ROWS; r++) {
    const row = worksheet.getRow(r);
    let found = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (normalize(cell.value)?.trim() === markerText) found = true;
    });
    if (found) return r;
  }
  throw new Error(`Не найдена строка заголовков с текстом "${markerText}"`);
}

function buildColumnIndex(
  worksheet: ExcelJS.Worksheet,
  headerRow: number
): Map<string, number> {
  const map = new Map<string, number>();
  const row = worksheet.getRow(headerRow);
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    // Не обрезаем пробелы: некоторые заголовки источника содержат конечные
    // пробелы, и COLUMN_MAPPING ссылается на них дословно.
    const name = normalize(cell.value);
    if (name) map.set(name, colNumber);
  });
  return map;
}

function cellPlainValue(cell: ExcelJS.Cell): unknown {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    const obj = v as {
      richText?: { text: string }[];
      result?: unknown;
      text?: string;
    };
    if (Array.isArray(obj.richText)) {
      return obj.richText.map((t) => t.text).join("");
    }
    if ("result" in obj) return obj.result ?? null;
    if (typeof obj.text === "string") return obj.text;
  }
  return v;
}

function isRowEmpty(row: ExcelJS.Row): boolean {
  let empty = true;
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (cellPlainValue(cell) !== null && cellPlainValue(cell) !== "") {
      empty = false;
    }
  });
  return empty;
}

export interface TransferResult {
  buffer: Buffer;
  rowsTransferred: number;
  photosTransferred: number;
  outputName: string;
}

export async function transferToTemplate(
  sourceBuffer: Buffer,
  profile: TemplateProfile
): Promise<TransferResult> {
  const sourceWb = new ExcelJS.Workbook();
  // exceljs ships its own (older) @types/node Buffer type which can mismatch
  // the project's Buffer type; both are the same Node.js Buffer at runtime.
  await sourceWb.xlsx.load(sourceBuffer as never);
  const sourceWs = sourceWb.worksheets[0];
  if (!sourceWs) throw new Error("В исходном файле нет листов с данными");

  const sourceHeaderRow = findHeaderRow(sourceWs, SOURCE_HEADER_MARKER);
  const sourceCols = buildColumnIndex(sourceWs, sourceHeaderRow);

  const templateBuffer = fs.readFileSync(profile.templatePath);
  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.load(templateBuffer as never);
  const templateWs = templateWb.worksheets[0];
  if (!templateWs) throw new Error("В шаблоне не найдено листов");

  const templateHeaderRow = findHeaderRow(
    templateWs,
    profile.templateHeaderMarker
  );
  const templateCols = buildColumnIndex(templateWs, templateHeaderRow);

  const missingTargets = Object.values(profile.columnMapping).filter(
    (name) => !templateCols.has(name)
  );
  if (missingTargets.length > 0) {
    throw new Error(`Не найдены столбцы шаблона: ${missingTargets.join(", ")}`);
  }

  // Очищаем демонстрационные строки шаблона (ниже заголовка)
  const dataStartRow = templateHeaderRow + 1;
  const lastTemplateRow = Math.max(templateWs.rowCount, dataStartRow + 20);
  for (let r = dataStartRow; r <= lastTemplateRow; r++) {
    const row = templateWs.getRow(r);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
    });
  }

  // Удаляем изображения-примеры, привязанные к строкам данных шаблона
  // (оставляем только то, что закреплено выше области данных, например логотип).
  type MediaEntry = { type: string; range?: { tl?: { row: number } } };
  const wsInternal = templateWs as unknown as { _media: MediaEntry[] };
  wsInternal._media = wsInternal._media.filter((m) => {
    if (m.type !== "image") return true;
    const anchorRow = Math.round(m.range?.tl?.row ?? 0) + 1;
    return anchorRow < dataStartRow;
  });

  const photoTemplateCol = profile.photoColumnName
    ? templateCols.get(profile.photoColumnName)
    : undefined;
  const photoSourceCol = sourceCols.get(SOURCE_PHOTO_COLUMN);

  // Индексируем изображения источника по номеру строки (в колонке "Фото товара")
  const photosByRow = new Map<
    number,
    { buffer: Buffer; extension: "jpeg" | "png" | "gif"; width: number; height: number }[]
  >();
  if (photoSourceCol) {
    const media = sourceWb.model.media;
    for (const img of sourceWs.getImages()) {
      const anchorCol = Math.round(img.range.tl.col) + 1; // 1-based
      if (anchorCol !== photoSourceCol) continue;
      const anchorRow = Math.round(img.range.tl.row) + 1; // 1-based
      const item = media[Number(img.imageId)];
      if (!item) continue;
      const ext = item.extension as "jpeg" | "png" | "gif";
      // Размещаем фото в исходном пиксельном разрешении (как в Excel-файле
      // источника), а не по маленькому отображаемому размеру привязки ячейки.
      let width = 120;
      let height = 120;
      try {
        const dims = imageSize(item.buffer as never);
        if (dims.width && dims.height) {
          width = dims.width;
          height = dims.height;
        }
      } catch {
        // не удалось определить размер — используем значение по умолчанию
      }
      const list = photosByRow.get(anchorRow) ?? [];
      list.push({
        buffer: item.buffer as never,
        extension: ext,
        width,
        height,
      });
      photosByRow.set(anchorRow, list);
    }
  }

  let rowsTransferred = 0;
  let photosTransferred = 0;
  let outRow = dataStartRow;

  const lastSourceRow = Math.max(sourceWs.rowCount, sourceHeaderRow);
  for (let r = sourceHeaderRow + 1; r <= lastSourceRow; r++) {
    const srcRow = sourceWs.getRow(r);
    if (isRowEmpty(srcRow)) continue;

    let wroteAny = false;
    for (const [sourceColName, templateColName] of Object.entries(
      profile.columnMapping
    )) {
      const sourceColIdx = sourceCols.get(sourceColName);
      if (!sourceColIdx) continue;
      const value = cellPlainValue(srcRow.getCell(sourceColIdx));
      if (value === null || value === "") continue;
      const templateColIdx = templateCols.get(templateColName)!;
      templateWs.getCell(outRow, templateColIdx).value = value as ExcelJS.CellValue;
      wroteAny = true;
    }

    const photos = photosByRow.get(r);
    if (photoTemplateCol && photos) {
      for (const photo of photos) {
        const imageId = templateWb.addImage({
          buffer: photo.buffer as never,
          extension: photo.extension,
        });
        templateWs.addImage(imageId, {
          tl: { col: photoTemplateCol - 1, row: outRow - 1 },
          ext: { width: photo.width, height: photo.height },
        });
        photosTransferred++;
        wroteAny = true;
      }
    }

    if (wroteAny) {
      outRow++;
      rowsTransferred++;
    }
  }

  const outBuffer = await templateWb.xlsx.writeBuffer();
  return {
    buffer: outBuffer as never,
    rowsTransferred,
    photosTransferred,
    outputName: profile.outputName,
  };
}

export async function transferSpecification(
  sourceBuffer: Buffer
): Promise<TransferResult> {
  return transferToTemplate(sourceBuffer, SPECIFICATION_PROFILE);
}

export async function transferBitrix(
  sourceBuffer: Buffer
): Promise<TransferResult> {
  return transferToTemplate(sourceBuffer, BITRIX_PROFILE);
}
