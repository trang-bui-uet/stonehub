import * as XLSX from "xlsx"

const EXCEL_DATE_MIN_SERIAL = 20_000
const EXCEL_DATE_MAX_SERIAL = 80_000
const EXCEL_DATE_EPOCH_UTC = Date.UTC(1899, 11, 30)

export function getCellValueAsString(sheet: XLSX.WorkSheet, cellAddress: string): string {
  const cell = sheet[cellAddress]
  if (!cell || cell.v === undefined || cell.v === null) {
    return ""
  }
  return String(cell.v).trim()
}

export function getColumnValueAsString(sheet: XLSX.WorkSheet, column: string, rowNumber: number): string {
  return getCellValueAsString(sheet, `${column}${rowNumber}`)
}

export function getColumnValueAsNumber(sheet: XLSX.WorkSheet, column: string, rowNumber: number): number | null {
  const rawValue = getColumnValueAsString(sheet, column, rowNumber)
  if (!rawValue) {
    return null
  }
  const normalizedValue = rawValue.replaceAll(",", "")
  const parsedValue = Number(normalizedValue)
  return Number.isNaN(parsedValue) ? null : parsedValue
}

export function getRangeValueAsString(sheet: XLSX.WorkSheet, rangeAddress: string | null): string {
  if (!rangeAddress) {
    return ""
  }
  const decodedRange = XLSX.utils.decode_range(rangeAddress)
  const collectedValues: string[] = []
  for (let rowIndex = decodedRange.s.r; rowIndex <= decodedRange.e.r; rowIndex += 1) {
    for (let columnIndex = decodedRange.s.c; columnIndex <= decodedRange.e.c; columnIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })
      const cellValue = getCellValueAsString(sheet, cellAddress)
      if (cellValue) {
        collectedValues.push(cellValue)
      }
    }
  }
  return collectedValues.join(" ").trim()
}

export function hasStopKeyword(rowValues: readonly string[], stopKeywords: readonly string[]): boolean {
  const normalizedValues = rowValues.map((value: string): string => value.toUpperCase())
  return stopKeywords.some((keyword: string): boolean =>
    normalizedValues.some((value: string): boolean => value.includes(keyword.toUpperCase())),
  )
}

export function isRowEmpty(rowValues: readonly string[]): boolean {
  return rowValues.every((value: string): boolean => value === "")
}

export function isRowComplete(rowValues: readonly string[]): boolean {
  return rowValues.every((value: string): boolean => value !== "")
}

export function hasInvalidRequiredSizes(
  lengthGross: number | null,
  widthGross: number | null,
  lengthNet: number | null,
  widthNet: number | null,
): boolean {
  return [lengthGross, widthGross, lengthNet, widthNet].some(
    (value: number | null): boolean => value === null || value <= 0,
  )
}

function formatDateToDisplayText(dateValue: Date): string {
  const day = String(dateValue.getUTCDate()).padStart(2, "0")
  const month = String(dateValue.getUTCMonth() + 1).padStart(2, "0")
  const year = dateValue.getUTCFullYear()
  return `${day}/${month}/${year}`
}

export function normalizeExcelDateValue(value: string): string {
  const normalizedValue = value.trim()
  if (!normalizedValue) {
    return ""
  }
  const serialValue = Number(normalizedValue)
  if (!Number.isFinite(serialValue) || !Number.isInteger(serialValue)) {
    return normalizedValue
  }
  if (serialValue < EXCEL_DATE_MIN_SERIAL || serialValue > EXCEL_DATE_MAX_SERIAL) {
    return normalizedValue
  }
  const dateValue = new Date(EXCEL_DATE_EPOCH_UTC + serialValue * 24 * 60 * 60 * 1000)
  return formatDateToDisplayText(dateValue)
}
