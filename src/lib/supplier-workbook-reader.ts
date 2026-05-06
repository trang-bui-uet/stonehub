import * as XLSX from "xlsx"
import supplierConfigJson from "@/config/supplier-config.json"

type SupplierColumnMapping = Readonly<{
  slabNoGross: string
  lengthGross: string
  widthGross: string
  slabNoNet: string
  lengthNet: string
  widthNet: string
  freshVar: string | null
}>

type SupplierGeneralInfoMapping = Readonly<{
  containerNumber: string | null
  materialName: string | null
  typeOfPolish: string | null
  numberOfSlabs: string | null
  loadingDate: string | null
  invoiceNumber: string | null
  invoiceDate: string | null
}>

type SupplierWorkbookConfig = Readonly<{
  sheetName: string
  headerRowIndex: number
  firstDataRowIndex: number
  stopKeywords: readonly string[]
  stopIfEmpty: boolean
  isFreshVarColor?: boolean
}>

type SupplierConfig = Readonly<{
  name: string
  workbook: SupplierWorkbookConfig
  columnMapping: SupplierColumnMapping
  generalInfoMapping: SupplierGeneralInfoMapping
}>

type SupplierConfigRoot = Readonly<{
  suppliers: readonly SupplierConfig[]
}>

type SupplierSlabRow = Readonly<{
  rowNumber: number
  slabNoGross: string
  lengthGross: number | null
  widthGross: number | null
  slabNoNet: string
  lengthNet: number | null
  widthNet: number | null
  freshVar: "Fresh" | "Var"
}>

type SupplierGeneralInfo = Readonly<{
  containerNumber: string
  materialName: string
  typeOfPolish: string
  numberOfSlabs: string
  loadingDate: string
  invoiceNumber: string
  invoiceDate: string
}>

type ReadSupplierWorkbookParams = Readonly<{
  file: File
  supplierName: string
}>

type ReadSupplierWorkbookResult = Readonly<{
  supplierName: string
  sheetName: string
  generalInfo: SupplierGeneralInfo
  rows: readonly SupplierSlabRow[]
}>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const EXCEL_DATE_MIN_SERIAL = 20_000
const EXCEL_DATE_MAX_SERIAL = 80_000
const EXCEL_DATE_EPOCH_UTC = Date.UTC(1899, 11, 30)
const WHITE_RGB_CODES: readonly string[] = ["FFFFFF", "FFFFFFFF"] as const

function getSupplierConfigByName(supplierName: string): SupplierConfig {
  const supplierConfig = SUPPLIER_CONFIG.suppliers.find((supplier: SupplierConfig): boolean => supplier.name === supplierName)
  if (!supplierConfig) {
    throw new Error(`Supplier "${supplierName}" is not configured.`)
  }
  return supplierConfig
}

function getCellValueAsString(sheet: XLSX.WorkSheet, cellAddress: string): string {
  const cell = sheet[cellAddress]
  if (!cell || cell.v === undefined || cell.v === null) {
    return ""
  }
  return String(cell.v).trim()
}

function getColumnValueAsString(sheet: XLSX.WorkSheet, column: string, rowNumber: number): string {
  return getCellValueAsString(sheet, `${column}${rowNumber}`)
}

function getColumnValueAsNumber(sheet: XLSX.WorkSheet, column: string, rowNumber: number): number | null {
  const rawValue = getColumnValueAsString(sheet, column, rowNumber)
  if (!rawValue) {
    return null
  }
  const normalizedValue = rawValue.replaceAll(",", "")
  const parsedValue = Number(normalizedValue)
  return Number.isNaN(parsedValue) ? null : parsedValue
}

function getRangeValueAsString(sheet: XLSX.WorkSheet, rangeAddress: string | null): string {
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

function formatDateToDisplayText(dateValue: Date): string {
  const day = String(dateValue.getUTCDate()).padStart(2, "0")
  const month = String(dateValue.getUTCMonth() + 1).padStart(2, "0")
  const year = dateValue.getUTCFullYear()
  return `${day}/${month}/${year}`
}

function normalizeExcelDateValue(value: string): string {
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

function hasStopKeyword(rowValues: readonly string[], stopKeywords: readonly string[]): boolean {
  const normalizedValues = rowValues.map((value: string): string => value.toUpperCase())
  return stopKeywords.some((keyword: string): boolean =>
    normalizedValues.some((value: string): boolean => value.includes(keyword.toUpperCase())),
  )
}

function isRowEmpty(rowValues: readonly string[]): boolean {
  return rowValues.every((value: string): boolean => value === "")
}

function isRowComplete(rowValues: readonly string[]): boolean {
  return rowValues.every((value: string): boolean => value !== "")
}
function hasInvalidRequiredSizes(
  lengthGross: number | null,
  widthGross: number | null,
  lengthNet: number | null,
  widthNet: number | null,
): boolean {
  return [lengthGross, widthGross, lengthNet, widthNet].some(
    (value: number | null): boolean => value === null || value <= 0,
  )
}
function getCellStyleColorCode(sheet: XLSX.WorkSheet, cellAddress: string): string {
  const cell = sheet[cellAddress] as (XLSX.CellObject & { s?: { fgColor?: { rgb?: string }; bgColor?: { rgb?: string } } }) | undefined
  const foregroundColor = cell?.s?.fgColor?.rgb?.trim().toUpperCase() ?? ""
  if (foregroundColor) {
    return foregroundColor
  }
  return cell?.s?.bgColor?.rgb?.trim().toUpperCase() ?? ""
}
function isWhiteColorCode(colorCode: string): boolean {
  if (!colorCode) {
    return true
  }
  if (WHITE_RGB_CODES.includes(colorCode)) {
    return true
  }
  return WHITE_RGB_CODES.includes(colorCode.startsWith("FF") ? colorCode.slice(2) : colorCode)
}
function resolveFreshVarByCellBackground(sheet: XLSX.WorkSheet, supplierConfig: SupplierConfig, rowNumber: number): "Fresh" | "Var" {
  if (!supplierConfig.workbook.isFreshVarColor || !supplierConfig.columnMapping.freshVar) {
    return "Fresh"
  }
  const cellAddress = `${supplierConfig.columnMapping.freshVar}${rowNumber}`
  const colorCode = getCellStyleColorCode(sheet, cellAddress)
  return isWhiteColorCode(colorCode) ? "Fresh" : "Var"
}

function parseRows(sheet: XLSX.WorkSheet, supplierConfig: SupplierConfig): readonly SupplierSlabRow[] {
  const rows: SupplierSlabRow[] = []
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) {
    return rows
  }
  const lastRowNumber = workbookRange.e.r + 1
  const firstRowNumber = supplierConfig.workbook.firstDataRowIndex
  for (let rowNumber = firstRowNumber; rowNumber <= lastRowNumber; rowNumber += 1) {
    const slabNoGross = getColumnValueAsString(sheet, supplierConfig.columnMapping.slabNoGross, rowNumber)
    const lengthGross = getColumnValueAsNumber(sheet, supplierConfig.columnMapping.lengthGross, rowNumber)
    const widthGross = getColumnValueAsNumber(sheet, supplierConfig.columnMapping.widthGross, rowNumber)
    const slabNoNet = getColumnValueAsString(sheet, supplierConfig.columnMapping.slabNoNet, rowNumber)
    const lengthNet = getColumnValueAsNumber(sheet, supplierConfig.columnMapping.lengthNet, rowNumber)
    const widthNet = getColumnValueAsNumber(sheet, supplierConfig.columnMapping.widthNet, rowNumber)
    const rawFreshVar = supplierConfig.columnMapping.freshVar
      ? getColumnValueAsString(sheet, supplierConfig.columnMapping.freshVar, rowNumber)
      : ""
    const normalizedFreshVar = rawFreshVar.trim().toUpperCase()
    const textFreshVar: "Fresh" | "Var" =
      normalizedFreshVar === "LINE / VARIATION" || normalizedFreshVar === "V" || normalizedFreshVar === "L"
        ? "Var"
        : "Fresh"
    const colorFreshVar = resolveFreshVarByCellBackground(sheet, supplierConfig, rowNumber)
    const freshVar: "Fresh" | "Var" =
      supplierConfig.workbook.isFreshVarColor && (textFreshVar === "Var" || colorFreshVar === "Var")
        ? "Var"
        : textFreshVar
    const rowValues: readonly string[] = [
      slabNoGross,
      lengthGross?.toString() ?? "",
      widthGross?.toString() ?? "",
      slabNoNet,
      lengthNet?.toString() ?? "",
      widthNet?.toString() ?? "",
    ]
    if (supplierConfig.workbook.stopIfEmpty && isRowEmpty(rowValues)) {
      break
    }
    if (hasStopKeyword(rowValues, supplierConfig.workbook.stopKeywords)) {
      break
    }
    if (isRowEmpty(rowValues)) {
      continue
    }
    if (hasInvalidRequiredSizes(lengthGross, widthGross, lengthNet, widthNet)) {
      continue
    }
    if (!isRowComplete(rowValues)) {
      continue
    }
    rows.push({ rowNumber, slabNoGross, lengthGross, widthGross, slabNoNet, lengthNet, widthNet, freshVar })
  }
  return rows
}

function parseGeneralInfo(
  sheet: XLSX.WorkSheet,
  supplierConfig: SupplierConfig,
  rows: readonly SupplierSlabRow[],
): SupplierGeneralInfo {
  const numberOfSlabsFromSheet = getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.numberOfSlabs)
  const resolvedNumberOfSlabs = numberOfSlabsFromSheet || rows.length.toString()
  return {
    containerNumber: getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.containerNumber),
    materialName: getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.materialName),
    typeOfPolish: getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.typeOfPolish),
    numberOfSlabs: resolvedNumberOfSlabs,
    loadingDate: normalizeExcelDateValue(getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.loadingDate)),
    invoiceNumber: getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.invoiceNumber),
    invoiceDate: normalizeExcelDateValue(getRangeValueAsString(sheet, supplierConfig.generalInfoMapping.invoiceDate)),
  }
}

/**
 * Read supplier workbook and map data by configured layout.
 */
export async function readSupplierWorkbookFromFile(params: ReadSupplierWorkbookParams): Promise<ReadSupplierWorkbookResult> {
  const supplierConfig = getSupplierConfigByName(params.supplierName)
  const fileBuffer = await params.file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array", cellStyles: true })
  const configuredSheetName = supplierConfig.workbook.sheetName
  const firstSheetName = workbook.SheetNames[0] ?? ""
  const resolvedSheetName = workbook.Sheets[configuredSheetName] ? configuredSheetName : firstSheetName
  const sheet = workbook.Sheets[resolvedSheetName]
  if (!sheet) {
    throw new Error(`Workbook does not contain any readable sheet.`)
  }
  const rows = parseRows(sheet, supplierConfig)
  return {
    supplierName: supplierConfig.name,
    sheetName: resolvedSheetName,
    generalInfo: parseGeneralInfo(sheet, supplierConfig, rows),
    rows,
  }
}
