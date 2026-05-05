import supplierConfigJson from "@/config/supplier-config.json"
import * as XLSX from "xlsx"

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
}>
type SupplierSectionConfig = Readonly<{
  name: string
  columnMapping: SupplierColumnMapping
  generalInfoMapping: SupplierGeneralInfoMapping
}>
type SupplierMixConfig = Readonly<{
  name: string
  workbook: SupplierWorkbookConfig
  sections: readonly SupplierSectionConfig[]
}>
type SupplierConfigRoot = Readonly<{
  suppliers: readonly unknown[]
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
type ReadSupplierWorkbookSectionResult = Readonly<{
  name: string
  generalInfo: SupplierGeneralInfo
  rows: readonly SupplierSlabRow[]
}>
type ReadAnMixWorkbookResult = Readonly<{
  supplierName: string
  sheetName: string
  sections: readonly ReadSupplierWorkbookSectionResult[]
}>
type ReadAnMixWorkbookParams = Readonly<{
  file: File
  supplierName: string
}>
type ParsedSectionsResult = Readonly<{
  sections: readonly ReadSupplierWorkbookSectionResult[]
  totalRows: number
  filledGeneralInfoCount: number
}>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const EXCEL_DATE_MIN_SERIAL = 20_000
const EXCEL_DATE_MAX_SERIAL = 80_000
const EXCEL_DATE_EPOCH_UTC = Date.UTC(1899, 11, 30)

function isMixConfig(value: unknown): value is SupplierMixConfig {
  if (!value || typeof value !== "object") {
    return false
  }
  const objectValue = value as { name?: unknown; workbook?: unknown; sections?: unknown }
  return typeof objectValue.name === "string" && Array.isArray(objectValue.sections) && Boolean(objectValue.workbook)
}

function getSupplierMixConfigByName(supplierName: string): SupplierMixConfig {
  const supplierConfig = SUPPLIER_CONFIG.suppliers.find((supplier: unknown): boolean => {
    if (!supplier || typeof supplier !== "object") {
      return false
    }
    const supplierValue = supplier as { name?: unknown }
    return supplierValue.name === supplierName
  })
  if (!isMixConfig(supplierConfig)) {
    throw new Error(`Supplier "${supplierName}" does not use mix sections.`)
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

function parseSectionRows(
  sheet: XLSX.WorkSheet,
  workbookConfig: SupplierWorkbookConfig,
  sectionConfig: SupplierSectionConfig,
): readonly SupplierSlabRow[] {
  const rows: SupplierSlabRow[] = []
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) {
    return rows
  }
  const lastRowNumber = workbookRange.e.r + 1
  const firstRowNumber = workbookConfig.firstDataRowIndex
  for (let rowNumber = firstRowNumber; rowNumber <= lastRowNumber; rowNumber += 1) {
    const slabNoGross = getColumnValueAsString(sheet, sectionConfig.columnMapping.slabNoGross, rowNumber)
    const lengthGross = getColumnValueAsNumber(sheet, sectionConfig.columnMapping.lengthGross, rowNumber)
    const widthGross = getColumnValueAsNumber(sheet, sectionConfig.columnMapping.widthGross, rowNumber)
    const slabNoNet = getColumnValueAsString(sheet, sectionConfig.columnMapping.slabNoNet, rowNumber)
    const lengthNet = getColumnValueAsNumber(sheet, sectionConfig.columnMapping.lengthNet, rowNumber)
    const widthNet = getColumnValueAsNumber(sheet, sectionConfig.columnMapping.widthNet, rowNumber)
    const rawFreshVar = sectionConfig.columnMapping.freshVar
      ? getColumnValueAsString(sheet, sectionConfig.columnMapping.freshVar, rowNumber)
      : ""
    const normalizedFreshVar = rawFreshVar.trim().toUpperCase()
    const freshVar: "Fresh" | "Var" =
      normalizedFreshVar === "LINE / VARIATION" || normalizedFreshVar === "V" || normalizedFreshVar === "L"
        ? "Var"
        : "Fresh"
    const rowValues: readonly string[] = [
      slabNoGross,
      lengthGross?.toString() ?? "",
      widthGross?.toString() ?? "",
      slabNoNet,
      lengthNet?.toString() ?? "",
      widthNet?.toString() ?? "",
    ]
    if (workbookConfig.stopIfEmpty && rows.length > 0 && isRowEmpty(rowValues)) {
      break
    }
    if (hasStopKeyword(rowValues, workbookConfig.stopKeywords)) {
      break
    }
    if (isRowEmpty(rowValues)) {
      continue
    }
    if (!isRowComplete(rowValues)) {
      continue
    }
    rows.push({
      rowNumber: rows.length + 1,
      slabNoGross,
      lengthGross,
      widthGross,
      slabNoNet,
      lengthNet,
      widthNet,
      freshVar,
    })
  }
  return rows
}

function hasFilledGeneralInfo(generalInfo: SupplierGeneralInfo): boolean {
  return [
    generalInfo.containerNumber,
    generalInfo.materialName,
    generalInfo.typeOfPolish,
    generalInfo.numberOfSlabs,
    generalInfo.loadingDate,
    generalInfo.invoiceNumber,
    generalInfo.invoiceDate,
  ].some((value: string): boolean => value.trim() !== "")
}

function parseSectionGeneralInfo(
  sheet: XLSX.WorkSheet,
  sectionConfig: SupplierSectionConfig,
  rows: readonly SupplierSlabRow[],
): SupplierGeneralInfo {
  const numberOfSlabsFromSheet = getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.numberOfSlabs)
  const resolvedNumberOfSlabs = numberOfSlabsFromSheet || rows.length.toString()
  return {
    containerNumber: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.containerNumber),
    materialName: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.materialName),
    typeOfPolish: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.typeOfPolish),
    numberOfSlabs: resolvedNumberOfSlabs,
    loadingDate: normalizeExcelDateValue(getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.loadingDate)),
    invoiceNumber: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.invoiceNumber),
    invoiceDate: normalizeExcelDateValue(getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.invoiceDate)),
  }
}

function parseSectionsFromSheet(sheet: XLSX.WorkSheet, supplierConfig: SupplierMixConfig): ParsedSectionsResult {
  let totalRows = 0
  let filledGeneralInfoCount = 0
  const sections = supplierConfig.sections.map(
    (sectionConfig: SupplierSectionConfig): ReadSupplierWorkbookSectionResult => {
      const rows = parseSectionRows(sheet, supplierConfig.workbook, sectionConfig)
      const generalInfo = parseSectionGeneralInfo(sheet, sectionConfig, rows)
      totalRows += rows.length
      if (hasFilledGeneralInfo(generalInfo)) {
        filledGeneralInfoCount += 1
      }
      return { name: sectionConfig.name, rows, generalInfo }
    },
  )
  return { sections, totalRows, filledGeneralInfoCount }
}

function resolveBestSheetName(workbook: XLSX.WorkBook, supplierConfig: SupplierMixConfig): string {
  const configuredSheetName = supplierConfig.workbook.sheetName
  const sheetNames = workbook.SheetNames
  if (sheetNames.length === 0) {
    return ""
  }
  const candidateSheetNames = [
    configuredSheetName,
    ...sheetNames.filter((sheetName: string): boolean => sheetName !== configuredSheetName),
  ]
  let bestSheetName = candidateSheetNames[0] ?? ""
  let bestTotalRows = -1
  let bestFilledGeneralInfoCount = -1
  for (const sheetName of candidateSheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      continue
    }
    const parsedResult = parseSectionsFromSheet(sheet, supplierConfig)
    if (parsedResult.totalRows > bestTotalRows) {
      bestSheetName = sheetName
      bestTotalRows = parsedResult.totalRows
      bestFilledGeneralInfoCount = parsedResult.filledGeneralInfoCount
      continue
    }
    if (
      parsedResult.totalRows === bestTotalRows &&
      parsedResult.filledGeneralInfoCount > bestFilledGeneralInfoCount
    ) {
      bestSheetName = sheetName
      bestFilledGeneralInfoCount = parsedResult.filledGeneralInfoCount
    }
  }
  return bestSheetName
}

/**
 * Read workbook data for AN mix supplier by configured sections.
 */
export async function readAnMixWorkbookFromFile(params: ReadAnMixWorkbookParams): Promise<ReadAnMixWorkbookResult> {
  const supplierConfig = getSupplierMixConfigByName(params.supplierName)
  const fileBuffer = await params.file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array" })
  const resolvedSheetName = resolveBestSheetName(workbook, supplierConfig)
  const sheet = workbook.Sheets[resolvedSheetName]
  if (!sheet) {
    throw new Error("Workbook does not contain any readable sheet.")
  }
  const parsedResult = parseSectionsFromSheet(sheet, supplierConfig)
  return {
    supplierName: supplierConfig.name,
    sheetName: resolvedSheetName,
    sections: parsedResult.sections,
  }
}
