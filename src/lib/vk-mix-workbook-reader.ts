import supplierConfigJson from "@/config/supplier-config.json"
import * as XLSX from "xlsx"
import {
  getCellValueAsString,
  getColumnValueAsNumber,
  getColumnValueAsString,
  hasInvalidRequiredSizes,
  isRowEmpty,
} from "@/lib/workbook-reader-shared"

type SupplierColumnMapping = Readonly<{ slabNoGross: string; lengthGross: string; widthGross: string; slabNoNet: string; lengthNet: string; widthNet: string; freshVar: string | null }>
type SupplierWorkbookConfig = Readonly<{ sheetName: string; stopKeywords: readonly string[] }>
type SupplierConfig = Readonly<{ name: string; workbook: SupplierWorkbookConfig }>
type SupplierConfigRoot = Readonly<{ suppliers: readonly SupplierConfig[] }>
type SupplierSlabRow = Readonly<{ rowNumber: number; slabNoGross: string; lengthGross: number | null; widthGross: number | null; slabNoNet: string; lengthNet: number | null; widthNet: number | null; freshVar: "Fresh" | "Var" }>
type SupplierGeneralInfo = Readonly<{ containerNumber: string; materialName: string; typeOfPolish: string; numberOfSlabs: string; loadingDate: string; invoiceNumber: string; invoiceDate: string }>
type ReadVkMixWorkbookSectionResult = Readonly<{ name: string; generalInfo: SupplierGeneralInfo; rows: readonly SupplierSlabRow[] }>
type ReadVkMixWorkbookResult = Readonly<{ supplierName: string; sheetName: string; sections: readonly ReadVkMixWorkbookSectionResult[] }>
type ReadVkMixWorkbookParams = Readonly<{ file: File; supplierName: string }>
type HeaderMarker = Readonly<{ slabNo: string; grossSize: string; netSize: string }>
type MaterialNameRange = Readonly<{ startColumn: string; endColumn: string; rowOffsetFromHeader: number }>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const VK_MIX_COLUMN_MAPPING: SupplierColumnMapping = { slabNoGross: "A", lengthGross: "B", widthGross: "D", slabNoNet: "G", lengthNet: "H", widthNet: "J", freshVar: null } as const
const HEADER_MARKER: HeaderMarker = { slabNo: "SLAB NO.", grossSize: "GROSS SIZE IN CM", netSize: "NET SIZE IN CM" } as const
const MATERIAL_NAME_RANGE: MaterialNameRange = { startColumn: "A", endColumn: "K", rowOffsetFromHeader: -3 } as const
const VAR_FRESH_TOKENS: readonly string[] = ["LINE / VARIATION", "LINE-VAR", "V", "L"] as const

function getSupplierConfigByName(supplierName: string): SupplierConfig { const supplierConfig = SUPPLIER_CONFIG.suppliers.find((supplier: SupplierConfig): boolean => supplier.name === supplierName); if (!supplierConfig) throw new Error(`Supplier "${supplierName}" is not configured.`); return supplierConfig }
function normalizeText(value: string): string { return value.trim().toUpperCase() }
function getRangeRowValueAsString(sheet: XLSX.WorkSheet, startColumn: string, endColumn: string, rowNumber: number): string {
  if (rowNumber < 1) return ""
  const startAddress = `${startColumn}${rowNumber}`
  const endAddress = `${endColumn}${rowNumber}`
  const decodedRange = XLSX.utils.decode_range(`${startAddress}:${endAddress}`)
  const values: string[] = []
  for (let columnIndex = decodedRange.s.c; columnIndex <= decodedRange.e.c; columnIndex += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: decodedRange.s.r, c: columnIndex })
    const cellValue = getCellValueAsString(sheet, cellAddress)
    if (cellValue !== "") values.push(cellValue)
  }
  return values.join(" ").trim()
}
function getMaterialNameByHeaderRow(sheet: XLSX.WorkSheet, headerRowNumber: number): string {
  const materialRowNumber = headerRowNumber + MATERIAL_NAME_RANGE.rowOffsetFromHeader
  return getRangeRowValueAsString(sheet, MATERIAL_NAME_RANGE.startColumn, MATERIAL_NAME_RANGE.endColumn, materialRowNumber)
}
function isHeaderRow(sheet: XLSX.WorkSheet, rowNumber: number): boolean {
  const slabNoValue = normalizeText(getColumnValueAsString(sheet, "A", rowNumber))
  const grossSizeValue = normalizeText(getColumnValueAsString(sheet, "B", rowNumber))
  const netSizeValue = normalizeText(getColumnValueAsString(sheet, "H", rowNumber))
  return slabNoValue === HEADER_MARKER.slabNo && grossSizeValue === HEADER_MARKER.grossSize && netSizeValue === HEADER_MARKER.netSize
}
function hasStopKeyword(sheet: XLSX.WorkSheet, rowNumber: number, stopKeywords: readonly string[]): boolean {
  const grossSlabNoValue = normalizeText(getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoGross, rowNumber))
  const netSlabNoValue = normalizeText(getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoNet, rowNumber))
  return stopKeywords.some((keyword: string): boolean => { const normalizedKeyword = normalizeText(keyword); return grossSlabNoValue.includes(normalizedKeyword) || netSlabNoValue.includes(normalizedKeyword) })
}
function isDataRowEmpty(sheet: XLSX.WorkSheet, rowNumber: number): boolean {
  const rowValues = [getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoGross, rowNumber), getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.lengthGross, rowNumber), getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.widthGross, rowNumber), getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoNet, rowNumber), getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.lengthNet, rowNumber), getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.widthNet, rowNumber)]
  return isRowEmpty(rowValues)
}
function resolveFreshVarByText(rawFreshVar: string): "Fresh" | "Var" {
  const normalizedFreshVar = rawFreshVar.trim().toUpperCase()
  return VAR_FRESH_TOKENS.includes(normalizedFreshVar) ? "Var" : "Fresh"
}
function parseSectionRows(sheet: XLSX.WorkSheet, headerRowNumber: number, stopKeywords: readonly string[], lastRowNumber: number): readonly SupplierSlabRow[] {
  const sectionRows: SupplierSlabRow[] = []
  for (let rowNumber = headerRowNumber + 1; rowNumber <= lastRowNumber; rowNumber += 1) {
    if (hasStopKeyword(sheet, rowNumber, stopKeywords)) break
    if (isDataRowEmpty(sheet, rowNumber)) continue
    const slabNoGross = getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoGross, rowNumber)
    const slabNoNet = getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.slabNoNet, rowNumber)
    const lengthGross = getColumnValueAsNumber(sheet, VK_MIX_COLUMN_MAPPING.lengthGross, rowNumber)
    const widthGross = getColumnValueAsNumber(sheet, VK_MIX_COLUMN_MAPPING.widthGross, rowNumber)
    const lengthNet = getColumnValueAsNumber(sheet, VK_MIX_COLUMN_MAPPING.lengthNet, rowNumber)
    const widthNet = getColumnValueAsNumber(sheet, VK_MIX_COLUMN_MAPPING.widthNet, rowNumber)
    const freshVar = resolveFreshVarByText(
      VK_MIX_COLUMN_MAPPING.freshVar ? getColumnValueAsString(sheet, VK_MIX_COLUMN_MAPPING.freshVar, rowNumber) : "",
    )
    if (!slabNoGross || !slabNoNet || hasInvalidRequiredSizes(lengthGross, widthGross, lengthNet, widthNet)) continue
    sectionRows.push({ rowNumber: sectionRows.length + 1, slabNoGross, lengthGross, widthGross, slabNoNet, lengthNet, widthNet, freshVar })
  }
  return sectionRows
}
function createDefaultGeneralInfo(rowCount: number, materialName: string): SupplierGeneralInfo { return { containerNumber: "", materialName, typeOfPolish: "", numberOfSlabs: rowCount.toString(), loadingDate: "", invoiceNumber: "", invoiceDate: "" } }
function parseSectionsFromSheet(sheet: XLSX.WorkSheet, stopKeywords: readonly string[]): readonly ReadVkMixWorkbookSectionResult[] {
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) return []
  const lastRowNumber = workbookRange.e.r + 1
  const sectionHeaders: number[] = []
  for (let rowNumber = 1; rowNumber <= lastRowNumber; rowNumber += 1) { if (isHeaderRow(sheet, rowNumber)) sectionHeaders.push(rowNumber) }
  return sectionHeaders.map((headerRowNumber: number, index: number): ReadVkMixWorkbookSectionResult => {
    const rows = parseSectionRows(sheet, headerRowNumber, stopKeywords, lastRowNumber)
    const materialName = getMaterialNameByHeaderRow(sheet, headerRowNumber)
    return { name: `mix${index + 1}`, generalInfo: createDefaultGeneralInfo(rows.length, materialName), rows }
  })
}

export async function readVkMixWorkbookFromFile(params: ReadVkMixWorkbookParams): Promise<ReadVkMixWorkbookResult> {
  const supplierConfig = getSupplierConfigByName(params.supplierName)
  const fileBuffer = await params.file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array" })
  const configuredSheetName = supplierConfig.workbook.sheetName
  const firstSheetName = workbook.SheetNames[0] ?? ""
  const resolvedSheetName = workbook.Sheets[configuredSheetName] ? configuredSheetName : firstSheetName
  const sheet = workbook.Sheets[resolvedSheetName]
  if (!sheet) throw new Error("Workbook does not contain any readable sheet.")
  const sections = parseSectionsFromSheet(sheet, supplierConfig.workbook.stopKeywords)
  if (sections.length === 0) throw new Error("Không tìm thấy block dữ liệu VK mix theo header marker.")
  return { supplierName: supplierConfig.name, sheetName: resolvedSheetName, sections }
}
