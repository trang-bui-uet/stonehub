import supplierConfigJson from "@/config/supplier-config.json"
import * as XLSX from "xlsx"
import {
  getColumnValueAsNumber,
  getColumnValueAsString,
  getRangeValueAsString,
  hasInvalidRequiredSizes,
  hasStopKeyword,
  isRowComplete,
  isRowEmpty,
} from "@/lib/workbook-reader-shared"

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
type SupplierConfigRoot = Readonly<{ suppliers: readonly unknown[] }>
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
type ReadSveMixWorkbookResult = Readonly<{
  supplierName: string
  sheetName: string
  sections: readonly ReadSupplierWorkbookSectionResult[]
}>
type ReadSveMixWorkbookParams = Readonly<{
  file: File
  supplierName: string
}>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const VAR_FRESH_TOKENS: readonly string[] = ["LINE / VARIATION", "V", "L"] as const

function isMixConfig(value: unknown): value is SupplierMixConfig {
  if (!value || typeof value !== "object") {
    return false
  }
  const objectValue = value as { name?: unknown; workbook?: unknown; sections?: unknown }
  return typeof objectValue.name === "string" && Boolean(objectValue.workbook) && Array.isArray(objectValue.sections)
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
function resolveFreshVarByText(rawFreshVar: string): "Fresh" | "Var" {
  const normalizedFreshVar = rawFreshVar.trim().toUpperCase()
  return VAR_FRESH_TOKENS.includes(normalizedFreshVar) ? "Var" : "Fresh"
}
function isSectionWithS(rawSlabNo: string): boolean {
  return rawSlabNo.toUpperCase().includes("(S)")
}
function parseRows(
  sheet: XLSX.WorkSheet,
  workbookConfig: SupplierWorkbookConfig,
  columnMapping: SupplierColumnMapping,
): readonly SupplierSlabRow[] {
  const rows: SupplierSlabRow[] = []
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) {
    return rows
  }
  const firstRowNumber = workbookConfig.firstDataRowIndex
  const lastRowNumber = workbookRange.e.r + 1
  for (let rowNumber = firstRowNumber; rowNumber <= lastRowNumber; rowNumber += 1) {
    const slabNoGross = getColumnValueAsString(sheet, columnMapping.slabNoGross, rowNumber)
    const lengthGross = getColumnValueAsNumber(sheet, columnMapping.lengthGross, rowNumber)
    const widthGross = getColumnValueAsNumber(sheet, columnMapping.widthGross, rowNumber)
    const slabNoNet = getColumnValueAsString(sheet, columnMapping.slabNoNet, rowNumber)
    const lengthNet = getColumnValueAsNumber(sheet, columnMapping.lengthNet, rowNumber)
    const widthNet = getColumnValueAsNumber(sheet, columnMapping.widthNet, rowNumber)
    const freshVar = resolveFreshVarByText(
      columnMapping.freshVar ? getColumnValueAsString(sheet, columnMapping.freshVar, rowNumber) : "",
    )
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
    if (hasInvalidRequiredSizes(lengthGross, widthGross, lengthNet, widthNet)) {
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
function parseGeneralInfo(sheet: XLSX.WorkSheet, sectionConfig: SupplierSectionConfig, rowCount: number): SupplierGeneralInfo {
  const numberOfSlabsFromSheet = getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.numberOfSlabs)
  const resolvedNumberOfSlabs = numberOfSlabsFromSheet || rowCount.toString()
  return {
    containerNumber: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.containerNumber),
    materialName: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.materialName),
    typeOfPolish: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.typeOfPolish),
    numberOfSlabs: resolvedNumberOfSlabs,
    loadingDate: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.loadingDate),
    invoiceNumber: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.invoiceNumber),
    invoiceDate: getRangeValueAsString(sheet, sectionConfig.generalInfoMapping.invoiceDate),
  }
}
function createSectionRows(
  allRows: readonly SupplierSlabRow[],
  shouldInclude: (row: SupplierSlabRow) => boolean,
): readonly SupplierSlabRow[] {
  return allRows
    .filter((row: SupplierSlabRow): boolean => shouldInclude(row))
    .map(
      (row: SupplierSlabRow, index: number): SupplierSlabRow => ({
        ...row,
        rowNumber: index + 1,
      }),
    )
}

/**
 * Read workbook data for SVE mix supplier and split by slab number marker "(S)".
 */
export async function readSveMixWorkbookFromFile(params: ReadSveMixWorkbookParams): Promise<ReadSveMixWorkbookResult> {
  const supplierConfig = getSupplierMixConfigByName(params.supplierName)
  const primarySection = supplierConfig.sections[0]
  if (!primarySection) {
    throw new Error(`Supplier "${params.supplierName}" does not define section config.`)
  }
  const fileBuffer = await params.file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array" })
  const configuredSheetName = supplierConfig.workbook.sheetName
  const firstSheetName = workbook.SheetNames[0] ?? ""
  const resolvedSheetName = workbook.Sheets[configuredSheetName] ? configuredSheetName : firstSheetName
  const sheet = workbook.Sheets[resolvedSheetName]
  if (!sheet) {
    throw new Error("Workbook does not contain any readable sheet.")
  }
  const allRows = parseRows(sheet, supplierConfig.workbook, primarySection.columnMapping)
  const rowsWithS = createSectionRows(allRows, (row: SupplierSlabRow): boolean => isSectionWithS(row.slabNoGross))
  const rowsWithoutS = createSectionRows(allRows, (row: SupplierSlabRow): boolean => !isSectionWithS(row.slabNoGross))
  const mix1SectionConfig = supplierConfig.sections.find(
    (sectionConfig: SupplierSectionConfig): boolean => sectionConfig.name === "mix1",
  )
  const mix2SectionConfig = supplierConfig.sections.find(
    (sectionConfig: SupplierSectionConfig): boolean => sectionConfig.name === "mix2",
  )
  if (!mix1SectionConfig || !mix2SectionConfig) {
    throw new Error(`Supplier "${params.supplierName}" must define "mix1" and "mix2".`)
  }
  return {
    supplierName: supplierConfig.name,
    sheetName: resolvedSheetName,
    sections: [
      { name: "mix1", generalInfo: parseGeneralInfo(sheet, mix1SectionConfig, rowsWithS.length), rows: rowsWithS },
      { name: "mix2", generalInfo: parseGeneralInfo(sheet, mix2SectionConfig, rowsWithoutS.length), rows: rowsWithoutS },
    ],
  }
}
