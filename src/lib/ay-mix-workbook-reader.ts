import supplierConfigJson from "@/config/supplier-config.json"
import {
  getColumnValueAsNumber,
  getColumnValueAsString,
  getRangeValueAsString,
  hasInvalidRequiredSizes,
  hasStopKeyword,
  isRowComplete,
  isRowEmpty,
} from "@/lib/workbook-reader-shared"
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
type ReadAyMixWorkbookResult = Readonly<{
  supplierName: string
  sheetName: string
  sections: readonly ReadSupplierWorkbookSectionResult[]
}>
type ReadAyMixWorkbookParams = Readonly<{
  file: File
  supplierName: string
}>
type ParsedMixSection = Readonly<{
  materialName: string
  rows: readonly SupplierSlabRow[]
}>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const DEFAULT_MAX_MIX_SECTION = 2
const HEADER_START_COLUMN_INDEX = 0
const HEADER_END_COLUMN_INDEX = 7

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
function createSectionRows(rows: readonly SupplierSlabRow[]): readonly SupplierSlabRow[] {
  return rows.map(
    (row: SupplierSlabRow, index: number): SupplierSlabRow => ({
      ...row,
      rowNumber: index + 1,
    }),
  )
}
function isMergedHeaderRow(sheet: XLSX.WorkSheet, rowNumber: number): boolean {
  const mergedRanges = sheet["!merges"] ?? []
  const hasMergedAtoH = mergedRanges.some(
    (mergedRange: XLSX.Range): boolean =>
      mergedRange.s.r + 1 === rowNumber &&
      mergedRange.e.r + 1 === rowNumber &&
      mergedRange.s.c === HEADER_START_COLUMN_INDEX &&
      mergedRange.e.c === HEADER_END_COLUMN_INDEX,
  )
  if (!hasMergedAtoH) {
    return false
  }
  const headerText = getColumnValueAsString(sheet, "A", rowNumber)
  if (!headerText) {
    return false
  }
  return true
}
function parseRowsByYellowMergedHeader(
  sheet: XLSX.WorkSheet,
  workbookConfig: SupplierWorkbookConfig,
  columnMapping: SupplierColumnMapping,
): readonly ParsedMixSection[] {
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) {
    return [
      { materialName: "", rows: [] },
      { materialName: "", rows: [] },
    ]
  }
  const parsedSections: { materialName: string; rows: SupplierSlabRow[] }[] = [
    { materialName: "", rows: [] },
    { materialName: "", rows: [] },
  ]
  let sectionIndex = -1
  const firstRowNumber = workbookConfig.firstDataRowIndex
  const lastRowNumber = workbookRange.e.r + 1
  for (let rowNumber = firstRowNumber; rowNumber <= lastRowNumber; rowNumber += 1) {
    if (isMergedHeaderRow(sheet, rowNumber)) {
      sectionIndex += 1
      if (sectionIndex >= DEFAULT_MAX_MIX_SECTION) {
        break
      }
      parsedSections[sectionIndex].materialName = getColumnValueAsString(sheet, "A", rowNumber)
      continue
    }
    if (sectionIndex < 0) {
      continue
    }
    const slabNoGross = getColumnValueAsString(sheet, columnMapping.slabNoGross, rowNumber)
    const lengthGross = getColumnValueAsNumber(sheet, columnMapping.lengthGross, rowNumber)
    const widthGross = getColumnValueAsNumber(sheet, columnMapping.widthGross, rowNumber)
    const slabNoNet = getColumnValueAsString(sheet, columnMapping.slabNoNet, rowNumber)
    const lengthNet = getColumnValueAsNumber(sheet, columnMapping.lengthNet, rowNumber)
    const widthNet = getColumnValueAsNumber(sheet, columnMapping.widthNet, rowNumber)
    const rowValues: readonly string[] = [
      slabNoGross,
      lengthGross?.toString() ?? "",
      widthGross?.toString() ?? "",
      slabNoNet,
      lengthNet?.toString() ?? "",
      widthNet?.toString() ?? "",
    ]
    if (hasStopKeyword(rowValues, workbookConfig.stopKeywords)) {
      continue
    }
    if (workbookConfig.stopIfEmpty && parsedSections[sectionIndex].rows.length > 0 && isRowEmpty(rowValues)) {
      continue
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
    parsedSections[sectionIndex].rows.push({
      rowNumber: parsedSections[sectionIndex].rows.length + 1,
      slabNoGross,
      lengthGross,
      widthGross,
      slabNoNet,
      lengthNet,
      widthNet,
      freshVar: "Fresh",
    })
  }
  return parsedSections
}

/**
 * Read AY mix workbook and split sections by yellow merged headers.
 */
export async function readAyMixWorkbookFromFile(params: ReadAyMixWorkbookParams): Promise<ReadAyMixWorkbookResult> {
  const supplierConfig = getSupplierMixConfigByName(params.supplierName)
  const mix1SectionConfig = supplierConfig.sections.find(
    (sectionConfig: SupplierSectionConfig): boolean => sectionConfig.name === "mix1",
  )
  const mix2SectionConfig = supplierConfig.sections.find(
    (sectionConfig: SupplierSectionConfig): boolean => sectionConfig.name === "mix2",
  )
  if (!mix1SectionConfig || !mix2SectionConfig) {
    throw new Error(`Supplier "${params.supplierName}" must define "mix1" and "mix2".`)
  }
  const fileBuffer = await params.file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array", cellStyles: true })
  const configuredSheetName = supplierConfig.workbook.sheetName
  const firstSheetName = workbook.SheetNames[0] ?? ""
  const resolvedSheetName = workbook.Sheets[configuredSheetName] ? configuredSheetName : firstSheetName
  const sheet = workbook.Sheets[resolvedSheetName]
  if (!sheet) {
    throw new Error("Workbook does not contain any readable sheet.")
  }
  const parsedSections = parseRowsByYellowMergedHeader(sheet, supplierConfig.workbook, mix1SectionConfig.columnMapping)
  const mix1Section = parsedSections[0] ?? { materialName: "", rows: [] }
  const mix2Section = parsedSections[1] ?? { materialName: "", rows: [] }
  const mix1Rows = createSectionRows(mix1Section.rows)
  const mix2Rows = createSectionRows(mix2Section.rows)
  const mix1GeneralInfo = parseGeneralInfo(sheet, mix1SectionConfig, mix1Rows.length)
  const mix2GeneralInfo = parseGeneralInfo(sheet, mix2SectionConfig, mix2Rows.length)
  return {
    supplierName: supplierConfig.name,
    sheetName: resolvedSheetName,
    sections: [
      {
        name: "mix1",
        generalInfo: { ...mix1GeneralInfo, materialName: mix1Section.materialName || mix1GeneralInfo.materialName },
        rows: mix1Rows,
      },
      {
        name: "mix2",
        generalInfo: { ...mix2GeneralInfo, materialName: mix2Section.materialName || mix2GeneralInfo.materialName },
        rows: mix2Rows,
      },
    ],
  }
}
