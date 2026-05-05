import templateWorkbookMixFileUrl from "@/assets/template-tp-mix.xlsx?url"
import templateMixConfigJson from "@/config/template-thien-phuc-config-mix.json"
import ExcelJS from "exceljs"
import * as XLSX from "xlsx"

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
type MixWorkbookSectionData = Readonly<{
  name: string
  generalInfo: SupplierGeneralInfo
  rows: readonly SupplierSlabRow[]
}>
type MixWorkbookData = Readonly<{
  supplierName: string
  sheetName: string
  sections: readonly MixWorkbookSectionData[]
}>
type TemplateSectionGeneralInfoMapping = Readonly<{
  materialName: string
}>
type TemplateSectionColumnMapping = Readonly<{
  slabNoGross: string
  lengthGross: string
  widthGross: string
  grossSquareMeter: string
  slabNoNet: string
  lengthNet: string
  widthNet: string
  netSquareMeter: string
  freshVar: string
}>
type TemplateMixSectionConfig = Readonly<{
  name: string
  dataStartRowIndex: number
  generalInfoMapping: TemplateSectionGeneralInfoMapping
  columnMapping: TemplateSectionColumnMapping
}>
type TemplateMixConfig = Readonly<{
  sheetIndex: number
  sections: readonly TemplateMixSectionConfig[]
}>

const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const GROSS_MEASUREMENT_TEXT = "GROSS MEASUREMENT"
const NET_MEASUREMENT_TEXT = "NET MEASUREMENT"
const VAR_LABEL_TEXT = "V"
const DEFAULT_MATERIAL_NAME_TEXT = "Material Name"
const TEMPLATE_MIX_CONFIG: TemplateMixConfig = templateMixConfigJson as TemplateMixConfig

function calculateSquareMeter(length: number | null, width: number | null): number {
  if (length === null || width === null) {
    return 0
  }
  return (length * width) / CENTIMETER_SQUARE_TO_METER_SQUARE
}

function roundNumberToTwoDigits(value: number): number {
  return Math.round(value * 100) / 100
}

function normalizeTextValue(value: string): string {
  return value.trim()
}

function setCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, value: string | number): void {
  worksheet.getCell(cellAddress).value = value
}

function setCellFormula(worksheet: ExcelJS.Worksheet, cellAddress: string, formula: string, result: number): void {
  worksheet.getCell(cellAddress).value = { formula, result }
}

function getRangeStartCellAddress(rangeAddress: string): string {
  const decodedRange = XLSX.utils.decode_range(rangeAddress)
  return XLSX.utils.encode_cell({ r: decodedRange.s.r, c: decodedRange.s.c })
}

function shiftCellOrRangeAddress(address: string, rowOffset: number): string {
  if (rowOffset === 0) {
    return address
  }
  const decodedRange = XLSX.utils.decode_range(address)
  const shiftedRange = {
    s: { c: decodedRange.s.c, r: decodedRange.s.r + rowOffset },
    e: { c: decodedRange.e.c, r: decodedRange.e.r + rowOffset },
  }
  return XLSX.utils.encode_range(shiftedRange)
}

function recreateMergedRange(worksheet: ExcelJS.Worksheet, rangeAddress: string): void {
  try {
    worksheet.unMergeCells(rangeAddress)
  } catch {
    // Keep silent when range is not merged.
  }
  try {
    worksheet.mergeCells(rangeAddress)
  } catch {
    // Keep silent when range cannot be merged.
  }
}

function recreateSectionHeaderMergedRanges(
  worksheet: ExcelJS.Worksheet,
  sectionConfig: TemplateMixSectionConfig,
  dataStartRowNumber: number,
): void {
  const titleRange = `${sectionConfig.columnMapping.slabNoGross}${dataStartRowNumber - 3}:${sectionConfig.columnMapping.netSquareMeter}${dataStartRowNumber - 3}`
  const grossHeaderRange = `${sectionConfig.columnMapping.slabNoGross}${dataStartRowNumber - 2}:${sectionConfig.columnMapping.grossSquareMeter}${dataStartRowNumber - 2}`
  const netHeaderRange = `${sectionConfig.columnMapping.slabNoNet}${dataStartRowNumber - 2}:${sectionConfig.columnMapping.netSquareMeter}${dataStartRowNumber - 2}`
  recreateMergedRange(worksheet, titleRange)
  recreateMergedRange(worksheet, grossHeaderRange)
  recreateMergedRange(worksheet, netHeaderRange)
}

function setFooterLabel(worksheet: ExcelJS.Worksheet, cellAddress: string, value: string): void {
  const cell = worksheet.getCell(cellAddress)
  if (typeof cell.value === "string" && cell.value.trim() !== "") {
    return
  }
  setCellValue(worksheet, cellAddress, value)
}

function prepareTemplateSectionDataRows(
  worksheet: ExcelJS.Worksheet,
  dataStartRowNumber: number,
  dataRowCount: number,
): number {
  const totalRowNumberWhenNoData = dataStartRowNumber + 1
  if (dataRowCount <= 0) {
    return totalRowNumberWhenNoData
  }
  const extraRowCount = dataRowCount - 1
  for (let index = 0; index < extraRowCount; index += 1) {
    worksheet.duplicateRow(dataStartRowNumber, 1, true)
  }
  return dataStartRowNumber + dataRowCount
}

function setSectionGeneralInfo(
  worksheet: ExcelJS.Worksheet,
  sectionConfig: TemplateMixSectionConfig,
  sectionData: MixWorkbookSectionData,
  rowOffset: number,
): void {
  const resolvedMaterialName = normalizeTextValue(sectionData.generalInfo.materialName) || DEFAULT_MATERIAL_NAME_TEXT
  const materialNameAddress = shiftCellOrRangeAddress(sectionConfig.generalInfoMapping.materialName, rowOffset)
  recreateMergedRange(worksheet, materialNameAddress)
  setCellValue(worksheet, getRangeStartCellAddress(materialNameAddress), resolvedMaterialName)
}

function setSectionRowData(
  worksheet: ExcelJS.Worksheet,
  sectionConfig: TemplateMixSectionConfig,
  sectionData: MixWorkbookSectionData,
  dataStartRowNumber: number,
): Readonly<{ grossTotalSquareMeter: number; netTotalSquareMeter: number }> {
  let grossTotalSquareMeter = 0
  let netTotalSquareMeter = 0
  for (const [index, row] of sectionData.rows.entries()) {
    const currentRowNumber = dataStartRowNumber + index
    const grossSquareMeter = calculateSquareMeter(row.lengthGross, row.widthGross)
    const netSquareMeter = calculateSquareMeter(row.lengthNet, row.widthNet)
    grossTotalSquareMeter += grossSquareMeter
    netTotalSquareMeter += netSquareMeter
    setCellValue(worksheet, `${sectionConfig.columnMapping.slabNoGross}${currentRowNumber}`, row.slabNoGross)
    setCellValue(worksheet, `${sectionConfig.columnMapping.lengthGross}${currentRowNumber}`, row.lengthGross ?? "")
    setCellValue(worksheet, `${sectionConfig.columnMapping.widthGross}${currentRowNumber}`, row.widthGross ?? "")
    setCellFormula(
      worksheet,
      `${sectionConfig.columnMapping.grossSquareMeter}${currentRowNumber}`,
      `=${sectionConfig.columnMapping.lengthGross}${currentRowNumber}*${sectionConfig.columnMapping.widthGross}${currentRowNumber}/${CENTIMETER_SQUARE_TO_METER_SQUARE}`,
      roundNumberToTwoDigits(grossSquareMeter),
    )
    setCellValue(worksheet, `${sectionConfig.columnMapping.slabNoNet}${currentRowNumber}`, row.slabNoNet)
    setCellValue(worksheet, `${sectionConfig.columnMapping.lengthNet}${currentRowNumber}`, row.lengthNet ?? "")
    setCellValue(worksheet, `${sectionConfig.columnMapping.widthNet}${currentRowNumber}`, row.widthNet ?? "")
    setCellFormula(
      worksheet,
      `${sectionConfig.columnMapping.netSquareMeter}${currentRowNumber}`,
      `=${sectionConfig.columnMapping.lengthNet}${currentRowNumber}*${sectionConfig.columnMapping.widthNet}${currentRowNumber}/${CENTIMETER_SQUARE_TO_METER_SQUARE}`,
      roundNumberToTwoDigits(netSquareMeter),
    )
    setCellValue(
      worksheet,
      `${sectionConfig.columnMapping.freshVar}${currentRowNumber}`,
      row.freshVar === "Var" ? VAR_LABEL_TEXT : "",
    )
  }
  return { grossTotalSquareMeter, netTotalSquareMeter }
}

function setSectionTotalRow(
  worksheet: ExcelJS.Worksheet,
  sectionConfig: TemplateMixSectionConfig,
  totalRowNumber: number,
  grossTotalSquareMeter: number,
  netTotalSquareMeter: number,
): void {
  const grossLabelRange = `${sectionConfig.columnMapping.slabNoGross}${totalRowNumber}:${sectionConfig.columnMapping.widthGross}${totalRowNumber}`
  const netLabelRange = `${sectionConfig.columnMapping.slabNoNet}${totalRowNumber}:${sectionConfig.columnMapping.widthNet}${totalRowNumber}`
  recreateMergedRange(worksheet, grossLabelRange)
  recreateMergedRange(worksheet, netLabelRange)
  setFooterLabel(worksheet, `${sectionConfig.columnMapping.slabNoGross}${totalRowNumber}`, GROSS_MEASUREMENT_TEXT)
  setFooterLabel(worksheet, `${sectionConfig.columnMapping.slabNoNet}${totalRowNumber}`, NET_MEASUREMENT_TEXT)
  setCellValue(
    worksheet,
    `${sectionConfig.columnMapping.grossSquareMeter}${totalRowNumber}`,
    roundNumberToTwoDigits(grossTotalSquareMeter),
  )
  setCellValue(
    worksheet,
    `${sectionConfig.columnMapping.netSquareMeter}${totalRowNumber}`,
    roundNumberToTwoDigits(netTotalSquareMeter),
  )
}

function formatCurrentDayMonthText(): string {
  const currentDate = new Date()
  const day = String(currentDate.getDate()).padStart(2, "0")
  const month = String(currentDate.getMonth() + 1).padStart(2, "0")
  return `${day}-${month}`
}

function buildOutputFileName(workbookData: MixWorkbookData): string {
  const firstSection = workbookData.sections[0]
  const resolvedMaterialName = firstSection
    ? normalizeTextValue(firstSection.generalInfo.materialName) || DEFAULT_MATERIAL_NAME_TEXT
    : DEFAULT_MATERIAL_NAME_TEXT
  const dayMonthText = formatCurrentDayMonthText()
  return `TP MIX - ${resolvedMaterialName} ${dayMonthText}.xlsx`
}

/**
 * Export list file for mix workbook data using shared mix template.
 */
export async function exportTemplateMixFile(workbookData: MixWorkbookData): Promise<void> {
  const templateResponse = await fetch(templateWorkbookMixFileUrl)
  if (!templateResponse.ok) {
    throw new Error("Không thể tải file template mix.")
  }
  const templateArrayBuffer = await templateResponse.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateArrayBuffer)
  const worksheet = workbook.getWorksheet(TEMPLATE_MIX_CONFIG.sheetIndex + 1) ?? workbook.worksheets[0]
  if (!worksheet) {
    throw new Error("Template mix không có sheet hợp lệ.")
  }
  const sectionByName = new Map(
    workbookData.sections.map(
      (sectionData: MixWorkbookSectionData): readonly [string, MixWorkbookSectionData] => [sectionData.name, sectionData],
    ),
  )
  let accumulatedRowOffset = 0
  for (const sectionConfig of TEMPLATE_MIX_CONFIG.sections) {
    const sectionData = sectionByName.get(sectionConfig.name)
    if (!sectionData) {
      continue
    }
    const resolvedDataStartRowNumber = sectionConfig.dataStartRowIndex + accumulatedRowOffset
    recreateSectionHeaderMergedRanges(worksheet, sectionConfig, resolvedDataStartRowNumber)
    const totalRowNumber = prepareTemplateSectionDataRows(worksheet, resolvedDataStartRowNumber, sectionData.rows.length)
    setSectionGeneralInfo(worksheet, sectionConfig, sectionData, accumulatedRowOffset)
    const { grossTotalSquareMeter, netTotalSquareMeter } = setSectionRowData(
      worksheet,
      sectionConfig,
      sectionData,
      resolvedDataStartRowNumber,
    )
    setSectionTotalRow(worksheet, sectionConfig, totalRowNumber, grossTotalSquareMeter, netTotalSquareMeter)
    accumulatedRowOffset += Math.max(0, sectionData.rows.length - 1)
  }
  const outputFileName = buildOutputFileName(workbookData)
  const outputBuffer = await workbook.xlsx.writeBuffer()
  const fileBlob = new Blob([outputBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  })
  const downloadUrl = URL.createObjectURL(fileBlob)
  const temporaryLink = document.createElement("a")
  temporaryLink.href = downloadUrl
  temporaryLink.download = outputFileName
  temporaryLink.click()
  URL.revokeObjectURL(downloadUrl)
}
