import templateWorkbookFileUrl from "@/assets/template-tp.xlsx?url"
import templateConfigJson from "@/config/template-thien-phuc-config.json"
import type { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import ExcelJS from "exceljs"

type SupplierWorkbookData = Awaited<ReturnType<typeof readSupplierWorkbookFromFile>>
type TemplateGeneralInfoMapping = Readonly<{
  materialName: string
}>
type TemplateColumnMapping = Readonly<{
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
type TemplateThienPhucConfig = Readonly<{
  sheetIndex: number
  dataStartRowIndex: number
  generalInfoMapping: TemplateGeneralInfoMapping
  columnMapping: TemplateColumnMapping
}>
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const GROSS_MEASUREMENT_TEXT = "GROSS MEASUREMENT"
const DEFAULT_MATERIAL_NAME_TEXT = "Material Name"
const NET_MEASUREMENT_TEXT = "NET MEASUREMENT"
const VAR_LABEL_TEXT = "V"
const TEMPLATE_CONFIG: TemplateThienPhucConfig = templateConfigJson as TemplateThienPhucConfig

function setCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, value: string | number): void {
  worksheet.getCell(cellAddress).value = value
}

function setCellFormula(
  worksheet: ExcelJS.Worksheet,
  cellAddress: string,
  formula: string,
  result: number,
): void {
  worksheet.getCell(cellAddress).value = { formula, result }
}
function recreateMergedRange(worksheet: ExcelJS.Worksheet, rangeAddress: string): void {
  try {
    worksheet.unMergeCells(rangeAddress)
  } catch {}
  try {
    worksheet.mergeCells(rangeAddress)
  } catch {}
}

function setFooterLabel(worksheet: ExcelJS.Worksheet, cellAddress: string, value: string): void {
  const cell = worksheet.getCell(cellAddress)
  if (typeof cell.value === "string" && cell.value.trim() !== "") {
    return
  }
  setCellValue(worksheet, cellAddress, value)
}

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

function formatCurrentDayMonthText(): string {
  const currentDate = new Date()
  const day = String(currentDate.getDate()).padStart(2, "0")
  const month = String(currentDate.getMonth() + 1).padStart(2, "0")
  return `${day}-${month}`
}

function setGeneralInfo(worksheet: ExcelJS.Worksheet, workbookData: SupplierWorkbookData): void {
  const resolvedMaterialName = normalizeTextValue(workbookData.generalInfo.materialName) || DEFAULT_MATERIAL_NAME_TEXT
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.materialName,
    resolvedMaterialName,
  )
}

function setRowData(
  worksheet: ExcelJS.Worksheet,
  workbookData: SupplierWorkbookData,
): { nextRowNumber: number; grossTotalSquareMeter: number; netTotalSquareMeter: number } {
  let grossTotalSquareMeter = 0
  let netTotalSquareMeter = 0
  for (const [index, row] of workbookData.rows.entries()) {
    const currentRowNumber = TEMPLATE_CONFIG.dataStartRowIndex + index
    const grossSquareMeter = calculateSquareMeter(row.lengthGross, row.widthGross)
    const netSquareMeter = calculateSquareMeter(row.lengthNet, row.widthNet)
    grossTotalSquareMeter += grossSquareMeter
    netTotalSquareMeter += netSquareMeter
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.slabNoGross}${currentRowNumber}`, row.slabNoGross)
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.lengthGross}${currentRowNumber}`, row.lengthGross ?? "")
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.widthGross}${currentRowNumber}`, row.widthGross ?? "")
    setCellFormula(
      worksheet,
      `${TEMPLATE_CONFIG.columnMapping.grossSquareMeter}${currentRowNumber}`,
      `=${TEMPLATE_CONFIG.columnMapping.lengthGross}${currentRowNumber}*${TEMPLATE_CONFIG.columnMapping.widthGross}${currentRowNumber}/${CENTIMETER_SQUARE_TO_METER_SQUARE}`,
      roundNumberToTwoDigits(grossSquareMeter),
    )
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.slabNoNet}${currentRowNumber}`, row.slabNoNet)
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.lengthNet}${currentRowNumber}`, row.lengthNet ?? "")
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.widthNet}${currentRowNumber}`, row.widthNet ?? "")
    setCellFormula(
      worksheet,
      `${TEMPLATE_CONFIG.columnMapping.netSquareMeter}${currentRowNumber}`,
      `=${TEMPLATE_CONFIG.columnMapping.lengthNet}${currentRowNumber}*${TEMPLATE_CONFIG.columnMapping.widthNet}${currentRowNumber}/${CENTIMETER_SQUARE_TO_METER_SQUARE}`,
      roundNumberToTwoDigits(netSquareMeter),
    )
    setCellValue(
      worksheet,
      `${TEMPLATE_CONFIG.columnMapping.freshVar}${currentRowNumber}`,
      row.freshVar === "Var" ? VAR_LABEL_TEXT : "",
    )
  }
  return { nextRowNumber: TEMPLATE_CONFIG.dataStartRowIndex + workbookData.rows.length, grossTotalSquareMeter, netTotalSquareMeter }
}

function prepareTemplateDataRows(worksheet: ExcelJS.Worksheet, dataRowCount: number): number {
  const dataStartRowNumber = TEMPLATE_CONFIG.dataStartRowIndex
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

function setTotalRow(
  worksheet: ExcelJS.Worksheet,
  totalRowNumber: number,
  grossTotalSquareMeter: number,
  netTotalSquareMeter: number,
): void {
  const grossLabelRange = `${TEMPLATE_CONFIG.columnMapping.slabNoGross}${totalRowNumber}:${TEMPLATE_CONFIG.columnMapping.widthGross}${totalRowNumber}`
  const netLabelRange = `${TEMPLATE_CONFIG.columnMapping.slabNoNet}${totalRowNumber}:${TEMPLATE_CONFIG.columnMapping.widthNet}${totalRowNumber}`
  recreateMergedRange(worksheet, grossLabelRange)
  recreateMergedRange(worksheet, netLabelRange)
  setFooterLabel(worksheet, `${TEMPLATE_CONFIG.columnMapping.slabNoGross}${totalRowNumber}`, GROSS_MEASUREMENT_TEXT)
  setFooterLabel(worksheet, `${TEMPLATE_CONFIG.columnMapping.slabNoNet}${totalRowNumber}`, NET_MEASUREMENT_TEXT)
  setCellValue(
    worksheet,
    `${TEMPLATE_CONFIG.columnMapping.grossSquareMeter}${totalRowNumber}`,
    roundNumberToTwoDigits(grossTotalSquareMeter),
  )
  setCellValue(
    worksheet,
    `${TEMPLATE_CONFIG.columnMapping.netSquareMeter}${totalRowNumber}`,
    roundNumberToTwoDigits(netTotalSquareMeter),
  )
}

function buildOutputFileName(workbookData: SupplierWorkbookData): string {
  const resolvedMaterialName = normalizeTextValue(workbookData.generalInfo.materialName) || DEFAULT_MATERIAL_NAME_TEXT
  const dayMonthText = formatCurrentDayMonthText()
  return `TP - ${resolvedMaterialName} ${dayMonthText}.xlsx`
}

/**
 * Export list file from workbook data using TP template.
 */
export async function exportSupplierListFile(workbookData: SupplierWorkbookData): Promise<void> {
  const templateResponse = await fetch(templateWorkbookFileUrl)
  if (!templateResponse.ok) {
    throw new Error("Không thể tải file template.")
  }
  const templateArrayBuffer = await templateResponse.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateArrayBuffer)
  const worksheet =
    workbook.getWorksheet(TEMPLATE_CONFIG.sheetIndex + 1) ?? workbook.worksheets[0]
  if (!worksheet) {
    throw new Error("Template không có sheet hợp lệ.")
  }
  const totalRowNumber = prepareTemplateDataRows(worksheet, workbookData.rows.length)
  setGeneralInfo(worksheet, workbookData)
  const { grossTotalSquareMeter, netTotalSquareMeter } = setRowData(worksheet, workbookData)
  setTotalRow(worksheet, totalRowNumber, grossTotalSquareMeter, netTotalSquareMeter)
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
