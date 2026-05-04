import templateWorkbookFileUrl from "@/assets/template-thien-phuc.xlsx?url"
import templateConfigJson from "@/config/template-thien-phuc-config.json"
import type { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import ExcelJS from "exceljs"

type SupplierWorkbookData = Awaited<ReturnType<typeof readSupplierWorkbookFromFile>>
type TemplateGeneralInfoMapping = Readonly<{
  "container#": string
  materialName: string
  typeOfPolish: string
  numberOfSlabs: string
  loadingDate: string
  invoiceNumber: string
  invoiceDate: string
  containerNumber: string
}>
type TemplateColumnMapping = Readonly<{
  slabNoGross: string
  lengthGross: string
  widthGross: string
  grossSquareMeter: string
  lengthNet: string
  widthNet: string
  netSquareMeter: string
  freshOrVariation: string
}>
type TemplateSummaryRowMapping = Readonly<{
  numberOfSlabs: string
  grossSquareMeter: string
  netSquareMeter: string
  percentage: string
}>
type TemplateSummaryMapping = Readonly<{
  fresh: TemplateSummaryRowMapping
  lineVariation: TemplateSummaryRowMapping
  total: TemplateSummaryRowMapping
}>
type TemplateThienPhucConfig = Readonly<{
  sheetIndex: number
  dataStartRowIndex: number
  generalInfoMapping: TemplateGeneralInfoMapping
  columnMapping: TemplateColumnMapping
  summaryMapping: TemplateSummaryMapping
}>
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const PERCENTAGE_100_TEXT = "100%"
const PERCENTAGE_0_TEXT = "0%"
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

function setGeneralInfo(worksheet: ExcelJS.Worksheet, workbookData: SupplierWorkbookData): void {
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping["container#"],
    normalizeTextValue(workbookData.generalInfo.containerNumber),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.containerNumber,
    normalizeTextValue(workbookData.generalInfo.containerNumber),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.materialName,
    normalizeTextValue(workbookData.generalInfo.materialName),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.typeOfPolish,
    normalizeTextValue(workbookData.generalInfo.typeOfPolish),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.numberOfSlabs,
    workbookData.rows.length,
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.loadingDate,
    normalizeTextValue(workbookData.generalInfo.loadingDate),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.invoiceNumber,
    normalizeTextValue(workbookData.generalInfo.invoiceNumber),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.generalInfoMapping.invoiceDate,
    normalizeTextValue(workbookData.generalInfo.invoiceDate),
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
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.lengthNet}${currentRowNumber}`, row.lengthNet ?? "")
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.widthNet}${currentRowNumber}`, row.widthNet ?? "")
    setCellFormula(
      worksheet,
      `${TEMPLATE_CONFIG.columnMapping.netSquareMeter}${currentRowNumber}`,
      `=${TEMPLATE_CONFIG.columnMapping.lengthNet}${currentRowNumber}*${TEMPLATE_CONFIG.columnMapping.widthNet}${currentRowNumber}/${CENTIMETER_SQUARE_TO_METER_SQUARE}`,
      roundNumberToTwoDigits(netSquareMeter),
    )
    setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.freshOrVariation}${currentRowNumber}`, "Fresh")
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
  workbookData: SupplierWorkbookData,
  grossTotalSquareMeter: number,
  netTotalSquareMeter: number,
): void {
  setCellValue(worksheet, `${TEMPLATE_CONFIG.columnMapping.slabNoGross}${totalRowNumber}`, "Total")
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
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.fresh.numberOfSlabs, workbookData.rows.length)
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.summaryMapping.fresh.grossSquareMeter,
    roundNumberToTwoDigits(grossTotalSquareMeter),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.summaryMapping.fresh.netSquareMeter,
    roundNumberToTwoDigits(netTotalSquareMeter),
  )
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.fresh.percentage, PERCENTAGE_100_TEXT)
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.lineVariation.numberOfSlabs, 0)
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.lineVariation.grossSquareMeter, 0)
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.lineVariation.netSquareMeter, 0)
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.lineVariation.percentage, PERCENTAGE_0_TEXT)
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.total.numberOfSlabs, workbookData.rows.length)
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.summaryMapping.total.grossSquareMeter,
    roundNumberToTwoDigits(grossTotalSquareMeter),
  )
  setCellValue(
    worksheet,
    TEMPLATE_CONFIG.summaryMapping.total.netSquareMeter,
    roundNumberToTwoDigits(netTotalSquareMeter),
  )
  setCellValue(worksheet, TEMPLATE_CONFIG.summaryMapping.total.percentage, PERCENTAGE_100_TEXT)
}

function buildOutputFileName(workbookData: SupplierWorkbookData): string {
  const normalizedSupplierName = workbookData.supplierName.toLowerCase().replaceAll(" ", "-")
  const normalizedContainerNumber = normalizeTextValue(workbookData.generalInfo.containerNumber).replaceAll(" ", "-")
  const suffix = normalizedContainerNumber || "no-container"
  return `list-${normalizedSupplierName}-${suffix}.xlsx`
}

/**
 * Export list file from workbook data using Thiên Phúc template.
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
  setTotalRow(worksheet, totalRowNumber, workbookData, grossTotalSquareMeter, netTotalSquareMeter)
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
