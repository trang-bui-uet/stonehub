import templateInvoiceWorkbookFileUrl from "@/assets/template-invoice.xlsx?url"
import templateInvoiceConfigJson from "@/config/template-invoice-config.json"
import ExcelJS from "exceljs"

type SupplierSlabRow = Readonly<{
  lengthNet: number | null
  widthNet: number | null
}>
type SupplierGeneralInfo = Readonly<{
  materialName: string
}>
type StandardSupplierWorkbookData = Readonly<{
  generalInfo: SupplierGeneralInfo
  rows: readonly SupplierSlabRow[]
}>
type MixSupplierWorkbookData = Readonly<{
  sections: readonly Readonly<{
    generalInfo: SupplierGeneralInfo
    rows: readonly SupplierSlabRow[]
  }>[]
}>
type SupplierWorkbookData = StandardSupplierWorkbookData | MixSupplierWorkbookData
type InvoiceExportInput = Readonly<{
  workbookData: SupplierWorkbookData
  customerName: string
  customerCode: string
  legalDocument: string
  phoneNumber: string
  address: string
  unitPrice: number
  requesterName: string
}>
type TemplateInvoiceCellMapping = Readonly<{
  issueDate: string
  customerLine: string
  legalDocumentLine: string
  addressLine: string
  itemNumber: string
  itemName: string
  quantity: string
  unitPrice: string
  amount: string
  subtotalAmount: string
  totalAmount: string
  requesterName: string
}>
type TemplateInvoiceConfig = Readonly<{
  sheetIndex: number
  cellMapping: TemplateInvoiceCellMapping
}>

const TEMPLATE_INVOICE_CONFIG: TemplateInvoiceConfig = templateInvoiceConfigJson as TemplateInvoiceConfig
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const DEFAULT_ITEM_NAME = "Đá tự nhiên"

function calculateSquareMeter(length: number | null, width: number | null): number {
  if (length === null || width === null) {
    return 0
  }
  return (length * width) / CENTIMETER_SQUARE_TO_METER_SQUARE
}

function roundNumberToThreeDigits(value: number): number {
  return Math.round(value * 1000) / 1000
}

function normalizeText(value: string): string {
  return value.trim()
}

function isMixWorkbookData(workbookData: SupplierWorkbookData): workbookData is MixSupplierWorkbookData {
  return "sections" in workbookData
}

function resolveTotalNetSquareMeter(workbookData: SupplierWorkbookData): number {
  if (!isMixWorkbookData(workbookData)) {
    const totalValue = workbookData.rows.reduce((sum: number, row: SupplierSlabRow): number => sum + calculateSquareMeter(row.lengthNet, row.widthNet), 0)
    return roundNumberToThreeDigits(totalValue)
  }
  const totalValue = workbookData.sections.reduce(
    (sectionSum: number, section): number =>
      sectionSum + section.rows.reduce((rowSum: number, row: SupplierSlabRow): number => rowSum + calculateSquareMeter(row.lengthNet, row.widthNet), 0),
    0,
  )
  return roundNumberToThreeDigits(totalValue)
}

function resolveMaterialName(workbookData: SupplierWorkbookData): string {
  if (!isMixWorkbookData(workbookData)) {
    return normalizeText(workbookData.generalInfo.materialName) || DEFAULT_ITEM_NAME
  }
  const firstSection = workbookData.sections[0]
  if (!firstSection) {
    return DEFAULT_ITEM_NAME
  }
  return normalizeText(firstSection.generalInfo.materialName) || DEFAULT_ITEM_NAME
}

function formatCurrentDateLabel(): string {
  const currentDate = new Date()
  const day = String(currentDate.getDate()).padStart(2, "0")
  const month = String(currentDate.getMonth() + 1).padStart(2, "0")
  const year = String(currentDate.getFullYear())
  return `Ngày ${day}/${month}/${year}`
}

type OutputFileNameInput = Readonly<{
  materialName: string
  customerName: string
  containerNumber: string
}>

function sanitizeFileNamePart(value: string, fallbackValue: string): string {
  const normalizedValue = normalizeText(value).replace(/[\\/:*?"<>|]/g, "-")
  if (normalizedValue === "") {
    return fallbackValue
  }
  return normalizedValue
}

function buildOutputFileName(input: OutputFileNameInput): string {
  const safeMaterialName = sanitizeFileNamePart(input.materialName, "material")
  const safeCustomerName = sanitizeFileNamePart(input.customerName, "khach-hang")
  const safeContainerNumber = sanitizeFileNamePart(input.containerNumber, "container")
  return `TP - ${safeMaterialName} - ${safeCustomerName} - ${safeContainerNumber}.xlsx`
}

function setCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, value: string | number): void {
  worksheet.getCell(cellAddress).value = value
}

/**
 * Export invoice file from the invoice template.
 */
export async function exportTemplateInvoiceFile(input: InvoiceExportInput): Promise<void> {
  const templateResponse = await fetch(templateInvoiceWorkbookFileUrl)
  if (!templateResponse.ok) {
    throw new Error("Không thể tải template hóa đơn.")
  }
  const templateArrayBuffer = await templateResponse.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateArrayBuffer)
  const worksheet = workbook.getWorksheet(TEMPLATE_INVOICE_CONFIG.sheetIndex + 1) ?? workbook.worksheets[0]
  if (!worksheet) {
    throw new Error("Template hóa đơn không có sheet hợp lệ.")
  }
  const materialName = resolveMaterialName(input.workbookData)
  const netSquareMeter = resolveTotalNetSquareMeter(input.workbookData)
  const totalAmount = Math.round(netSquareMeter * input.unitPrice)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.issueDate, formatCurrentDateLabel())
  setCellValue(
    worksheet,
    TEMPLATE_INVOICE_CONFIG.cellMapping.customerLine,
    `Họ tên người mua hàng: ${normalizeText(input.customerName)}                Mã cont: ${normalizeText(input.customerCode)}`,
  )
  setCellValue(
    worksheet,
    TEMPLATE_INVOICE_CONFIG.cellMapping.legalDocumentLine,
    `Số giấy tờ pháp lý của cá nhân: ${normalizeText(input.legalDocument)}`,
  )
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.addressLine, `Địa chỉ: ${normalizeText(input.address)}`)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.itemNumber, 1)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.itemName, materialName)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.quantity, netSquareMeter)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.unitPrice, input.unitPrice)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.amount, totalAmount)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.subtotalAmount, totalAmount)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.totalAmount, totalAmount)
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.requesterName, normalizeText(input.requesterName))
  const outputBuffer = await workbook.xlsx.writeBuffer()
  const fileBlob = new Blob([outputBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  })
  const downloadUrl = URL.createObjectURL(fileBlob)
  const temporaryLink = document.createElement("a")
  temporaryLink.href = downloadUrl
  temporaryLink.download = buildOutputFileName({
    materialName,
    customerName: input.customerName,
    containerNumber: input.customerCode,
  })
  temporaryLink.click()
  URL.revokeObjectURL(downloadUrl)
}
