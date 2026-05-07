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
  unitPricesBySection?: readonly number[]
  shippingFee?: number
  immediateDiscount?: number
  depositAmount?: number
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

function roundNumberToTwoDigits(value: number): number {
  return Math.round(value * 100) / 100
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
    return roundNumberToTwoDigits(totalValue)
  }
  const totalValue = workbookData.sections.reduce(
    (sectionSum: number, section): number =>
      sectionSum + section.rows.reduce((rowSum: number, row: SupplierSlabRow): number => rowSum + calculateSquareMeter(row.lengthNet, row.widthNet), 0),
    0,
  )
  return roundNumberToTwoDigits(totalValue)
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
function setQuantityCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, quantity: number): void {
  const quantityCell = worksheet.getCell(cellAddress)
  quantityCell.value = roundNumberToTwoDigits(quantity)
  quantityCell.numFmt = "0.00"
}
function setMoneyCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, amount: number): void {
  const amountCell = worksheet.getCell(cellAddress)
  amountCell.value = Math.round(amount)
  amountCell.numFmt = "#,##0"
}
function setSumFormulaCellValue(worksheet: ExcelJS.Worksheet, cellAddress: string, formula: string): void {
  const formulaCell = worksheet.getCell(cellAddress)
  formulaCell.value = { formula }
  formulaCell.numFmt = "#,##0"
}
function offsetCellAddress(baseCellAddress: string, rowOffset: number): string {
  if (rowOffset === 0) {
    return baseCellAddress
  }
  const match = /^([A-Z]+)(\d+)$/.exec(baseCellAddress)
  if (!match) {
    return baseCellAddress
  }
  const columnPart = match[1]
  const rowPart = Number(match[2])
  return `${columnPart}${rowPart + rowOffset}`
}
type InvoiceLineItem = Readonly<{
  itemName: string
  quantity: number
  unitPrice: number
}>
type InvoiceAdjustmentLine = Readonly<{
  itemName: string
  amount: number
}>
function resolveInvoiceLineItems(input: InvoiceExportInput): readonly InvoiceLineItem[] {
  if (!isMixWorkbookData(input.workbookData) || !input.unitPricesBySection || input.unitPricesBySection.length < 2) {
    const materialName = resolveMaterialName(input.workbookData)
    const netSquareMeter = resolveTotalNetSquareMeter(input.workbookData)
    return [{ itemName: materialName, quantity: netSquareMeter, unitPrice: input.unitPrice }]
  }
  return input.workbookData.sections.map((section, index: number): InvoiceLineItem => {
    const sectionMaterialName = normalizeText(section.generalInfo.materialName) || `${DEFAULT_ITEM_NAME} ${index + 1}`
    const sectionQuantity = roundNumberToTwoDigits(
      section.rows.reduce((sum: number, row: SupplierSlabRow): number => sum + calculateSquareMeter(row.lengthNet, row.widthNet), 0),
    )
    const sectionUnitPrice = input.unitPricesBySection?.[index] ?? input.unitPrice
    return {
      itemName: sectionMaterialName,
      quantity: sectionQuantity,
      unitPrice: sectionUnitPrice,
    }
  })
}
function isPositiveAmount(value: number): boolean {
  return Number.isFinite(value) && value > 0
}
function writeInvoiceAdjustmentLine(
  worksheet: ExcelJS.Worksheet,
  rowOffset: number,
  itemNumber: number,
  line: InvoiceAdjustmentLine,
): void {
  setCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.itemNumber, rowOffset), itemNumber)
  setCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.itemName, rowOffset), line.itemName)
  setMoneyCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.amount, rowOffset), line.amount)
}
function buildAmountSumFormula(): string {
  const startCellAddress = TEMPLATE_INVOICE_CONFIG.cellMapping.amount
  const endCellAddress = offsetCellAddress(startCellAddress, 9)
  return `SUM(${startCellAddress}:${endCellAddress})`
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
  const invoiceLineItems = resolveInvoiceLineItems(input)
  const shippingFee = input.shippingFee ?? 0
  const immediateDiscount = input.immediateDiscount ?? 0
  const depositAmount = input.depositAmount ?? 0
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
  invoiceLineItems.forEach((lineItem: InvoiceLineItem, index: number): void => {
    const amount = roundNumberToTwoDigits(lineItem.quantity * lineItem.unitPrice)
    setCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.itemNumber, index), index + 1)
    setCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.itemName, index), lineItem.itemName)
    setQuantityCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.quantity, index), lineItem.quantity)
    setMoneyCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.unitPrice, index), lineItem.unitPrice)
    setMoneyCellValue(worksheet, offsetCellAddress(TEMPLATE_INVOICE_CONFIG.cellMapping.amount, index), amount)
  })
  let nextItemNumber = invoiceLineItems.length + 1
  let nextRowOffset = invoiceLineItems.length
  if (isPositiveAmount(shippingFee)) {
    writeInvoiceAdjustmentLine(
      worksheet,
      nextRowOffset,
      nextItemNumber,
      {
        itemName: "Cước vận chuyển",
        amount: shippingFee,
      },
    )
    nextItemNumber += 1
    nextRowOffset += 1
  }
  const discountAndDepositAnchorItemNumber = 10
  let currentAdjustmentOffset = discountAndDepositAnchorItemNumber - 1
  let currentAdjustmentItemNumber = discountAndDepositAnchorItemNumber
  if (isPositiveAmount(immediateDiscount)) {
    writeInvoiceAdjustmentLine(
      worksheet,
      currentAdjustmentOffset,
      currentAdjustmentItemNumber,
      {
        itemName: "Trừ chiết khấu tiền ngay",
        amount: -immediateDiscount,
      },
    )
    currentAdjustmentOffset -= 1
    currentAdjustmentItemNumber -= 1
  }
  if (isPositiveAmount(depositAmount)) {
    writeInvoiceAdjustmentLine(
      worksheet,
      currentAdjustmentOffset,
      currentAdjustmentItemNumber,
      {
        itemName: "Trừ cọc",
        amount: -depositAmount,
      },
    )
  }
  setSumFormulaCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.subtotalAmount, buildAmountSumFormula())
  setCellValue(worksheet, TEMPLATE_INVOICE_CONFIG.cellMapping.totalAmount, "")
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
