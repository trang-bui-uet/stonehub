import { Button } from "@/components/ui/button"
import { readTpMixWorkbookFromFile } from "@/lib/tp-mix-workbook-reader"
import { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import { Input } from "@/components/ui/input"
import { InputGroup, InputGroupAddon, InputGroupInput, InputGroupText } from "@/components/ui/input-group"
import { Label } from "@/components/ui/label"
import { exportTemplateInvoiceFile } from "@/lib/template-invoice-exporter"
import { ImagePlus, LoaderCircle, Upload, X } from "lucide-react"
import type { ChangeEvent, ReactElement } from "react"
import { useEffect, useMemo, useRef, useState } from "react"
import * as XLSX from "xlsx"

type StandardSupplierWorkbookData = Awaited<ReturnType<typeof readSupplierWorkbookFromFile>>
type TpMixSupplierWorkbookData = Awaited<ReturnType<typeof readTpMixWorkbookFromFile>>
type MixSupplierWorkbookData = TpMixSupplierWorkbookData
type SupplierWorkbookData = StandardSupplierWorkbookData | MixSupplierWorkbookData
type SupplierSlabRow = StandardSupplierWorkbookData["rows"][number]
type SupplierGeneralInfo = StandardSupplierWorkbookData["generalInfo"]
type MixSectionData = MixSupplierWorkbookData["sections"][number]
type InvoiceFormData = Readonly<{
  customerName: string
  customerCode: string
  legalDocument: string
  address: string
  unitPrice: string
  sectionOneUnitPrice: string
  sectionTwoUnitPrice: string
  shippingFee: string
  immediateDiscount: string
  depositAmount: string
  requesterName: string
}>
type InvoiceExportDialogProps = Readonly<{
  workbookData: SupplierWorkbookData | null
  isDisabled: boolean
  defaultCustomerCode: string
}>
type InvoiceOcrData = Readonly<{
  organization_name: string | null
  tax_code: string | null
  address: string | null
  legal_representative: string | null
  phone_number: string | null
  email: string | null
  raw_text: string
}>
type InvoiceOcrResponse = Readonly<{
  status: string
  data: InvoiceOcrData
}>

const TP_MIX_SUPPLIER_NAME = "TP (mix)"
const TP_SUPPLIER_NAME = "TP"
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const EXCEL_FILE_ACCEPT = ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
const EXCEL_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
] as const
const EXCEL_FILE_EXTENSIONS = [".xlsx", ".xls"] as const
const IMAGE_FILE_ACCEPT = "image/png,image/jpeg,image/jpg,image/webp,image/gif"
const IMAGE_FILE_MIME_TYPES = ["image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif"] as const
const STONEHUB_BE_URL = ((import.meta.env.VITE_STONEHUB_BE_URL as string | undefined) ?? "http://127.0.0.1:8000").replace(/\/$/, "")
const OCR_API_ENDPOINT = `${STONEHUB_BE_URL}/api/v1/invoices/extract`
const SLAB_NO_TOKEN = "SLAB NO."
const PRICE_FIELD_NAMES: readonly (keyof InvoiceFormData)[] = [
  "unitPrice",
  "sectionOneUnitPrice",
  "sectionTwoUnitPrice",
  "shippingFee",
  "immediateDiscount",
  "depositAmount",
] as const
const DEFAULT_INVOICE_FORM_DATA: InvoiceFormData = {
  customerName: "",
  customerCode: "",
  legalDocument: "",
  address: "",
  unitPrice: "0",
  sectionOneUnitPrice: "0",
  sectionTwoUnitPrice: "0",
  shippingFee: "0",
  immediateDiscount: "0",
  depositAmount: "0",
  requesterName: "Vũ Thanh Thùy",
}

function resolveContainerNumber(workbookData: SupplierWorkbookData | null): string {
  if (!workbookData) {
    return ""
  }
  if (isMixWorkbookData(workbookData)) {
    return workbookData.sections[0]?.generalInfo.containerNumber ?? ""
  }
  return workbookData.generalInfo.containerNumber ?? ""
}
function normalizePriceInputValue(value: string): string {
  return value.replace(/[^\d]/g, "")
}
function formatPriceInputValue(value: string): string {
  const normalizedValue = normalizePriceInputValue(value)
  if (normalizedValue === "") {
    return "0"
  }
  return Number(normalizedValue).toLocaleString("en-US")
}
function parseFormattedPrice(value: string): number {
  const normalizedValue = normalizePriceInputValue(value)
  if (normalizedValue === "") {
    return 0
  }
  return Number(normalizedValue)
}
function formatCurrencyValue(value: number): string {
  return Math.round(value).toLocaleString("en-US")
}
function formatSquareMeterValue(value: number): string {
  return roundNumberToTwoDigits(value).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}
function InvoiceExportDialog(props: InvoiceExportDialogProps): ReactElement {
  const customerImageInputReference = useRef<HTMLInputElement | null>(null)
  const [isCreatingInvoice, setIsCreatingInvoice] = useState<boolean>(false)
  const [isExtractingCustomerInfo, setIsExtractingCustomerInfo] = useState<boolean>(false)
  const [uploadedCustomerImageName, setUploadedCustomerImageName] = useState<string>("")
  const [isCustomerImageDragOver, setIsCustomerImageDragOver] = useState<boolean>(false)
  const [invoiceErrorMessage, setInvoiceErrorMessage] = useState<string>("")
  const [invoiceFormData, setInvoiceFormData] = useState<InvoiceFormData>({
    ...DEFAULT_INVOICE_FORM_DATA,
    customerCode: props.defaultCustomerCode,
  })
  const isMixInvoice: boolean = Boolean(props.workbookData && isMixWorkbookData(props.workbookData))
  const parsedUnitPriceForDisplay = parseFormattedPrice(invoiceFormData.unitPrice)
  const parsedSectionOneUnitPriceForDisplay = parseFormattedPrice(invoiceFormData.sectionOneUnitPrice)
  const parsedSectionTwoUnitPriceForDisplay = parseFormattedPrice(invoiceFormData.sectionTwoUnitPrice)
  const parsedShippingFeeForDisplay = parseFormattedPrice(invoiceFormData.shippingFee)
  const parsedImmediateDiscountForDisplay = parseFormattedPrice(invoiceFormData.immediateDiscount)
  const parsedDepositAmountForDisplay = parseFormattedPrice(invoiceFormData.depositAmount)
  const mixOneRows: readonly SupplierSlabRow[] =
    props.workbookData && isMixWorkbookData(props.workbookData) ? props.workbookData.sections[0]?.rows ?? [] : []
  const mixTwoRows: readonly SupplierSlabRow[] =
    props.workbookData && isMixWorkbookData(props.workbookData) ? props.workbookData.sections[1]?.rows ?? [] : []
  const mixOneTotalNetSquareMeter = calculateTotalSquareMeter(
    mixOneRows,
    (row: SupplierSlabRow): number | null => row.lengthNet,
    (row: SupplierSlabRow): number | null => row.widthNet,
  )
  const mixTwoTotalNetSquareMeter = calculateTotalSquareMeter(
    mixTwoRows,
    (row: SupplierSlabRow): number | null => row.lengthNet,
    (row: SupplierSlabRow): number | null => row.widthNet,
  )
  const mixOneAmount = Math.round(mixOneTotalNetSquareMeter * parsedSectionOneUnitPriceForDisplay)
  const mixTwoAmount = Math.round(mixTwoTotalNetSquareMeter * parsedSectionTwoUnitPriceForDisplay)
  const mixTotalAmount = mixOneAmount + mixTwoAmount
  const mixFinalAmount = mixTotalAmount + parsedShippingFeeForDisplay - parsedImmediateDiscountForDisplay - parsedDepositAmountForDisplay
  const normalRows: readonly SupplierSlabRow[] =
    props.workbookData && !isMixWorkbookData(props.workbookData) ? props.workbookData.rows : []
  const normalTotalNetSquareMeter = calculateTotalSquareMeter(
    normalRows,
    (row: SupplierSlabRow): number | null => row.lengthNet,
    (row: SupplierSlabRow): number | null => row.widthNet,
  )
  const normalTotalAmount = Math.round(normalTotalNetSquareMeter * parsedUnitPriceForDisplay)
  const normalFinalAmount =
    normalTotalAmount + parsedShippingFeeForDisplay - parsedImmediateDiscountForDisplay - parsedDepositAmountForDisplay
  useEffect((): void => {
    setInvoiceFormData({
      ...DEFAULT_INVOICE_FORM_DATA,
      customerCode: props.defaultCustomerCode,
    })
    setUploadedCustomerImageName("")
    setIsCustomerImageDragOver(false)
    setInvoiceErrorMessage("")
  }, [props.defaultCustomerCode, props.workbookData])
  useEffect((): (() => void) | void => {
    async function handleWindowPaste(event: ClipboardEvent): Promise<void> {
      const clipboardFile = resolveImageFileFromClipboard(event.clipboardData)
      if (!clipboardFile) {
        return
      }
      event.preventDefault()
      await extractCustomerInfoFromImageFile(clipboardFile)
    }
    window.addEventListener("paste", handleWindowPaste)
    return (): void => {
      window.removeEventListener("paste", handleWindowPaste)
    }
  }, [invoiceFormData.customerCode])
  function handleInvoiceFieldChange(fieldName: keyof InvoiceFormData, value: string): void {
    if (PRICE_FIELD_NAMES.includes(fieldName)) {
      const formattedPrice = formatPriceInputValue(value)
      setInvoiceFormData((currentData: InvoiceFormData): InvoiceFormData => ({ ...currentData, [fieldName]: formattedPrice }))
      return
    }
    setInvoiceFormData((currentData: InvoiceFormData): InvoiceFormData => ({ ...currentData, [fieldName]: value }))
  }
  function resolveImageFileFromClipboard(clipboardData: DataTransfer | null): File | null {
    if (!clipboardData) {
      return null
    }
    const imageItem = Array.from(clipboardData.items).find(
      (clipboardItem: DataTransferItem): boolean => clipboardItem.kind === "file" && clipboardItem.type.startsWith("image/"),
    )
    if (!imageItem) {
      return null
    }
    return imageItem.getAsFile()
  }
  function isImageFile(file: File): boolean {
    return IMAGE_FILE_MIME_TYPES.includes(file.type as (typeof IMAGE_FILE_MIME_TYPES)[number])
  }
  async function extractCustomerInfoFromImageFile(file: File): Promise<void> {
    if (!isImageFile(file)) {
      setInvoiceErrorMessage("Chỉ chấp nhận ảnh PNG, JPG, WEBP hoặc GIF.")
      return
    }
    setIsExtractingCustomerInfo(true)
    setInvoiceErrorMessage("")
    setUploadedCustomerImageName(file.name || "clipboard-image.png")
    try {
      const formData = new FormData()
      formData.append("image", file)
      const response = await fetch(OCR_API_ENDPOINT, {
        method: "POST",
        body: formData,
      })
      const responseBody = (await response.json()) as InvoiceOcrResponse | { detail?: string }
      if (!response.ok) {
        const detailMessage = "detail" in responseBody && typeof responseBody.detail === "string" ? responseBody.detail : "Không thể trích xuất thông tin từ ảnh."
        throw new Error(detailMessage)
      }
      if (!("data" in responseBody)) {
        throw new Error("Phản hồi OCR không hợp lệ.")
      }
      const extractedData = responseBody.data
      setInvoiceFormData((currentData: InvoiceFormData): InvoiceFormData => ({
        ...currentData,
        customerName: extractedData.organization_name ?? currentData.customerName,
        legalDocument: extractedData.tax_code ?? currentData.legalDocument,
        address: extractedData.address ?? currentData.address,
      }))
    } catch (_err: unknown) {
      setInvoiceErrorMessage("Không thể đọc ảnh thông tin khách hàng")
    } finally {
      setIsExtractingCustomerInfo(false)
    }
  }
  async function handleCustomerImageInputChange(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    const imageFile = event.target.files?.[0]
    if (!imageFile) {
      return
    }
    await extractCustomerInfoFromImageFile(imageFile)
    event.target.value = ""
  }
  async function handleCustomerImageDrop(event: React.DragEvent<HTMLDivElement>): Promise<void> {
    event.preventDefault()
    setIsCustomerImageDragOver(false)
    const droppedFile = event.dataTransfer.files[0]
    if (!droppedFile) {
      return
    }
    await extractCustomerInfoFromImageFile(droppedFile)
  }
  function handleCustomerImageDragOver(event: React.DragEvent<HTMLDivElement>): void {
    event.preventDefault()
    setIsCustomerImageDragOver(true)
  }
  function handleCustomerImageDragLeave(event: React.DragEvent<HTMLDivElement>): void {
    event.preventDefault()
    setIsCustomerImageDragOver(false)
  }
  async function handleCreateInvoice(): Promise<void> {
    if (!props.workbookData) {
      setInvoiceErrorMessage("Vui lòng tải dữ liệu trước khi xuất hóa đơn.")
      return
    }
    const parsedUnitPrice = parseFormattedPrice(invoiceFormData.unitPrice)
    const parsedSectionOneUnitPrice = parseFormattedPrice(invoiceFormData.sectionOneUnitPrice)
    const parsedSectionTwoUnitPrice = parseFormattedPrice(invoiceFormData.sectionTwoUnitPrice)
    const parsedShippingFee = parseFormattedPrice(invoiceFormData.shippingFee)
    const parsedImmediateDiscount = parseFormattedPrice(invoiceFormData.immediateDiscount)
    const parsedDepositAmount = parseFormattedPrice(invoiceFormData.depositAmount)
    if (isMixInvoice) {
      if (Number.isNaN(parsedSectionOneUnitPrice) || parsedSectionOneUnitPrice <= 0) {
        setInvoiceErrorMessage("Đơn giá section 1 phải lớn hơn 0.")
        return
      }
      if (Number.isNaN(parsedSectionTwoUnitPrice) || parsedSectionTwoUnitPrice <= 0) {
        setInvoiceErrorMessage("Đơn giá section 2 phải lớn hơn 0.")
        return
      }
    } else if (Number.isNaN(parsedUnitPrice) || parsedUnitPrice <= 0) {
      setInvoiceErrorMessage("Đơn giá phải lớn hơn 0.")
      return
    }
    setInvoiceErrorMessage("")
    setIsCreatingInvoice(true)
    try {
      await exportTemplateInvoiceFile({
        workbookData: props.workbookData as StandardSupplierWorkbookData | MixSupplierWorkbookData,
        customerName: invoiceFormData.customerName,
        customerCode: invoiceFormData.customerCode,
        legalDocument: invoiceFormData.legalDocument,
        phoneNumber: "",
        address: invoiceFormData.address,
        unitPrice: isMixInvoice ? parsedSectionOneUnitPrice : parsedUnitPrice,
        unitPricesBySection: isMixInvoice ? [parsedSectionOneUnitPrice, parsedSectionTwoUnitPrice] : undefined,
        shippingFee: parsedShippingFee,
        immediateDiscount: parsedImmediateDiscount,
        depositAmount: parsedDepositAmount,
        requesterName: invoiceFormData.requesterName,
      })
      setInvoiceErrorMessage("Xuất hóa đơn thành công.")
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể xuất hóa đơn."
      setInvoiceErrorMessage(message)
    } finally {
      setIsCreatingInvoice(false)
    }
  }
  return (
    <section className="space-y-4 rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
      <div>
        <h2 className="text-base font-semibold">Xuất hóa đơn</h2>
        <p className="text-sm" style={{ color: "var(--muted-foreground)" }}>
          Nhập thông tin khách hàng và đơn giá. Số lượng sẽ lấy theo Net SQM.
        </p>
      </div>
      <div className="grid gap-3 py-1">
          <div className="grid gap-2">
            <Label htmlFor="invoice-customer-image-upload">Ảnh thông tin khách hàng</Label>
            <div
              className={`flex min-h-24 cursor-pointer flex-col items-center justify-center rounded-md border border-dashed px-3 py-2 text-center ${isCustomerImageDragOver ? "border-primary bg-primary/5" : ""}`}
              onClick={(): void => customerImageInputReference.current?.click()}
              onDrop={handleCustomerImageDrop}
              onDragOver={handleCustomerImageDragOver}
              onDragLeave={handleCustomerImageDragLeave}
            >
              <div className="flex items-center gap-2 text-sm font-medium">
                <Upload className="size-4" />
                <span>Kéo thả hoặc bấm để tải ảnh</span>
              </div>
              <p className="mt-1 text-xs text-muted-foreground">Hỗ trợ paste từ clipboard (Ctrl+V).</p>
              {uploadedCustomerImageName ? (
                <div className="mt-2 flex items-center gap-1 text-xs">
                  {isExtractingCustomerInfo ? <LoaderCircle className="size-3 animate-spin" /> : <ImagePlus className="size-3" />}
                  <span>{uploadedCustomerImageName}</span>
                </div>
              ) : null}
              <input
                ref={customerImageInputReference}
                id="invoice-customer-image-upload"
                className="hidden"
                type="file"
                accept={IMAGE_FILE_ACCEPT}
                onChange={handleCustomerImageInputChange}
                disabled={isExtractingCustomerInfo || isCreatingInvoice}
              />
            </div>
          </div>
          <div className="grid gap-2">
            <Label htmlFor="invoice-customer-name">Họ tên người mua hàng</Label>
            <Input
              id="invoice-customer-name"
              value={invoiceFormData.customerName}
              disabled={isExtractingCustomerInfo}
              onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("customerName", event.target.value)}
            />
          </div>
          <div className="grid gap-2">
            <Label htmlFor="invoice-customer-code">Mã cont</Label>
            <Input
              id="invoice-customer-code"
              value={invoiceFormData.customerCode}
              disabled={isExtractingCustomerInfo}
              onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("customerCode", event.target.value)}
            />
          </div>
          <div className="grid gap-2">
            <Label htmlFor="invoice-legal-document">Số giấy tờ pháp lý</Label>
            <Input
              id="invoice-legal-document"
              value={invoiceFormData.legalDocument}
              disabled={isExtractingCustomerInfo}
              onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("legalDocument", event.target.value)}
            />
          </div>
          <div className="grid gap-2">
            <Label htmlFor="invoice-address">Địa chỉ</Label>
            <Input
              id="invoice-address"
              value={invoiceFormData.address}
              disabled={isExtractingCustomerInfo}
              onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("address", event.target.value)}
            />
          </div>
          {isMixInvoice
            ? (
              <div className="grid gap-2 sm:grid-cols-2">
                <div className="grid gap-2">
                  <Label htmlFor="invoice-section-one-unit-price">Đơn giá mix1</Label>
                  <InputGroup>
                    <InputGroupInput
                      id="invoice-section-one-unit-price"
                      type="text"
                      inputMode="numeric"
                      value={invoiceFormData.sectionOneUnitPrice}
                      disabled={isExtractingCustomerInfo}
                      onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("sectionOneUnitPrice", event.target.value)}
                    />
                    <InputGroupAddon align="inline-end">
                      <InputGroupText>VND</InputGroupText>
                    </InputGroupAddon>
                  </InputGroup>
                </div>
                <div className="grid gap-2">
                  <Label htmlFor="invoice-section-two-unit-price">Đơn giá mix2</Label>
                  <InputGroup>
                    <InputGroupInput
                      id="invoice-section-two-unit-price"
                      type="text"
                      inputMode="numeric"
                      value={invoiceFormData.sectionTwoUnitPrice}
                      disabled={isExtractingCustomerInfo}
                      onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("sectionTwoUnitPrice", event.target.value)}
                    />
                    <InputGroupAddon align="inline-end">
                      <InputGroupText>VND</InputGroupText>
                    </InputGroupAddon>
                  </InputGroup>
                </div>
              </div>
            )
            : (
              <div className="grid gap-2">
                <Label htmlFor="invoice-unit-price">Đơn giá</Label>
                <InputGroup>
                  <InputGroupInput
                    id="invoice-unit-price"
                    type="text"
                    inputMode="numeric"
                    value={invoiceFormData.unitPrice}
                    disabled={isExtractingCustomerInfo}
                    onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("unitPrice", event.target.value)}
                  />
                  <InputGroupAddon align="inline-end">
                    <InputGroupText>VND</InputGroupText>
                  </InputGroupAddon>
                </InputGroup>
              </div>
            )}
          <div className="grid gap-2 sm:grid-cols-3">
            <div className="grid gap-2">
              <Label htmlFor="invoice-shipping-fee">Cước vận chuyển</Label>
              <InputGroup>
                <InputGroupInput
                  id="invoice-shipping-fee"
                  type="text"
                  inputMode="numeric"
                  value={invoiceFormData.shippingFee}
                  disabled={isExtractingCustomerInfo}
                  onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("shippingFee", event.target.value)}
                />
                <InputGroupAddon align="inline-end">
                  <InputGroupText>VND</InputGroupText>
                </InputGroupAddon>
              </InputGroup>
            </div>
            <div className="grid gap-2">
              <Label htmlFor="invoice-immediate-discount">Trừ chiết khấu tiền ngay</Label>
              <InputGroup>
                <InputGroupInput
                  id="invoice-immediate-discount"
                  type="text"
                  inputMode="numeric"
                  value={invoiceFormData.immediateDiscount}
                  disabled={isExtractingCustomerInfo}
                  onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("immediateDiscount", event.target.value)}
                />
                <InputGroupAddon align="inline-end">
                  <InputGroupText>VND</InputGroupText>
                </InputGroupAddon>
              </InputGroup>
            </div>
            <div className="grid gap-2">
              <Label htmlFor="invoice-deposit-amount">Trừ cọc</Label>
              <InputGroup>
                <InputGroupInput
                  id="invoice-deposit-amount"
                  type="text"
                  inputMode="numeric"
                  value={invoiceFormData.depositAmount}
                  disabled={isExtractingCustomerInfo}
                  onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("depositAmount", event.target.value)}
                />
                <InputGroupAddon align="inline-end">
                  <InputGroupText>VND</InputGroupText>
                </InputGroupAddon>
              </InputGroup>
            </div>
          </div>
          <div className="rounded-md border px-3 py-2 text-sm" style={{ borderColor: "var(--border)", backgroundColor: "var(--muted)" }}>
            {isMixInvoice ? (
              <p className="font-semibold">{`Tổng tiền = ${formatSquareMeterValue(mixOneTotalNetSquareMeter)} x ${formatCurrencyValue(parsedSectionOneUnitPriceForDisplay)} + ${formatSquareMeterValue(mixTwoTotalNetSquareMeter)} x ${formatCurrencyValue(parsedSectionTwoUnitPriceForDisplay)} + ${formatCurrencyValue(parsedShippingFeeForDisplay)} - ${formatCurrencyValue(parsedImmediateDiscountForDisplay)} - ${formatCurrencyValue(parsedDepositAmountForDisplay)} = ${formatCurrencyValue(mixFinalAmount)} VND`}</p>
            ) : (
              <p>{`Tổng tiền = ${formatSquareMeterValue(normalTotalNetSquareMeter)} x ${formatCurrencyValue(parsedUnitPriceForDisplay)} + ${formatCurrencyValue(parsedShippingFeeForDisplay)} - ${formatCurrencyValue(parsedImmediateDiscountForDisplay)} - ${formatCurrencyValue(parsedDepositAmountForDisplay)} = ${formatCurrencyValue(normalFinalAmount)} VND`}</p>
            )}
          </div>
          <div className="grid gap-2">
            <Label htmlFor="invoice-requester-name">Người đề nghị xuất hóa đơn</Label>
            <Input
              id="invoice-requester-name"
              value={invoiceFormData.requesterName}
              disabled={isExtractingCustomerInfo}
              onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("requesterName", event.target.value)}
            />
          </div>
      </div>
      {invoiceErrorMessage ? (
        <p className={`rounded-md border px-3 py-2 text-sm ${invoiceErrorMessage.includes("thành công") ? "border-emerald-400 bg-emerald-50 text-emerald-700" : "border-red-400 bg-red-50 text-red-700"}`}>
          {invoiceErrorMessage}
        </p>
      ) : null}
      <div className="flex justify-start">
        <Button type="button" onClick={handleCreateInvoice} disabled={props.isDisabled || isCreatingInvoice || isExtractingCustomerInfo}>
          {isCreatingInvoice ? "Đang xuất hóa đơn..." : "Xuất hóa đơn"}
        </Button>
      </div>
    </section>
  )
}
function roundNumberToTwoDigits(value: number): number {
  return Math.round(value * 100) / 100
}
function calculateSquareMeterRaw(length: number | null, width: number | null): number | null {
  if (length === null || width === null) {
    return null
  }
  return (length * width) / CENTIMETER_SQUARE_TO_METER_SQUARE
}
function calculateTotalSquareMeter(
  rows: readonly SupplierSlabRow[],
  getLength: (row: SupplierSlabRow) => number | null,
  getWidth: (row: SupplierSlabRow) => number | null,
): number {
  const totalValue = rows.reduce((sum: number, row: SupplierSlabRow): number => {
    const squareMeter = calculateSquareMeterRaw(getLength(row), getWidth(row))
    return sum + (squareMeter ?? 0)
  }, 0)
  return roundNumberToTwoDigits(totalValue)
}
function isExcelFile(file: File): boolean {
  const fileName = file.name.toLowerCase()
  const hasExcelExtension = EXCEL_FILE_EXTENSIONS.some((extension: string): boolean => fileName.endsWith(extension))
  if (hasExcelExtension) {
    return true
  }
  return EXCEL_FILE_MIME_TYPES.includes(file.type as (typeof EXCEL_FILE_MIME_TYPES)[number])
}
function isMixWorkbookData(workbookData: SupplierWorkbookData): workbookData is MixSupplierWorkbookData {
  return "sections" in workbookData
}
function getCellText(sheet: XLSX.WorkSheet, cellAddress: string): string {
  const cellValue = sheet[cellAddress]?.v
  if (cellValue === undefined || cellValue === null) {
    return ""
  }
  return String(cellValue).trim()
}
function hasSlabNoTokenInRow(sheet: XLSX.WorkSheet, rowNumber: number): boolean {
  const firstColumnCode = "A".charCodeAt(0)
  const lastColumnCode = "Z".charCodeAt(0)
  for (let columnCode = firstColumnCode; columnCode <= lastColumnCode; columnCode += 1) {
    const columnName = String.fromCharCode(columnCode)
    const cellText = getCellText(sheet, `${columnName}${rowNumber}`).toUpperCase()
    if (cellText.includes(SLAB_NO_TOKEN)) {
      return true
    }
  }
  return false
}
function countRowsContainingSlabNoToken(sheet: XLSX.WorkSheet): number {
  const workbookRange = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null
  if (!workbookRange) {
    return 0
  }
  let rowCount = 0
  const lastRowNumber = workbookRange.e.r + 1
  for (let rowNumber = 1; rowNumber <= lastRowNumber; rowNumber += 1) {
    if (hasSlabNoTokenInRow(sheet, rowNumber)) {
      rowCount += 1
    }
  }
  return rowCount
}
async function detectTpSupplierName(file: File): Promise<string> {
  const fileBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(fileBuffer, { type: "array", cellStyles: true })
  const firstSheetName = workbook.SheetNames[0] ?? ""
  const sheet = workbook.Sheets[firstSheetName]
  if (!sheet) {
    throw new Error("File Excel không có sheet hợp lệ.")
  }
  const rowCount = countRowsContainingSlabNoToken(sheet)
  if (rowCount === 1) {
    return TP_SUPPLIER_NAME
  }
  if (rowCount === 2) {
    return TP_MIX_SUPPLIER_NAME
  }
  if (rowCount < 1) {
    throw new Error('Không tìm thấy dòng chứa "SLAB NO." trong file Excel.')
  }
  throw new Error('File không hợp lệ: số dòng chứa "SLAB NO." phải là 1 (TP) hoặc 2 (TP mix).')
}
function renderGeneralInfoCard(
  supplierName: string,
  generalInfo: SupplierGeneralInfo,
  numberOfRows: number,
  totalGrossSquareMeter: number,
  totalNetSquareMeter: number,
  title: string,
): ReactElement {
  return (
    <div className="rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
      <h2 className="text-base font-semibold">{title}</h2>
      <div className="mt-3 grid gap-3 text-sm md:grid-cols-2">
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Nhà cung cấp: </span>
          <span className="truncate">{supplierName || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Container Number: </span>
          <span className="truncate">{generalInfo.containerNumber || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Material: </span>
          <span className="truncate">{generalInfo.materialName || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Type of polish: </span>
          <span className="truncate">{generalInfo.typeOfPolish || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Số lượng slabs: </span>
          <span className="truncate">{numberOfRows || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Loading date: </span>
          <span className="truncate">{generalInfo.loadingDate || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Invoice Number: </span>
          <span className="truncate">{generalInfo.invoiceNumber || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Invoice Date: </span>
          <span className="truncate">{generalInfo.invoiceDate || "XXX"}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Total Gross SQM: </span>
          <span className="truncate">{totalGrossSquareMeter}</span>
        </div>
        <div className="flex min-w-0 items-center gap-1">
          <span className="shrink-0 font-medium">Total Net SQM: </span>
          <span className="truncate">{totalNetSquareMeter}</span>
        </div>
      </div>
    </div>
  )
}

export default function MenuBPage(): ReactElement {
  const fileInputReference = useRef<HTMLInputElement | null>(null)
  const [selectedExcelFileName, setSelectedExcelFileName] = useState<string>("")
  const [isReadingFile, setIsReadingFile] = useState<boolean>(false)
  const [errorMessage, setErrorMessage] = useState<string>("")
  const [workbookData, setWorkbookData] = useState<SupplierWorkbookData | null>(null)
  const invoiceCustomerCode = useMemo((): string => resolveContainerNumber(workbookData), [workbookData])
  function clearWorkbookState(): void {
    setSelectedExcelFileName("")
    setWorkbookData(null)
    setErrorMessage("")
    if (fileInputReference.current) {
      fileInputReference.current.value = ""
    }
  }
  function handleClearSelectedFile(): void {
    clearWorkbookState()
  }
  async function handleUploadFile(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    const inputFile = event.target.files?.[0]
    if (!inputFile) {
      return
    }
    if (!isExcelFile(inputFile)) {
      setErrorMessage("Chỉ chấp nhận file Excel (.xlsx, .xls).")
      event.target.value = ""
      return
    }
    setIsReadingFile(true)
    setErrorMessage("")
    setSelectedExcelFileName(inputFile.name)
    setWorkbookData(null)
    try {
      const detectedSupplierName = await detectTpSupplierName(inputFile)
      let parsedWorkbookData: SupplierWorkbookData
      if (detectedSupplierName === TP_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readTpMixWorkbookFromFile({ file: inputFile, supplierName: detectedSupplierName })
      } else {
        parsedWorkbookData = await readSupplierWorkbookFromFile({ file: inputFile, supplierName: detectedSupplierName })
      }
      setWorkbookData(parsedWorkbookData)
    } catch (err: unknown) {
      const fallbackMessage = err instanceof Error ? err.message : "Không thể đọc file Excel."
      const uploadErrorMessage = fallbackMessage.includes("SLAB NO.") ? "List không hợp lệ" : fallbackMessage
      setErrorMessage(uploadErrorMessage)
      setSelectedExcelFileName("")
    } finally {
      setIsReadingFile(false)
      event.target.value = ""
    }
  }

  return (
    <div className="max-w-6xl space-y-6 px-3 pb-6 sm:px-4">
      <div className="space-y-2">
        <h1 className="text-2xl font-semibold">Xuất hóa đơn</h1>
        <p className="text-sm" style={{ color: "var(--muted-foreground)" }}>
          Cách dùng: Chọn file Excel (.xlsx/.xls), sau khi tải thành công nhập thông tin và bấm Xuất hóa đơn.
        </p>
      </div>
      <section className="space-y-4 rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
        <label className="space-y-2">
          <span className="text-sm font-medium">File Excel</span>
          <div className="flex items-center gap-3 rounded-md border px-3 py-2" style={{ borderColor: "var(--input)" }}>
            <Button
              type="button"
              variant="outline"
              className="border-dashed"
              onClick={(): void => fileInputReference.current?.click()}
              disabled={isReadingFile}
            >
              Chọn file
            </Button>
            <span className="truncate text-sm" style={{ color: "var(--muted-foreground)" }}>
              {selectedExcelFileName || "Chưa chọn file"}
            </span>
            {selectedExcelFileName ? (
              <Button
                type="button"
                variant="ghost"
                size="icon-sm"
                onClick={handleClearSelectedFile}
                disabled={isReadingFile}
                aria-label="Xóa file đã chọn"
              >
                <X />
              </Button>
            ) : null}
            <input
              ref={fileInputReference}
              className="hidden"
              type="file"
              accept={EXCEL_FILE_ACCEPT}
              onChange={handleUploadFile}
              disabled={isReadingFile}
            />
          </div>
        </label>
        {isReadingFile ? <p className="text-sm">Đang đọc dữ liệu từ file Excel...</p> : null}
        {errorMessage ? (
          <p className="mt-2 rounded-md border border-red-400 bg-red-50 px-3 py-2 text-sm text-red-700">{errorMessage}</p>
        ) : null}
      </section>
      {workbookData ? (
        <section className="space-y-4">
          {isMixWorkbookData(workbookData)
            ? (
              <div className="grid gap-4 md:grid-cols-2">
                {workbookData.sections.map((section: MixSectionData): ReactElement => (
                  <div key={section.name} className="space-y-4">
                    {renderGeneralInfoCard(
                      workbookData.supplierName,
                      section.generalInfo,
                      section.rows.length,
                      calculateTotalSquareMeter(
                        section.rows,
                        (row: SupplierSlabRow): number | null => row.lengthGross,
                        (row: SupplierSlabRow): number | null => row.widthGross,
                      ),
                      calculateTotalSquareMeter(
                        section.rows,
                        (row: SupplierSlabRow): number | null => row.lengthNet,
                        (row: SupplierSlabRow): number | null => row.widthNet,
                      ),
                      `Thông tin chung ${section.name}`,
                    )}
                  </div>
                ))}
              </div>
            )
            : (
              <div className="space-y-4">
                {renderGeneralInfoCard(
                  workbookData.supplierName,
                  workbookData.generalInfo,
                  workbookData.rows.length,
                  calculateTotalSquareMeter(
                    workbookData.rows,
                    (row: SupplierSlabRow): number | null => row.lengthGross,
                    (row: SupplierSlabRow): number | null => row.widthGross,
                  ),
                  calculateTotalSquareMeter(
                    workbookData.rows,
                    (row: SupplierSlabRow): number | null => row.lengthNet,
                    (row: SupplierSlabRow): number | null => row.widthNet,
                  ),
                  "Thông tin chung",
                )}
              </div>
            )}
        </section>
      ) : null}
      {workbookData ? (
        <section className="mt-2">
          <InvoiceExportDialog workbookData={workbookData} isDisabled={isReadingFile} defaultCustomerCode={invoiceCustomerCode} />
        </section>
      ) : null}
    </div>
  )
}

