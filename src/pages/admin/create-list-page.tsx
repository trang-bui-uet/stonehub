import supplierConfigJson from "@/config/supplier-config.json"
import { readAjMixWorkbookFromFile } from "@/lib/aj-mix-workbook-reader"
import { readAyMixWorkbookFromFile } from "@/lib/ay-mix-workbook-reader"
import { Button } from "@/components/ui/button"
import {
  Combobox,
  ComboboxContent,
  ComboboxEmpty,
  ComboboxInput,
  ComboboxItem,
  ComboboxList,
} from "@/components/ui/combobox"
import { readAnMixWorkbookFromFile } from "@/lib/an-mix-workbook-reader"
import { readTpMixWorkbookFromFile } from "@/lib/tp-mix-workbook-reader"
import { readSveMixWorkbookFromFile } from "@/lib/sve-mix-workbook-reader"
import { readVkMixWorkbookFromFile } from "@/lib/vk-mix-workbook-reader"
import { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog"
import { Input } from "@/components/ui/input"
import { InputGroup, InputGroupAddon, InputGroupInput, InputGroupText } from "@/components/ui/input-group"
import { Label } from "@/components/ui/label"
import { exportTemplateInvoiceFile } from "@/lib/template-invoice-exporter"
import { exportTemplateMixFile } from "@/lib/template-mix-exporter"
import { exportSupplierListFile } from "@/lib/template-list-exporter"
import { ImagePlus, LoaderCircle, Minus, Plus, Upload, X } from "lucide-react"
import type { ChangeEvent, ReactElement } from "react"
import { useEffect, useMemo, useRef, useState } from "react"

type SupplierConfigItem = Readonly<{
  name: string
  mixed?: boolean
}>

type SupplierConfigRoot = Readonly<{
  suppliers: readonly SupplierConfigItem[]
}>

type StandardSupplierWorkbookData = Awaited<ReturnType<typeof readSupplierWorkbookFromFile>>
type AjMixSupplierWorkbookData = Awaited<ReturnType<typeof readAjMixWorkbookFromFile>>
type AyMixSupplierWorkbookData = Awaited<ReturnType<typeof readAyMixWorkbookFromFile>>
type AnMixSupplierWorkbookData = Awaited<ReturnType<typeof readAnMixWorkbookFromFile>>
type SveMixSupplierWorkbookData = Awaited<ReturnType<typeof readSveMixWorkbookFromFile>>
type VkMixSupplierWorkbookData = Awaited<ReturnType<typeof readVkMixWorkbookFromFile>>
type TpMixSupplierWorkbookData = Awaited<ReturnType<typeof readTpMixWorkbookFromFile>>
type MixSupplierWorkbookData =
  | AjMixSupplierWorkbookData
  | AyMixSupplierWorkbookData
  | AnMixSupplierWorkbookData
  | TpMixSupplierWorkbookData
  | VkMixSupplierWorkbookData
  | SveMixSupplierWorkbookData
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

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const AJ_MIX_SUPPLIER_NAME = "AJ (mix)"
const AY_MIX_SUPPLIER_NAME = "AY (mix)"
const AN_MIX_SUPPLIER_NAME = "AN (mix)"
const SVE_MIX_SUPPLIER_NAME = "SVE (mix)"
const VK_MIX_SUPPLIER_NAME = "VK (mix)"
const TP_MIX_SUPPLIER_NAME = "TP (mix)"
const TP_SUPPLIER_NAME = "TP"
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const DEFAULT_DIMENSION_ADJUSTMENT_INPUT = "0"
const EXCEL_FILE_ACCEPT = ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
const EXCEL_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
] as const
const EXCEL_FILE_EXTENSIONS = [".xlsx", ".xls"] as const
const IMAGE_FILE_ACCEPT = "image/png,image/jpeg,image/jpg,image/webp,image/gif"
const IMAGE_FILE_MIME_TYPES = ["image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif"] as const
const OCR_API_ENDPOINT = `${(import.meta.env.VITE_STONEHUB_BE_URL as string | undefined) ?? "http://127.0.0.1:8000"}/api/v1/invoices/extract`
const VARIANT_DISPLAY_VALUE = "V"
const DEFAULT_INVOICE_FORM_DATA: InvoiceFormData = {
  customerName: "",
  customerCode: "",
  legalDocument: "",
  address: "",
  unitPrice: "0",
  requesterName: "Vũ Thanh Thùy",
}

function getFreshVariantDisplayValue(value: SupplierSlabRow["freshVar"] | undefined): string {
  if (value === "Var") {
    return VARIANT_DISPLAY_VALUE
  }
  return ""
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

function InvoiceExportDialog(props: InvoiceExportDialogProps): ReactElement {
  const customerImageInputReference = useRef<HTMLInputElement | null>(null)
  const [isOpen, setIsOpen] = useState<boolean>(false)
  const [isCreatingInvoice, setIsCreatingInvoice] = useState<boolean>(false)
  const [isExtractingCustomerInfo, setIsExtractingCustomerInfo] = useState<boolean>(false)
  const [uploadedCustomerImageName, setUploadedCustomerImageName] = useState<string>("")
  const [isCustomerImageDragOver, setIsCustomerImageDragOver] = useState<boolean>(false)
  const [invoiceErrorMessage, setInvoiceErrorMessage] = useState<string>("")
  const [invoiceFormData, setInvoiceFormData] = useState<InvoiceFormData>({
    ...DEFAULT_INVOICE_FORM_DATA,
    customerCode: props.defaultCustomerCode,
  })
  useEffect((): void => {
    if (!isOpen) {
      return
    }
    setInvoiceFormData({
      ...DEFAULT_INVOICE_FORM_DATA,
      customerCode: props.defaultCustomerCode,
    })
    setUploadedCustomerImageName("")
    setIsCustomerImageDragOver(false)
    setInvoiceErrorMessage("")
  }, [isOpen, props.defaultCustomerCode])
  useEffect((): (() => void) | void => {
    if (!isOpen) {
      return
    }
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
  }, [isOpen, invoiceFormData.customerCode])
  function handleInvoiceFieldChange(fieldName: keyof InvoiceFormData, value: string): void {
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
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể đọc thông tin khách hàng từ ảnh."
      setInvoiceErrorMessage(message)
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
    const trimmedUnitPrice = invoiceFormData.unitPrice.trim()
    const parsedUnitPrice = trimmedUnitPrice === "" ? 0 : Number(trimmedUnitPrice)
    if (Number.isNaN(parsedUnitPrice) || parsedUnitPrice <= 0) {
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
        unitPrice: parsedUnitPrice,
        requesterName: invoiceFormData.requesterName,
      })
      setIsOpen(false)
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể xuất hóa đơn."
      setInvoiceErrorMessage(message)
    } finally {
      setIsCreatingInvoice(false)
    }
  }
  return (
    <Dialog open={isOpen} onOpenChange={setIsOpen}>
      <DialogTrigger asChild>
        <Button type="button" size="default" variant="outline" disabled={props.isDisabled}>
          Xuất hóa đơn
        </Button>
      </DialogTrigger>
      <DialogContent className="sm:max-w-xl">
        <DialogHeader>
          <DialogTitle>Xuất hóa đơn</DialogTitle>
          <DialogDescription>Nhập thông tin khách hàng và đơn giá. Số lượng sẽ lấy theo Net SQM.</DialogDescription>
        </DialogHeader>
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
              <p className="mt-1 text-xs text-muted-foreground">Hỗ trợ paste từ clipboard (Ctrl+V) khi popup đang mở.</p>
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
          <div className="grid gap-2">
            <Label htmlFor="invoice-unit-price">Đơn giá</Label>
            <InputGroup>
              <InputGroupInput
                id="invoice-unit-price"
                type="number"
                min="0"
                step="1000"
                value={invoiceFormData.unitPrice}
                disabled={isExtractingCustomerInfo}
                onChange={(event: ChangeEvent<HTMLInputElement>): void => handleInvoiceFieldChange("unitPrice", event.target.value)}
              />
              <InputGroupAddon align="inline-end">
                <InputGroupText>VND</InputGroupText>
              </InputGroupAddon>
            </InputGroup>
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
          <p className="rounded-md border border-red-400 bg-red-50 px-3 py-2 text-sm text-red-700">{invoiceErrorMessage}</p>
        ) : null}
        <DialogFooter>
          <Button type="button" onClick={handleCreateInvoice} disabled={isCreatingInvoice || isExtractingCustomerInfo}>
            {isCreatingInvoice ? "Đang xuất hóa đơn..." : "Xuất hóa đơn"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
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

function calculateSquareMeter(length: number | null, width: number | null): number | null {
  const squareMeter = calculateSquareMeterRaw(length, width)
  if (squareMeter === null) {
    return null
  }
  return roundNumberToTwoDigits(squareMeter)
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

function getAdjustedDimension(value: number | null, adjustmentInCentimeter: number): number | null {
  if (value === null) {
    return null
  }
  return value + adjustmentInCentimeter
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

function getAdjustedRows(rows: readonly SupplierSlabRow[], adjustmentInCentimeter: number): readonly SupplierSlabRow[] {
  return rows.map((row: SupplierSlabRow): SupplierSlabRow => ({
    ...row,
    lengthGross: getAdjustedDimension(row.lengthGross, adjustmentInCentimeter),
    widthGross: getAdjustedDimension(row.widthGross, adjustmentInCentimeter),
    lengthNet: getAdjustedDimension(row.lengthNet, adjustmentInCentimeter),
    widthNet: getAdjustedDimension(row.widthNet, adjustmentInCentimeter),
  }))
}

function renderRowTable(rows: readonly SupplierSlabRow[]): ReactElement {
  return (
    <div className="rounded-xl border" style={{ borderColor: "var(--border)" }}>
      <div className="hidden overflow-x-auto md:block">
        <table className="w-full text-sm">
          <thead style={{ backgroundColor: "var(--muted)" }}>
            <tr className="text-left">
              <th className="px-3 py-2">Slab Gross</th>
              <th className="px-3 py-2">L Gross</th>
              <th className="px-3 py-2">W Gross</th>
              <th className="px-3 py-2">GROSS SQM</th>
              <th className="px-3 py-2">Slab Net</th>
              <th className="px-3 py-2">L Net</th>
              <th className="px-3 py-2">W Net</th>
              <th className="px-3 py-2">NET SQM</th>
              <th className="px-3 py-2"></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row: SupplierSlabRow): ReactElement => (
              <tr key={`${row.rowNumber}-${row.slabNoGross}-${row.slabNoNet}`} className="border-t" style={{ borderColor: "var(--border)" }}>
                <td className="px-3 py-2">{row.slabNoGross || "-"}</td>
                <td className="px-3 py-2">{row.lengthGross ?? "-"}</td>
                <td className="px-3 py-2">{row.widthGross ?? "-"}</td>
                <td className="px-3 py-2">{calculateSquareMeter(row.lengthGross, row.widthGross) ?? "-"}</td>
                <td className="px-3 py-2">{row.slabNoNet || "-"}</td>
                <td className="px-3 py-2">{row.lengthNet ?? "-"}</td>
                <td className="px-3 py-2">{row.widthNet ?? "-"}</td>
                <td className="px-3 py-2">{calculateSquareMeter(row.lengthNet, row.widthNet) ?? "-"}</td>
                <td className="px-3 py-2">{getFreshVariantDisplayValue(row.freshVar)}</td>
              </tr>
            ))}
            <tr className="border-t font-semibold" style={{ borderColor: "var(--border)" }}>
              <td className="px-3 py-2">Total</td>
              <td className="px-3 py-2"></td>
              <td className="px-3 py-2">-</td>
              <td className="px-3 py-2">{calculateTotalSquareMeter(rows, (row: SupplierSlabRow): number | null => row.lengthGross, (row: SupplierSlabRow): number | null => row.widthGross)}</td>
              <td className="px-3 py-2">-</td>
              <td className="px-3 py-2">-</td>
              <td className="px-3 py-2">-</td>
              <td className="px-3 py-2">{calculateTotalSquareMeter(rows, (row: SupplierSlabRow): number | null => row.lengthNet, (row: SupplierSlabRow): number | null => row.widthNet)}</td>
              <td className="px-3 py-2">-</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div className="space-y-3 p-3 md:hidden">
        {rows.map((row: SupplierSlabRow): ReactElement => (
          <div key={`${row.rowNumber}-${row.slabNoGross}-${row.slabNoNet}`} className="rounded-lg border p-3 text-sm" style={{ borderColor: "var(--border)" }}>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <span className="font-medium">Slab Gross: </span>
                <span>{row.slabNoGross || "-"}</span>
              </div>
              <div>
                <span className="font-medium">Slab Net: </span>
                <span>{row.slabNoNet || "-"}</span>
              </div>
              <div>
                <span className="font-medium">L Gross: </span>
                <span>{row.lengthGross ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">W Gross: </span>
                <span>{row.widthGross ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">L Net: </span>
                <span>{row.lengthNet ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">W Net: </span>
                <span>{row.widthNet ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">GROSS SQM: </span>
                <span>{calculateSquareMeter(row.lengthGross, row.widthGross) ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">NET SQM: </span>
                <span>{calculateSquareMeter(row.lengthNet, row.widthNet) ?? "-"}</span>
              </div>
              <div>
                <span className="font-medium">Variant: </span>
                <span>{getFreshVariantDisplayValue(row.freshVar)}</span>
              </div>
            </div>
          </div>
        ))}
      </div>
      <div className="border-t px-3 py-2 text-sm" style={{ borderColor: "var(--border)" }}>
        Tổng số dòng đọc được: {rows.length}
      </div>
    </div>
  )
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

export default function CreateListPage(): ReactElement {
  const fileInputReference = useRef<HTMLInputElement | null>(null)
  const [selectedSupplierName, setSelectedSupplierName] = useState<string>(SUPPLIER_CONFIG.suppliers[0]?.name ?? "")
  const [selectedExcelFileName, setSelectedExcelFileName] = useState<string>("")
  const [isReadingFile, setIsReadingFile] = useState<boolean>(false)
  const [isCreatingList, setIsCreatingList] = useState<boolean>(false)
  const [errorMessage, setErrorMessage] = useState<string>("")
  const [workbookData, setWorkbookData] = useState<SupplierWorkbookData | null>(null)
  const [dimensionAdjustmentInput, setDimensionAdjustmentInput] = useState<string>(
    DEFAULT_DIMENSION_ADJUSTMENT_INPUT,
  )
  const supplierNames = useMemo(
    (): readonly string[] =>
      SUPPLIER_CONFIG.suppliers
        .map((supplier: SupplierConfigItem): string => supplier.name)
        .filter((supplierName: string): boolean => supplierName !== TP_SUPPLIER_NAME && supplierName !== TP_MIX_SUPPLIER_NAME),
    [],
  )
  const selectedSupplierConfig = useMemo(
    (): SupplierConfigItem | undefined => SUPPLIER_CONFIG.suppliers.find((supplier: SupplierConfigItem): boolean => supplier.name === selectedSupplierName),
    [selectedSupplierName],
  )
  const dimensionAdjustmentInCentimeter = useMemo((): number => {
    const trimmedInput = dimensionAdjustmentInput.trim()
    if (trimmedInput === "") {
      return 0
    }
    const parsedValue = Number(trimmedInput)
    if (Number.isNaN(parsedValue) || !Number.isInteger(parsedValue)) {
      return 0
    }
    return parsedValue
  }, [dimensionAdjustmentInput])
  const adjustedWorkbookData = useMemo((): SupplierWorkbookData | null => {
    if (!workbookData) {
      return null
    }
    if (isMixWorkbookData(workbookData)) {
      return {
        ...workbookData,
        sections: workbookData.sections.map((section: MixSectionData): MixSectionData => ({
          ...section,
          rows: getAdjustedRows(section.rows, dimensionAdjustmentInCentimeter),
        })),
      }
    }
    return { ...workbookData, rows: getAdjustedRows(workbookData.rows, dimensionAdjustmentInCentimeter) }
  }, [dimensionAdjustmentInCentimeter, workbookData])
  const invoiceCustomerCode = useMemo((): string => resolveContainerNumber(adjustedWorkbookData), [adjustedWorkbookData])
  function clearWorkbookState(): void {
    setSelectedExcelFileName("")
    setWorkbookData(null)
    setErrorMessage("")
    setDimensionAdjustmentInput(DEFAULT_DIMENSION_ADJUSTMENT_INPUT)
    if (fileInputReference.current) {
      fileInputReference.current.value = ""
    }
  }
  function handleSupplierNameChange(value: string | null): void {
    if (!value) {
      return
    }
    if (value === selectedSupplierName) {
      return
    }
    setSelectedSupplierName(value)
    clearWorkbookState()
  }
  function handleDimensionAdjustmentChange(value: string): void {
    if (value.trim() === "") {
      setDimensionAdjustmentInput("")
      return
    }
    const parsedValue = Number(value)
    if (Number.isNaN(parsedValue) || !Number.isInteger(parsedValue)) {
      return
    }
    setDimensionAdjustmentInput(value)
  }
  function increaseDimensionAdjustment(): void {
    setDimensionAdjustmentInput((dimensionAdjustmentInCentimeter + 1).toString())
  }
  function decreaseDimensionAdjustment(): void {
    setDimensionAdjustmentInput((dimensionAdjustmentInCentimeter - 1).toString())
  }
  function handleClearSelectedFile(): void {
    clearWorkbookState()
  }
  async function handleUploadFile(event: ChangeEvent<HTMLInputElement>): Promise<void> {
    const inputFile = event.target.files?.[0]
    if (!inputFile) {
      return
    }
    if (!selectedSupplierName) {
      setErrorMessage("Vui lòng chọn nhà cung cấp trước khi tải file.")
      event.target.value = ""
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
    setDimensionAdjustmentInput(DEFAULT_DIMENSION_ADJUSTMENT_INPUT)
    try {
      let parsedWorkbookData: SupplierWorkbookData
      if (selectedSupplierName === AJ_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readAjMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else if (selectedSupplierName === AY_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readAyMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else if (selectedSupplierName === AN_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readAnMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else if (selectedSupplierName === SVE_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readSveMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else if (selectedSupplierName === VK_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readVkMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else if (selectedSupplierName === TP_MIX_SUPPLIER_NAME) {
        parsedWorkbookData = await readTpMixWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      } else {
        parsedWorkbookData = await readSupplierWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      }
      setWorkbookData(parsedWorkbookData)
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể đọc file Excel."
      setErrorMessage(message)
      setSelectedExcelFileName("")
    } finally {
      setIsReadingFile(false)
      event.target.value = ""
    }
  }
  async function handleCreateList(): Promise<void> {
    if (!adjustedWorkbookData) {
      setErrorMessage("Vui lòng tải dữ liệu trước khi tạo list.")
      return
    }
    setErrorMessage("")
    setIsCreatingList(true)
    try {
      if (selectedSupplierConfig?.mixed === true) {
        await exportTemplateMixFile(adjustedWorkbookData as MixSupplierWorkbookData)
      } else {
        await exportSupplierListFile(adjustedWorkbookData as StandardSupplierWorkbookData)
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể tạo file list."
      setErrorMessage(message)
    } finally {
      setIsCreatingList(false)
    }
  }
  return (
    <div className="max-w-6xl space-y-6 px-3 pb-6 sm:px-4">
      <div className="space-y-2">
        <h1 className="text-2xl font-semibold">Tạo List</h1>
        <p className="text-sm" style={{ color: "var(--muted-foreground)" }}>
          Cách dùng: Chọn nhà cung cấp, chọn file Excel (.xlsx/.xls), điều chỉnh kích thước (+/- cm) nếu cần, sau đó bấm Tạo List để tải file kết quả.
        </p>
      </div>
      <section className="space-y-4 rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
        <div className="grid gap-4 md:grid-cols-2">
          <label className="space-y-2">
            <span className="text-sm font-medium">Nhà cung cấp</span>
            <Combobox items={supplierNames} value={selectedSupplierName} onValueChange={handleSupplierNameChange}>
              <ComboboxInput className="w-full" placeholder="Tìm nhà cung cấp..." />
              <ComboboxContent>
                <ComboboxEmpty>Không tìm thấy nhà cung cấp.</ComboboxEmpty>
                <ComboboxList>
                  {(supplierName: string): ReactElement => (
                    <ComboboxItem key={supplierName} value={supplierName}>
                      {supplierName}
                    </ComboboxItem>
                  )}
                </ComboboxList>
              </ComboboxContent>
            </Combobox>
          </label>
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
        </div>
        {isReadingFile ? <p className="text-sm">Đang đọc dữ liệu từ file Excel...</p> : null}
        {errorMessage ? (
          <p className="rounded-md border border-red-400 bg-red-50 px-3 py-2 text-sm text-red-700">{errorMessage}</p>
        ) : null}
        <div className="space-y-2">
          <label className="block space-y-1">
            <span className="text-sm font-medium">Điều chỉnh kích thước (+/- cm)</span>
            <div className="flex max-w-xs items-center gap-2">
              <Button
                type="button"
                size="icon-sm"
                variant="outline"
                onClick={decreaseDimensionAdjustment}
                disabled={!workbookData || isReadingFile || isCreatingList}
                aria-label="Giảm điều chỉnh kích thước"
              >
                <Minus />
              </Button>
              <input
                className="w-20 rounded-md border px-2 py-2 text-center text-sm sm:w-24 sm:px-3"
                style={{ borderColor: "var(--input)", backgroundColor: "var(--background)" }}
                type="number"
                step="1"
                value={dimensionAdjustmentInput}
                onChange={(event: ChangeEvent<HTMLInputElement>): void => handleDimensionAdjustmentChange(event.target.value)}
                disabled={!workbookData || isReadingFile || isCreatingList}
              />
              <span className="text-sm font-medium">cm</span>
              <Button
                type="button"
                size="icon-sm"
                variant="outline"
                onClick={increaseDimensionAdjustment}
                disabled={!workbookData || isReadingFile || isCreatingList}
                aria-label="Tăng điều chỉnh kích thước"
              >
                <Plus />
              </Button>
            </div>
          </label>
        </div>
        <div className="flex justify-start">
          <div className="flex w-full flex-col gap-2 sm:w-auto sm:flex-row">
            <Button
              type="button"
              size="default"
              className="w-full sm:w-auto"
              onClick={handleCreateList}
              disabled={!workbookData || isCreatingList || isReadingFile}
            >
              {isCreatingList ? "Đang tạo list..." : "Tạo List"}
            </Button>
            <InvoiceExportDialog
              workbookData={adjustedWorkbookData}
              isDisabled={!workbookData || isReadingFile || isCreatingList}
              defaultCustomerCode={invoiceCustomerCode}
            />
          </div>
        </div>
      </section>
      {adjustedWorkbookData ? (
        <section className="space-y-4">
          {isMixWorkbookData(adjustedWorkbookData)
            ? (
              <div className="grid gap-4 md:grid-cols-2">
                {adjustedWorkbookData.sections.map((section: MixSectionData): ReactElement => (
                  <div key={section.name} className="space-y-4">
                    {renderGeneralInfoCard(
                      adjustedWorkbookData.supplierName,
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
                    {renderRowTable(section.rows)}
                  </div>
                ))}
              </div>
            )
            : (
              <div className="space-y-4">
                {renderGeneralInfoCard(
                  adjustedWorkbookData.supplierName,
                  adjustedWorkbookData.generalInfo,
                  adjustedWorkbookData.rows.length,
                  calculateTotalSquareMeter(
                    adjustedWorkbookData.rows,
                    (row: SupplierSlabRow): number | null => row.lengthGross,
                    (row: SupplierSlabRow): number | null => row.widthGross,
                  ),
                  calculateTotalSquareMeter(
                    adjustedWorkbookData.rows,
                    (row: SupplierSlabRow): number | null => row.lengthNet,
                    (row: SupplierSlabRow): number | null => row.widthNet,
                  ),
                  "Thông tin chung",
                )}
                {renderRowTable(adjustedWorkbookData.rows)}
              </div>
            )}
        </section>
      ) : null}
    </div>
  )
}
