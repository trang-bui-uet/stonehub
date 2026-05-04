import supplierConfigJson from "@/config/supplier-config.json"
import { Button } from "@/components/ui/button"
import { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import { exportSupplierListFile } from "@/lib/template-list-exporter"
import { Minus, Plus } from "lucide-react"
import type { ChangeEvent, ReactElement } from "react"
import { useMemo, useState } from "react"

type SupplierConfigItem = Readonly<{
  name: string
}>

type SupplierConfigRoot = Readonly<{
  suppliers: readonly SupplierConfigItem[]
}>

type SupplierWorkbookData = Awaited<ReturnType<typeof readSupplierWorkbookFromFile>>

const SUPPLIER_CONFIG: SupplierConfigRoot = supplierConfigJson as SupplierConfigRoot
const CENTIMETER_SQUARE_TO_METER_SQUARE = 10_000
const DEFAULT_DIMENSION_ADJUSTMENT_INPUT = "0"
const EXCEL_FILE_ACCEPT = ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
const EXCEL_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
] as const
const EXCEL_FILE_EXTENSIONS = [".xlsx", ".xls"] as const

function calculateSquareMeter(length: number | null, width: number | null): number | null {
  if (length === null || width === null) {
    return null
  }
  return Math.round(((length * width) / CENTIMETER_SQUARE_TO_METER_SQUARE) * 100) / 100
}

function calculateTotalSquareMeter(
  rows: SupplierWorkbookData["rows"],
  getLength: (row: SupplierWorkbookData["rows"][number]) => number | null,
  getWidth: (row: SupplierWorkbookData["rows"][number]) => number | null,
): number {
  const totalValue = rows.reduce((sum: number, row: SupplierWorkbookData["rows"][number]): number => {
    const squareMeter = calculateSquareMeter(getLength(row), getWidth(row))
    return sum + (squareMeter ?? 0)
  }, 0)
  return Math.round(totalValue * 100) / 100
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

export default function CreateListPage(): ReactElement {
  const [selectedSupplierName, setSelectedSupplierName] = useState<string>(SUPPLIER_CONFIG.suppliers[0]?.name ?? "")
  const [isReadingFile, setIsReadingFile] = useState<boolean>(false)
  const [isCreatingList, setIsCreatingList] = useState<boolean>(false)
  const [errorMessage, setErrorMessage] = useState<string>("")
  const [workbookData, setWorkbookData] = useState<SupplierWorkbookData | null>(null)
  const [dimensionAdjustmentInput, setDimensionAdjustmentInput] = useState<string>(
    DEFAULT_DIMENSION_ADJUSTMENT_INPUT,
  )
  const supplierNames = useMemo(
    (): readonly string[] => SUPPLIER_CONFIG.suppliers.map((supplier: SupplierConfigItem): string => supplier.name),
    [],
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
    const adjustedRows = workbookData.rows.map((row: SupplierWorkbookData["rows"][number]) => ({
      ...row,
      lengthGross: getAdjustedDimension(row.lengthGross, dimensionAdjustmentInCentimeter),
      widthGross: getAdjustedDimension(row.widthGross, dimensionAdjustmentInCentimeter),
      lengthNet: getAdjustedDimension(row.lengthNet, dimensionAdjustmentInCentimeter),
      widthNet: getAdjustedDimension(row.widthNet, dimensionAdjustmentInCentimeter),
    }))
    return {
      ...workbookData,
      rows: adjustedRows,
    }
  }, [dimensionAdjustmentInCentimeter, workbookData])
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
    setWorkbookData(null)
    setDimensionAdjustmentInput(DEFAULT_DIMENSION_ADJUSTMENT_INPUT)
    try {
      const parsedWorkbookData = await readSupplierWorkbookFromFile({ file: inputFile, supplierName: selectedSupplierName })
      setWorkbookData(parsedWorkbookData)
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể đọc file Excel."
      setErrorMessage(message)
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
      await exportSupplierListFile(adjustedWorkbookData)
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
          Chọn nhà cung cấp, tải file Excel và hệ thống sẽ đọc dữ liệu theo cấu hình tương ứng.
        </p>
      </div>
      <section className="space-y-4 rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
        <div className="grid gap-4 md:grid-cols-2">
          <label className="space-y-2">
            <span className="text-sm font-medium">Nhà cung cấp</span>
            <select
              className="w-full rounded-md border px-3 py-2 text-sm"
              style={{ borderColor: "var(--input)", backgroundColor: "var(--background)" }}
              value={selectedSupplierName}
              onChange={(event: ChangeEvent<HTMLSelectElement>): void => setSelectedSupplierName(event.target.value)}
            >
              {supplierNames.map((supplierName: string): ReactElement => (
                <option key={supplierName} value={supplierName}>
                  {supplierName}
                </option>
              ))}
            </select>
          </label>
          <label className="space-y-2">
            <span className="text-sm font-medium">File Excel</span>
            <input
              className="w-full rounded-md border px-3 py-2 text-sm file:mr-4 file:rounded-md file:border-0 file:bg-slate-900 file:px-3 file:py-2 file:text-sm file:text-white"
              style={{ borderColor: "var(--input)" }}
              type="file"
              accept={EXCEL_FILE_ACCEPT}
              onChange={handleUploadFile}
              disabled={isReadingFile}
            />
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
        <div className="flex justify-end">
          <button
            className="w-full rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50 sm:w-auto"
            type="button"
            onClick={handleCreateList}
            disabled={!workbookData || isCreatingList || isReadingFile}
          >
            {isCreatingList ? "Đang tạo list..." : "Tạo List"}
          </button>
        </div>
      </section>
      {adjustedWorkbookData ? (
        <section className="space-y-4">
          <div className="rounded-xl border p-3 sm:p-4" style={{ borderColor: "var(--border)" }}>
            <h2 className="text-base font-semibold">Thông tin chung</h2>
            <div className="mt-3 grid gap-3 text-sm md:grid-cols-2">
              <div>
                <span className="font-medium">Nhà cung cấp: </span>
                <span>{adjustedWorkbookData.supplierName || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Container Number: </span>
                <span>{adjustedWorkbookData.generalInfo.containerNumber || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Material: </span>
                <span>{adjustedWorkbookData.generalInfo.materialName || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Type of polish: </span>
                <span>{adjustedWorkbookData.generalInfo.typeOfPolish || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Số lượng slabs: </span>
                <span>{adjustedWorkbookData.rows.length || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Loading date: </span>
                <span>{adjustedWorkbookData.generalInfo.loadingDate || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Invoice Number: </span>
                <span>{adjustedWorkbookData.generalInfo.invoiceNumber || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Invoice Date: </span>
                <span>{adjustedWorkbookData.generalInfo.invoiceDate || "XXX"}</span>
              </div>
            </div>
          </div>
          <div className="rounded-xl border" style={{ borderColor: "var(--border)" }}>
            <div className="space-y-3 p-3 md:hidden">
              {adjustedWorkbookData.rows.map((row): ReactElement => (
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
                  </div>
                </div>
              ))}
              <div className="rounded-lg border p-3 text-sm font-semibold" style={{ borderColor: "var(--border)" }}>
                <div className="flex items-center justify-between">
                  <span>Tổng GROSS SQM</span>
                  <span>{calculateTotalSquareMeter(adjustedWorkbookData.rows, (row) => row.lengthGross, (row) => row.widthGross)}</span>
                </div>
                <div className="mt-2 flex items-center justify-between">
                  <span>Tổng NET SQM</span>
                  <span>{calculateTotalSquareMeter(adjustedWorkbookData.rows, (row) => row.lengthNet, (row) => row.widthNet)}</span>
                </div>
              </div>
            </div>
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
                  </tr>
                </thead>
                <tbody>
                  {adjustedWorkbookData.rows.map((row): ReactElement => (
                    <tr key={`${row.rowNumber}-${row.slabNoGross}-${row.slabNoNet}`} className="border-t" style={{ borderColor: "var(--border)" }}>
                      <td className="px-3 py-2">{row.slabNoGross || "-"}</td>
                      <td className="px-3 py-2">{row.lengthGross ?? "-"}</td>
                      <td className="px-3 py-2">{row.widthGross ?? "-"}</td>
                      <td className="px-3 py-2">{calculateSquareMeter(row.lengthGross, row.widthGross) ?? "-"}</td>
                      <td className="px-3 py-2">{row.slabNoNet || "-"}</td>
                      <td className="px-3 py-2">{row.lengthNet ?? "-"}</td>
                      <td className="px-3 py-2">{row.widthNet ?? "-"}</td>
                      <td className="px-3 py-2">{calculateSquareMeter(row.lengthNet, row.widthNet) ?? "-"}</td>
                    </tr>
                  ))}
                  <tr className="border-t font-semibold" style={{ borderColor: "var(--border)" }}>
                    <td className="px-3 py-2">Total</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">
                      {calculateTotalSquareMeter(adjustedWorkbookData.rows, (row) => row.lengthGross, (row) => row.widthGross)}
                    </td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">
                      {calculateTotalSquareMeter(adjustedWorkbookData.rows, (row) => row.lengthNet, (row) => row.widthNet)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="border-t px-3 py-2 text-sm" style={{ borderColor: "var(--border)" }}>
              Tổng số dòng đọc được: {adjustedWorkbookData.rows.length}
            </div>
          </div>
        </section>
      ) : null}
    </div>
  )
}
