import supplierConfigJson from "@/config/supplier-config.json"
import { readSupplierWorkbookFromFile } from "@/lib/supplier-workbook-reader"
import { exportSupplierListFile } from "@/lib/template-list-exporter"
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

export default function CreateListPage(): ReactElement {
  const [selectedSupplierName, setSelectedSupplierName] = useState<string>(SUPPLIER_CONFIG.suppliers[0]?.name ?? "")
  const [isReadingFile, setIsReadingFile] = useState<boolean>(false)
  const [isCreatingList, setIsCreatingList] = useState<boolean>(false)
  const [errorMessage, setErrorMessage] = useState<string>("")
  const [workbookData, setWorkbookData] = useState<SupplierWorkbookData | null>(null)
  const supplierNames = useMemo(
    (): readonly string[] => SUPPLIER_CONFIG.suppliers.map((supplier: SupplierConfigItem): string => supplier.name),
    [],
  )
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
    setIsReadingFile(true)
    setErrorMessage("")
    setWorkbookData(null)
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
    if (!workbookData) {
      setErrorMessage("Vui lòng tải dữ liệu trước khi tạo list.")
      return
    }
    setErrorMessage("")
    setIsCreatingList(true)
    try {
      await exportSupplierListFile(workbookData)
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : "Không thể tạo file list."
      setErrorMessage(message)
    } finally {
      setIsCreatingList(false)
    }
  }
  return (
    <div className="max-w-6xl space-y-6">
      <div className="space-y-2">
        <h1 className="text-2xl font-semibold">Tạo List</h1>
        <p className="text-sm" style={{ color: "var(--muted-foreground)" }}>
          Chọn nhà cung cấp, tải file Excel và hệ thống sẽ đọc dữ liệu theo cấu hình tương ứng.
        </p>
      </div>
      <section className="rounded-xl border p-4 space-y-4" style={{ borderColor: "var(--border)" }}>
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
              accept=".xlsx,.xls"
              onChange={handleUploadFile}
              disabled={isReadingFile}
            />
          </label>
        </div>
        {isReadingFile ? <p className="text-sm">Đang đọc dữ liệu từ file Excel...</p> : null}
        {errorMessage ? (
          <p className="rounded-md border border-red-400 bg-red-50 px-3 py-2 text-sm text-red-700">{errorMessage}</p>
        ) : null}
        <div className="flex justify-end">
          <button
            className="rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
            type="button"
            onClick={handleCreateList}
            disabled={!workbookData || isCreatingList || isReadingFile}
          >
            {isCreatingList ? "Đang tạo list..." : "Tạo List"}
          </button>
        </div>
      </section>
      {workbookData ? (
        <section className="space-y-4">
          <div className="rounded-xl border p-4" style={{ borderColor: "var(--border)" }}>
            <h2 className="text-base font-semibold">Thông tin chung</h2>
            <div className="mt-3 grid gap-3 text-sm md:grid-cols-2">
              <div>
                <span className="font-medium">Nhà cung cấp: </span>
                <span>{workbookData.supplierName || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Container Number: </span>
                <span>{workbookData.generalInfo.containerNumber || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Material: </span>
                <span>{workbookData.generalInfo.materialName || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Type of polish: </span>
                <span>{workbookData.generalInfo.typeOfPolish || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Số lượng slabs: </span>
                <span>{workbookData.rows.length || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Loading date: </span>
                <span>{workbookData.generalInfo.loadingDate || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Invoice Number: </span>
                <span>{workbookData.generalInfo.invoiceNumber || "XXX"}</span>
              </div>
              <div>
                <span className="font-medium">Invoice Date: </span>
                <span>{workbookData.generalInfo.invoiceDate || "XXX"}</span>
              </div>
            </div>
          </div>
          <div className="rounded-xl border" style={{ borderColor: "var(--border)" }}>
            <div className="overflow-x-auto">
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
                  {workbookData.rows.map((row): ReactElement => (
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
                      {calculateTotalSquareMeter(workbookData.rows, (row) => row.lengthGross, (row) => row.widthGross)}
                    </td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">-</td>
                    <td className="px-3 py-2">
                      {calculateTotalSquareMeter(workbookData.rows, (row) => row.lengthNet, (row) => row.widthNet)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="border-t px-3 py-2 text-sm" style={{ borderColor: "var(--border)" }}>
              Tổng số dòng đọc được: {workbookData.rows.length}
            </div>
          </div>
        </section>
      ) : null}
    </div>
  )
}
