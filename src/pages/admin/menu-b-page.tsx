import type { ReactElement } from "react"

export default function MenuBPage(): ReactElement {
  return (
    <div className="max-w-2xl">
      <div
        className="rounded-lg border p-6"
        style={{ borderColor: "var(--border)", backgroundColor: "var(--card)" }}
      >
        <h2 className="text-lg font-semibold">Menu B</h2>
        <p className="text-sm mt-2" style={{ color: "var(--muted-foreground)" }}>
          Placeholder: nội dung trang Menu B.
        </p>
      </div>
    </div>
  )
}

