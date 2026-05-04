import { LayoutGrid, FileSpreadsheet } from "lucide-react"
import type { ComponentType, ReactElement } from "react"
import { useMemo } from "react"
import { NavLink, Outlet, useLocation } from "react-router-dom"

type SidebarNavKey = "create-list" | "menu-b"

type SidebarNavItem = Readonly<{
  key: SidebarNavKey
  label: string
  icon: ComponentType<{ size?: number }>
  to: string
}>

const SIDEBAR_NAV_ITEMS: SidebarNavItem[] = [
  { key: "create-list", label: "Tạo List", icon: FileSpreadsheet, to: "/admin/create-list" },
  { key: "menu-b", label: "Menu B", icon: LayoutGrid, to: "/admin/menu-b" }
]

/**
 * Dashboard quản trị đơn giản: sidebar trái + nội dung bên phải.
 * Không yêu cầu đăng nhập.
 */
export default function AdminDashboard(): ReactElement {
  const location = useLocation()

  const activeTitle = useMemo((): string => {
    const normalizedPath = location.pathname.endsWith("/") ? location.pathname.slice(0, -1) : location.pathname
    const activeItem = SIDEBAR_NAV_ITEMS.find((item: SidebarNavItem): boolean => item.to === normalizedPath)
    return activeItem?.label ?? "Admin"
  }, [location.pathname])

  return (
    <div className="min-h-screen flex" style={{ backgroundColor: "var(--background)", color: "var(--foreground)" }}>
      <aside
        className="w-72 shrink-0 border-r"
        style={{ borderColor: "var(--sidebar-border)", backgroundColor: "var(--sidebar)", color: "var(--sidebar-foreground)" }}
      >
        <div className="h-16 flex items-center px-4" style={{ borderBottom: "1px solid var(--sidebar-border)" }}>
          <div className="font-semibold">Thiên Phúc Stone</div>
        </div>

        <nav className="p-2 space-y-1" aria-label="Quản lý">
          {SIDEBAR_NAV_ITEMS.map((item: SidebarNavItem): ReactElement => (
            <NavLink
              key={item.key}
              to={item.to}
              end
              className="w-full flex items-center gap-3 rounded-lg px-3 py-2 text-sm text-left transition-colors focus:outline-none"
              style={({ isActive }): { backgroundColor: string; color: string } =>
                isActive
                  ? { backgroundColor: "var(--sidebar-accent)", color: "var(--sidebar-accent-foreground)" }
                  : { backgroundColor: "transparent", color: "var(--sidebar-foreground)" }
              }
            >
              <item.icon size={18} />
              <span>{item.label}</span>
            </NavLink>
          ))}
        </nav>
      </aside>

      <div className="flex-1 flex flex-col">
        <header
          className="h-16 flex items-center justify-between px-6 border-b"
          style={{ borderBottomColor: "var(--border)" }}
        >
          <h1 className="text-lg font-semibold">{activeTitle}</h1>
          <div className="text-sm" style={{ color: "var(--muted-foreground)" }}>
            Dashboard đơn giản
          </div>
        </header>

        <main className="flex-1 p-6">
          <Outlet />
        </main>
      </div>
    </div>
  )
}

