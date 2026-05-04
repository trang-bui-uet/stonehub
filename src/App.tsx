import AdminDashboard from "./components/admin-dashboard/admin-dashboard"
import CreateListPage from "./pages/admin/create-list-page"
import MenuBPage from "./pages/admin/menu-b-page"

import type { ReactElement } from "react"
import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom"

function App(): ReactElement {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/admin" element={<AdminDashboard />}>
          <Route path="create-list" element={<CreateListPage />} />
          <Route path="menu-b" element={<MenuBPage />} />
          <Route index element={<Navigate to="create-list" replace />} />
        </Route>

        <Route path="*" element={<Navigate to="/admin" replace />} />
      </Routes>
    </BrowserRouter>
  )
}

export default App
