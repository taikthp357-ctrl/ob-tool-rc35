# -*- coding: utf-8 -*-
"""MA CONTROL FG — Standalone Native Desktop GUI (Zero-Port, Dell 8GB RAM Compliant).

Provides native Desktop GUI with:
- 100% Zero-Port (No open sockets or network listeners)
- 5-step stepper bar & 2 Branch Tabs (Body 3001 / EDC 4007)
- Central Data Grid with material editing and status chips
- Enforced Manual CF Gate: Pick Order starts in WAITING_USER_CF, blocks BBBG until confirmed
- Fast in-process synchronization with Edge App Mode (data.js)
"""

from __future__ import annotations

import datetime
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
from typing import Any, Dict, List, Optional

APP_ROOT = Path(__file__).resolve().parent.parent.parent
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from app.core.config import APP_TITLE, APP_VERSION, DB_PATH
from app.core.database import get_connection, init_db
from app.services.ob_service import OBService
from app.services.pick_cf_service import PickCFService
from app.services.bbbg_service import BBBGService
from app.web_bridge import sync_web_dashboard, ZeroPortActionDispatcher


class MAControlMainWindow:
    """Main Desktop Window for MA CONTROL FG."""

    def __init__(self, root: Optional[tk.Tk] = None, db_path: Optional[str] = None):
        self.root = root or tk.Tk()
        self.db_path = db_path
        self.root.title(f"{APP_TITLE} ({APP_VERSION}) — Zero-Port Standalone")
        self.root.geometry("1120x740")
        self.root.minsize(960, 600)
        self.root.configure(bg="#0f172a")

        self.current_branch = "BODY_3001"
        self.ob_service = OBService()
        self.pick_service = PickCFService()
        self.bbbg_service = BBBGService()

        init_db(self.db_path)

        self._configure_styles()
        self._build_header()
        self._build_stepper()
        self._build_body()
        self._build_pick_section()
        self._build_statusbar()

        self._load_data()

    def _configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        # Tab styling
        style.configure("TNotebook", background="#0f172a", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1e293b", foreground="#94a3b8", padding=[16, 6], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "#2563eb")], foreground=[("selected", "#ffffff")])

        # Treeview styling
        style.configure("Treeview", background="#ffffff", foreground="#0f172a", rowheight=26, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#f1f5f9", foreground="#334155", font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", "#1e40af")])

    def _build_header(self):
        header_frame = tk.Frame(self.root, bg="#1e293b", height=50)
        header_frame.pack(fill=tk.X, padx=0, pady=0)

        # Title
        lbl_title = tk.Label(header_frame, text="MA CONTROL — Điều hành FG", font=("Segoe UI", 13, "bold"), fg="#ffffff", bg="#1e293b")
        lbl_title.pack(side=tk.LEFT, padx=16, pady=10)

        # Status Badges
        lbl_sap = tk.Label(header_frame, text="● SAP đã kết nối", font=("Segoe UI", 9, "bold"), fg="#10b981", bg="#064e3b", padx=8, pady=2)
        lbl_sap.pack(side=tk.LEFT, padx=6)

        lbl_zero = tk.Label(header_frame, text="🔒 Zero-Port: ENABLED", font=("Segoe UI", 9, "bold"), fg="#38bdf8", bg="#0c4a6e", padx=8, pady=2)
        lbl_zero.pack(side=tk.LEFT, padx=6)

        lbl_ver = tk.Label(header_frame, text=f"{APP_VERSION}", font=("Segoe UI", 9), fg="#94a3b8", bg="#1e293b")
        lbl_ver.pack(side=tk.LEFT, padx=6)

        # User info
        lbl_user = tk.Label(header_frame, text="👤 GL_BODY_01 • Phân xưởng Thân vỏ", font=("Segoe UI", 9), fg="#cbd5e1", bg="#1e293b")
        lbl_user.pack(side=tk.RIGHT, padx=16)

    def _build_stepper(self):
        step_frame = tk.Frame(self.root, bg="#0f172a", height=38)
        step_frame.pack(fill=tk.X, padx=16, pady=8)

        steps = [
            ("1. Tải tồn", "#10b981"),
            ("➔", "#64748b"),
            ("2. Điền OB", "#38bdf8"),
            ("➔", "#64748b"),
            ("3. Ra đơn pick", "#94a3b8"),
            ("➔", "#64748b"),
            ("4. Người dùng CF", "#f59e0b"),
            ("➔", "#64748b"),
            ("5. BBBG", "#94a3b8"),
        ]
        for txt, color in steps:
            lbl = tk.Label(step_frame, text=txt, font=("Segoe UI", 9, "bold" if "➔" not in txt else "normal"), fg=color, bg="#0f172a")
            lbl.pack(side=tk.LEFT, padx=4)

    def _build_body(self):
        # Toolbar & Tabs
        container = tk.Frame(self.root, bg="#0f172a")
        container.pack(fill=tk.BOTH, expand=True, padx=16, pady=4)

        # Tabs Notebook
        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_body = tk.Frame(self.notebook, bg="#1e293b")
        self.tab_edc = tk.Frame(self.notebook, bg="#1e293b")

        self.notebook.add(self.tab_body, text=" Body • 3001 ")
        self.notebook.add(self.tab_edc, text=" EDC • 4007 ")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Build Grid in Tab Body
        self.tree_body = self._create_treeview(self.tab_body)
        self.tree_edc = self._create_treeview(self.tab_edc)

        # Action Buttons bar
        btn_bar = tk.Frame(container, bg="#0f172a")
        btn_bar.pack(fill=tk.X, pady=8)

        btn_check = tk.Button(btn_bar, text="🔍 Kiểm tra OB", font=("Segoe UI", 9, "bold"), bg="#334155", fg="#ffffff", padx=12, pady=4, relief=tk.FLAT, command=self._action_check_ob)
        btn_check.pack(side=tk.LEFT, padx=4)

        btn_fill = tk.Button(btn_bar, text="⚡ Điền OB (RC52)", font=("Segoe UI", 9, "bold"), bg="#2563eb", fg="#ffffff", padx=12, pady=4, relief=tk.FLAT, command=self._action_fill_ob)
        btn_fill.pack(side=tk.LEFT, padx=4)

        btn_pick = tk.Button(btn_bar, text="📦 Ra đơn pick", font=("Segoe UI", 9, "bold"), bg="#107c41", fg="#ffffff", padx=12, pady=4, relief=tk.FLAT, command=self._action_create_pick)
        btn_pick.pack(side=tk.LEFT, padx=4)

        btn_sync = tk.Button(btn_bar, text="🔄 Đồng bộ sang Edge App", font=("Segoe UI", 9), bg="#475569", fg="#ffffff", padx=12, pady=4, relief=tk.FLAT, command=self._action_sync_edge)
        btn_sync.pack(side=tk.RIGHT, padx=4)

    def _create_treeview(self, parent: tk.Frame) -> ttk.Treeview:
        cols = ("model", "material_code", "required_qty", "batch", "shop", "sto_number", "ob_number", "status")
        tree = ttk.Treeview(parent, columns=cols, show="headings", selectmode="extended")

        tree.heading("model", text="Model")
        tree.heading("material_code", text="Mã hàng")
        tree.heading("required_qty", text="SL yêu cầu")
        tree.heading("batch", text="Batch")
        tree.heading("shop", text="Shop")
        tree.heading("sto_number", text="STO")
        tree.heading("ob_number", text="OB")
        tree.heading("status", text="Trạng thái")

        tree.column("model", width=90, anchor=tk.CENTER)
        tree.column("material_code", width=140, anchor=tk.W)
        tree.column("required_qty", width=90, anchor=tk.E)
        tree.column("batch", width=90, anchor=tk.CENTER)
        tree.column("shop", width=80, anchor=tk.CENTER)
        tree.column("sto_number", width=110, anchor=tk.CENTER)
        tree.column("ob_number", width=110, anchor=tk.CENTER)
        tree.column("status", width=130, anchor=tk.CENTER)

        tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        return tree

    def _build_pick_section(self):
        pick_frame = tk.LabelFrame(self.root, text=" 📦 Đơn pick đã tạo (Chờ người dùng CF thủ công) ", font=("Segoe UI", 9, "bold"), fg="#38bdf8", bg="#1e293b", padx=8, pady=6)
        pick_frame.pack(fill=tk.X, padx=16, pady=4)

        cols = ("pick_order_no", "created_date", "shop", "model", "total_items", "total_qty", "status", "action")
        self.tree_pick = ttk.Treeview(pick_frame, columns=cols, show="headings", height=4)

        self.tree_pick.heading("pick_order_no", text="Số đơn pick")
        self.tree_pick.heading("created_date", text="Ngày tạo")
        self.tree_pick.heading("shop", text="Shop")
        self.tree_pick.heading("model", text="Model")
        self.tree_pick.heading("total_items", text="Số mã hàng")
        self.tree_pick.heading("total_qty", text="Tổng SL (PCS)")
        self.tree_pick.heading("status", text="Trạng thái")
        self.tree_pick.heading("action", text="Thao tác")

        self.tree_pick.column("pick_order_no", width=160, anchor=tk.W)
        self.tree_pick.column("created_date", width=100, anchor=tk.CENTER)
        self.tree_pick.column("shop", width=80, anchor=tk.CENTER)
        self.tree_pick.column("model", width=100, anchor=tk.CENTER)
        self.tree_pick.column("total_items", width=90, anchor=tk.CENTER)
        self.tree_pick.column("total_qty", width=100, anchor=tk.E)
        self.tree_pick.column("status", width=150, anchor=tk.CENTER)
        self.tree_pick.column("action", width=110, anchor=tk.CENTER)

        self.tree_pick.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Action buttons on right
        action_btn_box = tk.Frame(pick_frame, bg="#1e293b")
        action_btn_box.pack(side=tk.RIGHT, padx=8, fill=tk.Y)

        btn_cf = tk.Button(action_btn_box, text="✅ Kiểm tra CF", font=("Segoe UI", 9, "bold"), bg="#f59e0b", fg="#ffffff", padx=8, pady=4, relief=tk.FLAT, command=self._action_manual_cf)
        btn_cf.pack(fill=tk.X, pady=2)

        btn_bbbg = tk.Button(action_btn_box, text="📄 Tạo BBBG", font=("Segoe UI", 9, "bold"), bg="#107c41", fg="#ffffff", padx=8, pady=4, relief=tk.FLAT, command=self._action_generate_bbbg)
        btn_bbbg.pack(fill=tk.X, pady=2)

    def _build_statusbar(self):
        status_bar = tk.Frame(self.root, bg="#0f172a", height=24)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.lbl_status = tk.Label(status_bar, text="Sẵn sàng • Chế độ Zero-Port 100% không mở cổng mạng • Tuân thủ chính sách IT công ty", font=("Segoe UI", 8), fg="#64748b", bg="#0f172a")
        self.lbl_status.pack(side=tk.LEFT, padx=16, pady=4)

    def _on_tab_changed(self, event):
        idx = self.notebook.index(self.notebook.select())
        self.current_branch = "BODY_3001" if idx == 0 else "EDC_4007"
        self._load_data()

    def _load_data(self):
        conn = get_connection(self.db_path)
        try:
            # Load active branch items
            active_tree = self.tree_body if self.current_branch == "BODY_3001" else self.tree_edc
            for row in active_tree.get_children():
                active_tree.delete(row)

            rows = conn.execute("SELECT * FROM ob_items WHERE branch = ? ORDER BY id ASC", (self.current_branch,)).fetchall()
            for r in rows:
                active_tree.insert("", tk.END, values=(
                    r["model"], r["material_code"], r["required_qty"], r["batch"], r["shop"], r["sto_number"], r["ob_number"], r["status"]
                ))

            # Load Pick orders
            for row in self.tree_pick.get_children():
                self.tree_pick.delete(row)

            p_rows = conn.execute("SELECT * FROM pick_orders ORDER BY created_at DESC LIMIT 10").fetchall()
            for p in p_rows:
                st = p["status"]
                st_label = "Chờ người dùng CF" if st in ("WAITING_USER_CF", "Chờ CF") else st
                self.tree_pick.insert("", tk.END, values=(
                    p.get("pick_order_code") or p.get("pick_order_no"),
                    p.get("created_date"),
                    p.get("shop"),
                    p.get("model"),
                    p.get("total_items") or len(p_rows),
                    p.get("total_qty"),
                    st_label,
                    "Kiểm tra CF"
                ))
        finally:
            conn.close()

    def _action_check_ob(self):
        res = ZeroPortActionDispatcher.check_ob_items(self.current_branch, self.db_path)
        self._load_data()
        messagebox.showinfo("Kiểm tra OB", f"Đã kiểm tra xong các dòng OB nhánh {self.current_branch}!\nĐã cập nhật trạng thái sẵn sàng xuất.")

    def _action_fill_ob(self):
        res = ZeroPortActionDispatcher.fill_ob_items(self.current_branch, self.db_path)
        self._load_data()
        messagebox.showinfo("Điền OB RC52", f"Đã điền xong OB theo chuẩn RC52 cho nhánh {self.current_branch}!\nĐã sẵn sàng chuyển sang bước 3: Ra đơn pick.")

    def _action_create_pick(self):
        conn = get_connection(self.db_path)
        try:
            items = conn.execute("SELECT * FROM ob_items WHERE branch = ? AND (status IN ('READY_FOR_PICK', 'Sẵn sàng điền', 'WAITING_PICK') OR ob_number != '')", (self.current_branch,)).fetchall()
            dict_items = [dict(it) for it in items]
            if not dict_items:
                messagebox.showwarning("Ra đơn pick", "Không có dòng OB nào đủ điều kiện tạo đơn pick!")
                return

            shop = "BODY" if "BODY" in self.current_branch else "EDC"
            model = dict_items[0].get("model") or "VF8"
            res = ZeroPortActionDispatcher.create_pick_order(shop, model, dict_items, self.current_branch, self.db_path)
            self._load_data()
            messagebox.showinfo(
                "Ra đơn pick thành công",
                f"Đã tạo đơn pick: {res.get('pick_order_code') or res.get('pick_order_no')}\n"
                f"Trạng thái BẮT BUỘC: 'Chờ người dùng CF'.\n\n"
                f"⚠️ Lưu ý: Người dùng phải kiểm tra thực tế và nhấn 'Kiểm tra CF' trước khi sang bước BBBG!"
            )
        finally:
            conn.close()

    def _action_manual_cf(self):
        sel = self.tree_pick.selection()
        if not sel:
            messagebox.showwarning("Xác nhận CF", "Vui lòng chọn 1 đơn pick trong danh sách để kiểm tra CF!")
            return
        vals = self.tree_pick.item(sel[0])["values"]
        pick_no = vals[0]
        st = vals[6]

        if "Đã CF" in st or "USER_CONFIRMED" in st:
            messagebox.showinfo("Xác nhận CF", f"Đơn pick {pick_no} đã được xác nhận CF trước đó!")
            return

        confirm = messagebox.askyesno(
            "Cổng kiểm tra CF thủ công (Bước 4)",
            f"Bạn có xác nhận đã kiểm tra THỰC TẾ hàng hóa tại bãi cho đơn pick:\n{pick_no}?\n\n"
            f"• Khớp mã hàng, số lượng và lô (batch).\n"
            f"• Đủ điều kiện chuyển giao sang Bước 5 (BBBG).\n\n"
            f"Nhấn YES để xác nhận CF."
        )
        if confirm:
            ZeroPortActionDispatcher.confirm_pick_order(pick_no, "GL_BODY_01", db_path=self.db_path)
            self._load_data()
            messagebox.showinfo("CF thành công", f"Đã xác nhận CF thành công cho đơn {pick_no}!\nĐơn đã đủ điều kiện để tạo BBBG.")

    def _action_generate_bbbg(self):
        sel = self.tree_pick.selection()
        if not sel:
            messagebox.showwarning("Tạo BBBG", "Vui lòng chọn 1 đơn pick đã CF thành công để tạo BBBG!")
            return
        vals = self.tree_pick.item(sel[0])["values"]
        pick_no = vals[0]
        st = vals[6]

        try:
            res = ZeroPortActionDispatcher.generate_bbbg(pick_no, "GL_BODY_01", db_path=self.db_path)
            self._load_data()
            messagebox.showinfo("Tạo BBBG", f"Tạo BBBG thành công theo chuẩn V2.2.14!\nSố phiếu: {res.get('bbbg_no')}\nĐã lưu sổ tổng và chuẩn bị đính kèm GOS.")
        except Exception as ex:
            messagebox.showerror("Bảo vệ Bất biến Nghiệp vụ", str(ex))

    def _action_sync_edge(self):
        res = sync_web_dashboard(self.db_path)
        messagebox.showinfo("Đồng bộ", f"Đã đồng bộ toàn bộ dữ liệu ({len(res['pick_orders'])} đơn pick, {len(res['ob_items_body_3001'])} mã Body) sang data.js cho Edge App Mode!")


def launch_gui(db_path: Optional[str] = None):
    root = tk.Tk()
    app = MAControlMainWindow(root, db_path=db_path)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
