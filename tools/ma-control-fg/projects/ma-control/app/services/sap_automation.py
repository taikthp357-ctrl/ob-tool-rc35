# -*- coding: utf-8 -*-
"""SAP GUI COM Automation Service.

Implements native SAP Scripting COM automation for:
1. MB52: Direct inventory download & snapshot sync.
2. MB51: Direct material document history extraction for BBBG reconciliation.
3. VL06I: Direct inbound delivery item extraction for BBBG reconciliation.

Conventions & Baseline references:
- MB51 baseline: Plant 3093, Layout /khoma, XBLNR multiple selection (mb51.vbs, Mod_SAP_Integration.bas).
- VL06I baseline: BUTTON7, clear date, IT_LIFEX multiple selection, IT_WBSTK='a', btn[18] item view (vl06i.vbs).
- MB52 baseline: Layout /KHOMA, plants 3001/3093 (Dowton V3.7.0 Module1.bas).
- Single-thread STA COM executor with pythoncom.CoInitialize/CoUninitialize.
- Intermediate Excel export auto-cleanup upon validation and diagnostic retention upon failure.
"""

import os
import sys
import uuid
import time
import tempfile
import threading
from pathlib import Path
from typing import Dict, Any, List, Optional
from concurrent.futures import ThreadPoolExecutor

import openpyxl

# Thread-isolated COM executor for SAP GUI Scripting
_com_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="ma-sap-com")
_com_lock = threading.Lock()


def set_windows_clipboard(text: str) -> bool:
    """Set Windows clipboard text using win32clipboard or ctypes fallback."""
    try:
        import win32clipboard
        import win32con
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
            return True
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass

    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        if not user32.OpenClipboard(0):
            return False
        try:
            user32.EmptyClipboard()
            encoded = text.encode("utf-16le") + b"\x00\x00"
            h_mem = kernel32.GlobalAlloc(0x0002, len(encoded))  # GMEM_MOVEABLE
            if not h_mem:
                return False
            p_mem = kernel32.GlobalLock(h_mem)
            if not p_mem:
                return False
            ctypes.memmove(p_mem, encoded, len(encoded))
            kernel32.GlobalUnlock(h_mem)
            # 13 is CF_UNICODETEXT
            return bool(user32.SetClipboardData(13, h_mem))
        finally:
            user32.CloseClipboard()
    except Exception:
        return False


def get_sap_com_session(connection_index: int = 0, session_index: int = 0):
    """Retrieve an active SAP GUI scripting session COM object.
    
    Must be called from a COM-initialized thread.
    """
    import win32com.client
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        raise RuntimeError("Không tìm thấy đối tượng COM 'SAPGUI'. Hãy đảm bảo SAP Logon đang chạy.") from exc

    try:
        engine = sap_gui.GetScriptingEngine
    except Exception as exc:
        raise RuntimeError("Không lấy được Scripting Engine từ SAPGUI. Vui lòng kiểm tra quyền SAP GUI Scripting.") from exc

    if not engine or engine.Children.Count <= connection_index:
        raise RuntimeError(f"Chưa có kết nối SAP GUI nào (connection index {connection_index}). Hãy đăng nhập SAP.")

    connection = engine.Children(connection_index)
    if not connection or connection.Children.Count <= session_index:
        raise RuntimeError(f"Chưa có phiên làm việc SAP GUI nào (session index {session_index}).")

    session = connection.Children(session_index)
    if bool(session.Busy):
        raise RuntimeError("Phiên SAP GUI đang bận xử lý tác vụ khác. Vui lòng thử lại sau giây lát.")

    return session


def write_excel_dataset(headers: List[str], rows: List[List[Any]], prefix: str = "SAP_") -> Path:
    """Create an intermediate Excel file with headers and rows."""
    temp_dir = Path(tempfile.gettempdir())
    out_path = temp_dir / f"{prefix}{uuid.uuid4().hex[:8]}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(str(out_path))
    return out_path


# ---------------------------------------------------------------------------
# 1. MB52 Direct COM Automation
# ---------------------------------------------------------------------------
def _execute_mb52_com(
    plants: Optional[List[str]] = None,
    sloc: Optional[str] = None,
    layout: str = "/KHOMA",
    materials: Optional[List[str]] = None,
    conn_idx: int = 0,
    sess_idx: int = 0
) -> Path:
    """Worker implementation of MB52 export over SAP GUI COM."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        session = get_sap_com_session(conn_idx, sess_idx)
        wnd0 = session.findById("wnd[0]")
        wnd0.maximize()

        # Open MB52 transaction
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
        wnd0.sendVKey(0)

        # Set Layout / Variant (e.g. /KHOMA)
        try:
            session.findById("wnd[0]/usr/ctxtP_VARI").Text = layout
        except Exception:
            pass

        # Plant selection
        target_plants = plants or ["3001", "3093"]
        if len(target_plants) == 1:
            try:
                session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = target_plants[0]
            except Exception:
                pass
        elif len(target_plants) > 1:
            try:
                session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").Press()
                # Clear table button
                try:
                    session.findById("wnd[1]/tbar[0]/btn[16]").Press()
                except Exception:
                    pass
                set_windows_clipboard("\r\n".join(target_plants))
                # Paste from clipboard button
                try:
                    session.findById("wnd[1]/tbar[0]/btn[24]").Press()
                except Exception:
                    pass
                session.findById("wnd[1]").sendVKey(8)  # F8 confirm popup
            except Exception:
                pass

        # SLoc filter if provided
        if sloc:
            try:
                session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = sloc
            except Exception:
                pass

        # Materials filter if provided
        if materials:
            try:
                session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").Press()
                try:
                    session.findById("wnd[1]/tbar[0]/btn[16]").Press()
                except Exception:
                    pass
                set_windows_clipboard("\r\n".join(materials))
                session.findById("wnd[1]/tbar[0]/btn[24]").Press()
                session.findById("wnd[1]").sendVKey(8)
            except Exception:
                pass

        # Execute F8
        wnd0.sendVKey(8)

        # Verify Status Bar
        try:
            sbar = session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                raise RuntimeError(f"Lỗi SAP MB52: {sbar.Text}")
        except Exception as exc:
            if "Lỗi SAP MB52:" in str(exc):
                raise

        # Read ALV Grid
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        if row_count <= 0:
            raise ValueError("SAP MB52 không trả về dòng dữ liệu nào cho phạm vi đã chọn.")

        headers = ["MATNR", "WERKS", "LGORT", "CHARG", "Unrestricted", "MEINS"]
        extracted_rows: List[List[Any]] = []

        # Extract rows from ALV Grid
        for i in range(row_count):
            mat = str(grid.GetCellValue(i, "MATNR")).strip()
            plant_val = str(grid.GetCellValue(i, "WERKS")).strip()
            sloc_val = str(grid.GetCellValue(i, "LGORT")).strip()
            charg_val = str(grid.GetCellValue(i, "CHARG")).strip()
            qty_raw = grid.GetCellValue(i, "LABST") or grid.GetCellValue(i, "MENGE") or "0"
            unit_val = str(grid.GetCellValue(i, "MEINS")).strip()
            extracted_rows.append([mat, plant_val, sloc_val, charg_val, qty_raw, unit_val])

        # Write intermediate Excel file
        return write_excel_dataset(headers, extracted_rows, prefix="SAP_MB52_")
    finally:
        pythoncom.CoUninitialize()


def download_sap_mb52_live(
    plants: Optional[List[str]] = None,
    sloc: Optional[str] = None,
    layout: str = "/KHOMA",
    materials: Optional[List[str]] = None,
    timeout: int = 60
) -> Path:
    """Execute MB52 download on COM thread with timeout."""
    future = _com_executor.submit(_execute_mb52_com, plants, sloc, layout, materials)
    return future.result(timeout=timeout)


# ---------------------------------------------------------------------------
# 2. MB51 Direct COM Automation
# ---------------------------------------------------------------------------
def _execute_mb51_com(
    obs: List[str],
    plant: str = "3093",
    layout: str = "/khoma",
    conn_idx: int = 0,
    sess_idx: int = 0
) -> Path:
    """Worker implementation of MB51 export over SAP GUI COM."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        session = get_sap_com_session(conn_idx, sess_idx)
        wnd0 = session.findById("wnd[0]")
        wnd0.maximize()

        # Open MB51 transaction
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMB51"
        wnd0.sendVKey(0)

        # Set Layout /khoma
        try:
            session.findById("wnd[0]/usr/ctxtALV_DEF").Text = layout
        except Exception:
            pass

        # Set source plant (default 3093 per baseline)
        if plant:
            try:
                session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = plant
            except Exception:
                pass

        # Set Multiple Selection for XBLNR (Outbound Delivery)
        session.findById("wnd[0]/usr/btn%_XBLNR_%_APP_%-VALU_PUSH").Press()
        try:
            session.findById("wnd[1]/tbar[0]/btn[16]").Press()
        except Exception:
            pass

        # Paste OB list from clipboard
        set_windows_clipboard("\r\n".join(obs))
        session.findById("wnd[1]/tbar[0]/btn[24]").Press()
        session.findById("wnd[1]").sendVKey(8)  # Confirm popup F8

        # Run MB51 F8
        wnd0.sendVKey(8)

        # Verify Status Bar
        try:
            sbar = session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                raise RuntimeError(f"Lỗi SAP MB51: {sbar.Text}")
        except Exception as exc:
            if "Lỗi SAP MB51:" in str(exc):
                raise

        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        if row_count <= 0:
            raise ValueError("SAP MB51 không có dữ liệu nào cho các số OB đã chỉ định.")

        headers = ["MBLNR", "MATNR", "CHARG", "MENGE", "XBLNR", "EBELN", "BWART", "MEINS"]
        extracted_rows: List[List[Any]] = []

        for i in range(row_count):
            mblnr = str(grid.GetCellValue(i, "MBLNR")).strip()
            matnr = str(grid.GetCellValue(i, "MATNR")).strip()
            charg = str(grid.GetCellValue(i, "CHARG")).strip()
            menge = grid.GetCellValue(i, "MENGE")
            xblnr = str(grid.GetCellValue(i, "XBLNR")).strip()
            ebeln = str(grid.GetCellValue(i, "EBELN")).strip()
            bwart = str(grid.GetCellValue(i, "BWART")).strip()
            meins = str(grid.GetCellValue(i, "MEINS")).strip()
            extracted_rows.append([mblnr, matnr, charg, menge, xblnr, ebeln, bwart, meins])

        return write_excel_dataset(headers, extracted_rows, prefix="SAP_MB51_")
    finally:
        pythoncom.CoUninitialize()


def extract_sap_mb51_live(
    obs: List[str],
    plant: str = "3093",
    layout: str = "/khoma",
    timeout: int = 60
) -> Path:
    """Execute MB51 extraction on COM thread with timeout."""
    future = _com_executor.submit(_execute_mb51_com, obs, plant, layout)
    return future.result(timeout=timeout)


# ---------------------------------------------------------------------------
# 3. VL06I Direct COM Automation
# ---------------------------------------------------------------------------
def _execute_vl06i_com(
    obs: List[str],
    conn_idx: int = 0,
    sess_idx: int = 0
) -> Path:
    """Worker implementation of VL06I export over SAP GUI COM."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        session = get_sap_com_session(conn_idx, sess_idx)
        wnd0 = session.findById("wnd[0]")
        wnd0.maximize()

        # Open VL06I transaction
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL06I"
        wnd0.sendVKey(0)

        # 1. Press BUTTON7: List Inbound Deliveries
        session.findById("wnd[0]/usr/btnBUTTON7").Press()

        # 2. Clear delivery date filter
        try:
            session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").Text = ""
        except Exception:
            pass

        # 3. Multiple selection for IT_LIFEX (OB)
        session.findById("wnd[0]/usr/btn%_IT_LIFEX_%_APP_%-VALU_PUSH").Press()
        try:
            session.findById("wnd[1]/tbar[0]/btn[16]").Press()
        except Exception:
            pass

        set_windows_clipboard("\r\n".join(obs))
        session.findById("wnd[1]/tbar[0]/btn[24]").Press()
        session.findById("wnd[1]").sendVKey(8)  # F8 confirm popup

        # 4. Set goods movement status IT_WBSTK-LOW = "a"
        try:
            session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").Text = "a"
        except Exception:
            pass

        # 5. Execute F8
        wnd0.sendVKey(8)

        # Verify Status Bar
        try:
            sbar = session.findById("wnd[0]/sbar")
            if sbar.MessageType == "E":
                raise RuntimeError(f"Lỗi SAP VL06I: {sbar.Text}")
        except Exception as exc:
            if "Lỗi SAP VL06I:" in str(exc):
                raise

        # 6. Switch to Item detail view (btn[18] on tbar[1])
        session.findById("wnd[0]/tbar[1]/btn[18]").Press()

        # 7. Read Item ALV Grid
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        row_count = grid.RowCount
        if row_count <= 0:
            raise ValueError("SAP VL06I không có dữ liệu chi tiết item cho các số OB đã chỉ định.")

        headers = ["LIFEX", "MATNR", "CHARG", "LFIMG", "VBELN", "WERKS", "LGORT", "VRKME", "WBSTK"]
        extracted_rows: List[List[Any]] = []

        for i in range(row_count):
            lifex = str(grid.GetCellValue(i, "LIFEX")).strip()
            matnr = str(grid.GetCellValue(i, "MATNR")).strip()
            charg = str(grid.GetCellValue(i, "CHARG")).strip()
            lfimg = grid.GetCellValue(i, "LFIMG")
            vbeln = str(grid.GetCellValue(i, "VBELN")).strip()
            werks = str(grid.GetCellValue(i, "WERKS")).strip()
            lgort = str(grid.GetCellValue(i, "LGORT")).strip()
            vrkme = str(grid.GetCellValue(i, "VRKME")).strip()
            wbstk = str(grid.GetCellValue(i, "WBSTK")).strip()
            extracted_rows.append([lifex, matnr, charg, lfimg, vbeln, werks, lgort, vrkme, wbstk])

        return write_excel_dataset(headers, extracted_rows, prefix="SAP_VL06I_")
    finally:
        pythoncom.CoUninitialize()


def extract_sap_vl06i_live(
    obs: List[str],
    timeout: int = 60
) -> Path:
    """Execute VL06I extraction on COM thread with timeout."""
    future = _com_executor.submit(_execute_vl06i_com, obs)
    return future.result(timeout=timeout)
