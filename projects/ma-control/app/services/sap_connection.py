"""Read-only SAP session probe on one COM-initialized worker thread."""
from concurrent.futures import ThreadPoolExecutor, TimeoutError
import threading

_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="ma-sap-read")
_lock = threading.Lock()
_pending = None


def _probe():
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return {"state": "UNAVAILABLE", "message": "Máy chưa có pywin32/SAP GUI Scripting"}
    pythoncom.CoInitialize()
    try:
        sap = win32com.client.GetObject("SAPGUI")
        engine = sap.GetScriptingEngine
        sessions = []
        for ci in range(engine.Children.Count):
            connection = engine.Children(ci)
            for si in range(connection.Children.Count):
                session = connection.Children(si)
                info = session.Info
                sessions.append({"connection_index": ci, "session_index": si,
                                 "system": str(info.SystemName), "client": str(info.Client),
                                 "busy": bool(session.Busy)})
        return {"state": "CONNECTED" if sessions else "DISCONNECTED", "sessions": sessions,
                "message": "Đã đọc được phiên SAP; chưa chứng nhận thao tác ghi" if sessions else "Chưa có phiên SAP đăng nhập"}
    except Exception:
        return {"state": "DISCONNECTED", "message": "Chưa đọc được phiên SAP GUI đang đăng nhập"}
    finally:
        pythoncom.CoUninitialize()


def probe_connection(timeout=3):
    global _pending
    with _lock:
        if _pending is None or _pending.done():
            _pending = _executor.submit(_probe)
        future = _pending
    try:
        return future.result(timeout=timeout)
    except TimeoutError:
        return {"state": "TIMEOUT", "message": "SAP chưa phản hồi; không thực hiện thao tác ghi"}
