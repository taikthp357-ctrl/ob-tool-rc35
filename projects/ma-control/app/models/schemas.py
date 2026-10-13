# -*- coding: utf-8 -*-
"""Pydantic schemas for MA CONTROL 5-step workflow."""

from typing import List, Optional, Any, Dict
from pydantic import BaseModel, Field, StrictFloat, StrictInt
import datetime

# -------------------------------------------------------------
# 1. Materials & Model Mappings
# -------------------------------------------------------------
class MaterialSchema(BaseModel):
    material_code: str
    description: Optional[str] = None
    model: Optional[str] = None
    shop: Optional[str] = None
    standard_pack: float = 1.0
    default_sloc: Optional[str] = None
    is_external: int = 0

class ModelMappingSchema(BaseModel):
    material_code: str
    model_code: str
    shop_code: Optional[str] = None
    note: Optional[str] = None

# -------------------------------------------------------------
# 2. Stock Schemas (Step 1)
# -------------------------------------------------------------
class StockItemSchema(BaseModel):
    id: Optional[int] = None
    snapshot_id: Optional[int] = None
    plant: str
    sloc: str
    material_code: str
    batch: str
    stock_type: str = "UNRESTRICTED"
    unrestricted_qty: float = 0.0
    unit: str = "PCS"
    priority_rank: int = 0
    pull_line_warning: Optional[str] = None

class StockSnapshotSchema(BaseModel):
    id: Optional[int] = None
    source_type: str
    snapshot_time: str
    plant: Optional[str] = None
    sloc: Optional[str] = None
    total_rows: int = 0
    total_materials: int = 0
    total_qty: float = 0.0
    notes: Optional[str] = None

class StockUploadRequest(BaseModel):
    source: str = "MANUAL_UPLOAD"
    plant: Optional[str] = "3001"
    sloc: Optional[str] = "1003"
    items: List[Dict[str, Any]]

# -------------------------------------------------------------
# 3. OB Schemas (Step 2)
# -------------------------------------------------------------
class OBItemSchema(BaseModel):
    id: Optional[int] = None
    item_uid: Optional[str] = None
    branch_tab: str = "BODY_3001"
    row_order: int = 0
    selected: int = 1
    material_code: str
    model: Optional[str] = None
    req_qty: float = 0.0
    avail_qty: float = 0.0
    missing_qty: float = 0.0
    batch: Optional[str] = None
    shop: Optional[str] = None
    sto_number: Optional[str] = None
    ob_number: Optional[str] = None
    status: str = "Chưa kiểm tra"
    rc52_code: Optional[str] = None
    note: Optional[str] = None
    error_message: Optional[str] = None
    pick_order_id: Optional[int] = None

OB_STATUS_MAP = {
    # 8 distinct status categories per 13/10 directive, supporting legacy/test aliases:
    # 1. UNCHECKED / UNVERIFIED
    "UNCHECKED": ("UNCHECKED", "Mới nhập / Chưa kiểm tra", "Chưa kiểm tra"),
    "UNVERIFIED": ("UNVERIFIED", "Chưa kiểm tra", "Chưa kiểm tra"),
    # 2. INVALID_OR_CANCELLED / INVALID_CANCELLED
    "INVALID_OR_CANCELLED": ("INVALID_OR_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "INVALID": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "CANCELLED": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "INVALID_CANCELLED": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    # 3. LOCKED_OR_INSUFFICIENT / LOCKED_SHORTAGE
    "LOCKED_OR_INSUFFICIENT": ("LOCKED_OR_INSUFFICIENT", "Đang khóa / Thiếu tồn", "Cần xử lý"),
    "LOCKED": ("LOCKED_SHORTAGE", "Đang khóa / Thiếu tồn", "Cần xử lý"),
    "SHORTAGE": ("LOCKED_SHORTAGE", "Đang khóa / Thiếu tồn", "Cần xử lý"),
    "LOCKED_SHORTAGE": ("LOCKED_SHORTAGE", "Đang khóa / Thiếu tồn", "Cần xử lý"),
    # 4. ELIGIBLE_FOR_PICK / READY_FOR_PICK
    "ELIGIBLE_FOR_PICK": ("ELIGIBLE_FOR_PICK", "Đủ điều kiện tạo pick", "Đủ điều kiện"),
    "READY_FOR_PICK": ("READY_FOR_PICK", "Đủ điều kiện tạo pick", "Đủ điều kiện"),
    # 5. PENDING_CF / WAITING_CF
    "PENDING_CF": ("PENDING_CF", "Chờ CF", "Chờ CF"),
    "WAITING_CF": ("WAITING_CF", "Chờ CF", "Chờ CF"),
    # 6. PARTIAL_CF
    "PARTIAL_CF": ("PARTIAL_CF", "CF một phần", "Chờ CF"),
    # 7. CONFIRMED_CF / CF_COMPLETED
    "CONFIRMED_CF": ("CONFIRMED_CF", "Đã xác nhận CF đầy đủ", "Đã hoàn thành"),
    "CF_COMPLETED": ("CF_COMPLETED", "Đã CF đầy đủ", "Đã hoàn thành"),
    "CONFIRMED": ("CONFIRMED_CF", "Đã xác nhận CF đầy đủ", "Đã hoàn thành"),
    # 8. ELIGIBLE_BBBG / READY_FOR_BBBG
    "ELIGIBLE_BBBG": ("ELIGIBLE_BBBG", "Đủ điều kiện tạo/đối soát BBBG", "Đã hoàn thành"),
    "READY_FOR_BBBG": ("READY_FOR_BBBG", "Đủ điều kiện BBBG", "Đã hoàn thành"),
    # Vietnamese aliases
    "Chưa kiểm tra": ("UNVERIFIED", "Chưa kiểm tra", "Chưa kiểm tra"),
    "Mới nhập/chưa kiểm tra": ("UNCHECKED", "Mới nhập / Chưa kiểm tra", "Chưa kiểm tra"),
    "Không hợp lệ / Đã hủy": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "Không hợp lệ": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "Đã hủy": ("INVALID_CANCELLED", "Không hợp lệ / Đã hủy", "Cần xử lý"),
    "Đang khóa / Thiếu tồn": ("LOCKED_SHORTAGE", "Đang khóa / Thiếu tồn", "Cần xử lý"),
    "Thiếu SL OB": ("LOCKED_SHORTAGE", "Thiếu SL OB", "Cần xử lý"),
    "Đủ điều kiện tạo pick": ("READY_FOR_PICK", "Đủ điều kiện tạo pick", "Đủ điều kiện"),
    "Sẵn sàng điền": ("READY_FOR_PICK", "Sẵn sàng điền", "Đủ điều kiện"),
    "READY": ("READY_FOR_PICK", "Sẵn sàng điền", "Đủ điều kiện"),
    "Chờ CF": ("WAITING_CF", "Chờ CF", "Chờ CF"),
    "Chờ pick / CF": ("WAITING_CF", "Chờ pick / CF", "Chờ CF"),
    "CF một phần": ("PARTIAL_CF", "CF một phần", "Chờ CF"),
    "Đã xác nhận CF đầy đủ": ("CONFIRMED_CF", "Đã xác nhận CF đầy đủ", "Đã hoàn thành"),
    "Đã CF đầy đủ": ("CF_COMPLETED", "Đã CF đầy đủ", "Đã hoàn thành"),
    "Đã CF": ("CF_COMPLETED", "Đã CF", "Đã hoàn thành"),
    "USER_CONFIRMED": ("CF_COMPLETED", "Đã CF", "Đã hoàn thành"),
    "Đủ điều kiện BBBG": ("READY_FOR_BBBG", "Đủ điều kiện BBBG", "Đã hoàn thành"),
    "Đủ điều kiện tạo/đối soát BBBG": ("ELIGIBLE_BBBG", "Đủ điều kiện tạo/đối soát BBBG", "Đã hoàn thành"),
    "Đã lập BBBG": ("READY_FOR_BBBG", "Đã lập BBBG", "Đã hoàn thành"),
    "HANDED_OVER": ("READY_FOR_BBBG", "Đã lập BBBG", "Đã hoàn thành"),
}

def normalize_ob_status(raw_status: Optional[str]) -> Dict[str, Any]:
    """Normalize an OB item status into standardized code, label, group, and action indicators."""
    s = (raw_status or "UNVERIFIED").strip()
    if s in OB_STATUS_MAP:
        code, label, group = OB_STATUS_MAP[s]
    else:
        code, label, group = ("UNVERIFIED", s, "Chưa kiểm tra")
    return {
        "status_code": code,
        "status_label": label,
        "status_group": group,
        "is_ready_for_pick": code in ("ELIGIBLE_FOR_PICK", "READY_FOR_PICK"),
        "is_waiting_cf": code in ("PENDING_CF", "WAITING_CF", "PARTIAL_CF"),
        "is_cf_completed": code in ("CONFIRMED_CF", "CF_COMPLETED", "ELIGIBLE_BBBG", "READY_FOR_BBBG"),
        "is_action_needed": group == "Cần xử lý",
    }

class OBBatchSaveRequest(BaseModel):
    branch: Optional[str] = "BODY_3001"
    branch_tab: Optional[str] = None
    items: List[Dict[str, Any]]

class OBCheckRequest(BaseModel):
    branch_tab: str = "BODY_3001"
    source_plant: Optional[str] = None
    source_sloc: Optional[str] = None
    item_ids: Optional[List[int]] = None

class OBFillRequest(BaseModel):
    branch_tab: str = "BODY_3001"
    item_ids: Optional[List[int]] = None

# -------------------------------------------------------------
# 4. Pick Order Schemas (Step 3 & 4)
# -------------------------------------------------------------
PICK_STATUS_CODE_TO_LABEL = {
    "WAITING_USER_CF": "Chờ người dùng CF",
    "Chờ người dùng CF": "Chờ người dùng CF",
    "WAITING_CF": "Chờ người dùng CF",
    "PARTIAL_CF": "CF một phần",
    "CF một phần": "CF một phần",
    "CF_COMPLETED": "Đã CF",
    "USER_CONFIRMED": "Đã CF",
    "Đã CF": "Đã CF",
    "HANDED_OVER": "Đã lập BBBG",
    "Đã lập BBBG": "Đã lập BBBG",
    "REJECTED": "Từ chối",
    "Từ chối": "Từ chối",
}

PICK_STATUS_LABEL_TO_CODE = {
    "Chờ người dùng CF": "WAITING_USER_CF",
    "WAITING_USER_CF": "WAITING_USER_CF",
    "WAITING_CF": "WAITING_USER_CF",
    "CF một phần": "PARTIAL_CF",
    "PARTIAL_CF": "PARTIAL_CF",
    "CF_COMPLETED": "USER_CONFIRMED",
    "Đã CF": "USER_CONFIRMED",
    "USER_CONFIRMED": "USER_CONFIRMED",
    "Đã lập BBBG": "HANDED_OVER",
    "HANDED_OVER": "HANDED_OVER",
    "Từ chối": "REJECTED",
    "REJECTED": "REJECTED",
}

def normalize_pick_status(raw_status: Optional[str]) -> Dict[str, str]:
    """Normalize pick order status into both internal code and Vietnamese UI label."""
    s = (raw_status or "Chờ người dùng CF").strip()
    code = PICK_STATUS_LABEL_TO_CODE.get(s, s)
    label = PICK_STATUS_CODE_TO_LABEL.get(s, s)
    return {
        "status_code": code,
        "status_label": label,
        "pick_status": code,
    }

class PickOrderCreateRequest(BaseModel):
    branch_tab: Optional[str] = "BODY_3001"
    shop: str = "BODY"
    model: str = "VF8"
    items: List[Dict[str, Any]]

class PickOrderItemSchema(BaseModel):
    id: Optional[int] = None
    pick_order_id: int
    ob_item_id: Optional[int] = None
    material_code: str
    model: Optional[str] = None
    batch: Optional[str] = None
    shop: Optional[str] = None
    sto_number: Optional[str] = None
    ob_number: Optional[str] = None
    qty: float = 0.0
    cf_status: str = "PENDING_CF"
    cf_actual_qty: Optional[float] = None

class PickOrderSchema(BaseModel):
    id: Optional[int] = None
    pick_order_code: str
    pick_order_no: Optional[str] = None
    branch_tab: Optional[str] = None
    shop: str
    model: str
    item_count: int = 0
    total_items: Optional[int] = None
    total_qty: float = 0.0
    status: str = "Chờ người dùng CF"
    status_code: str = "WAITING_USER_CF"
    status_label: str = "Chờ người dùng CF"
    pick_status: str = "WAITING_USER_CF"
    cf_verified_by: Optional[str] = None
    cf_verified_at: Optional[str] = None
    notes: Optional[str] = None
    created_date: Optional[str] = None
    created_at: Optional[str] = None

class PickVerifyCFRequest(BaseModel):
    verified_by: str = Field(min_length=1)
    actual_quantities: Dict[str, StrictFloat | StrictInt]
    notes: Optional[str] = None

class ConfirmPickLegacyRequest(BaseModel):
    confirmed_by: str = "User_Operator"
    actual_quantities: Optional[Dict[str, StrictFloat | StrictInt]] = None

class WaitingCFLineSaveRequest(BaseModel):
    line_id: int
    actual_qty: float
    verified_by: str = Field(min_length=1)
    reason: Optional[str] = None
    evidence: Optional[str] = None

class WaitingCFBulkSaveRequest(BaseModel):
    verified_by: str = Field(min_length=1)
    lines: List[Dict[str, Any]]
    notes: Optional[str] = None

# -------------------------------------------------------------
# 5. BBBG Schemas (Step 5)
# -------------------------------------------------------------
class BBBGCreateRequest(BaseModel):
    pick_order_id: Optional[int] = None
    pick_order_code: Optional[str] = None
    pick_order_no: Optional[str] = None
    slip_no: Optional[str] = None
    doc_no: Optional[str] = None
    handover_person: Optional[str] = None
    receiver_person: Optional[str] = None
    sender_name: Optional[str] = None
    receiver_name: Optional[str] = None
    storage_location_note: Optional[str] = None
    storage_bin: Optional[str] = None
    is_external: int = 0
    template_mode: str = "Auto"
    font_size: int = 14
    auto_email: bool = False
    auto_print: bool = False
    attach_gos: bool = True

class BBBGRecordSchema(BaseModel):
    id: Optional[int] = None
    bbbg_uuid: str
    bbbg_no: Optional[str] = None
    dedup_key: str
    pick_order_id: Optional[int] = None
    slip_no: Optional[str] = None
    doc_no: Optional[str] = None
    shop: str
    model: str
    total_qty: float
    sender_name: Optional[str] = None
    receiver_name: Optional[str] = None
    storage_bin: Optional[str] = None
    email_status: str = "PENDING"
    print_status: str = "PENDING"
    unc_export_status: str = "PENDING"
    gos_bbbg_attached: int = 0
    gos_vat_attached: int = 0
    status: str = "CREATED"
    created_at: Optional[str] = None

# -------------------------------------------------------------
# 6. Job Events & Settings Schemas
# -------------------------------------------------------------
class JobEventSchema(BaseModel):
    id: Optional[int] = None
    job_id: str
    business_key: Optional[str] = None
    job_type: str
    step_name: Optional[str] = None
    event_type: Optional[str] = None
    target_ref: Optional[str] = None
    status: str
    duration_ms: int = 0
    sap_status_msg: Optional[str] = None
    evidence_json: Optional[str] = None
    created_at: Optional[str] = None

class SettingsUpdateSchema(BaseModel):
    key: str
    value: Any

# -------------------------------------------------------------
# 7. Chờ CF Direct Paste & User-Managed Schemas
# -------------------------------------------------------------
class WaitingCFDirectPasteRequest(BaseModel):
    raw_text: Optional[str] = None
    items: Optional[List[Dict[str, Any]]] = None
    allow_duplicates: bool = False
    direct_order_ref: Optional[str] = None

class WaitingCFUpdateLineRequest(BaseModel):
    line_id: int
    ob_number: Optional[str] = None
    material_code: Optional[str] = None
    plant: Optional[str] = None
    batch: Optional[str] = None
    qty: Optional[float] = None
    packing_box: Optional[str] = None
    rack: Optional[str] = None

# -------------------------------------------------------------
# 8. Low Stock & Multi-Batch Priority Allocation Schemas
# -------------------------------------------------------------
class LowStockConfigUpdate(BaseModel):
    default_threshold: Optional[float] = None
    max_age_hours: Optional[float] = None
    thresholds: Optional[List[Dict[str, Any]]] = None

class LowStockInspectRequest(BaseModel):
    material_code: str
    plant: str = "3001"
    required_qty: float
    scope_sloc: Optional[str] = None
    shop: Optional[str] = None
    max_batches: int = 3

class LowStockBatchAllocation(BaseModel):
    batch: str
    sloc: str
    allocated_qty: float

class LowStockApplyRequest(BaseModel):
    material_code: str
    plant: str = "3001"
    required_qty: float
    shop: Optional[str] = None
    scope_sloc: Optional[str] = None
    allocations: List[Dict[str, Any]]
    target: str = "PROPOSAL"  # "PROPOSAL" | "WAITING_CF" | "OB_ITEMS"
    ob_number: Optional[str] = None
    branch_tab: Optional[str] = None
    evidence_note: Optional[str] = None
    other_shop_reserves: Optional[Dict[str, Any]] = None

# -------------------------------------------------------------
# 9. Direct SAP MB52 & Direct Reconciliation Schemas
# -------------------------------------------------------------
class SAPDownloadMB52Request(BaseModel):
    plants: Optional[List[str]] = None
    sloc: Optional[str] = None
    layout: str = "/KHOMA"
    scope: Optional[str] = None
    test_bytes_b64: Optional[str] = None
    test_filename: Optional[str] = None

class DirectReconcileSAPRequest(BaseModel):
    test_mb51_bytes_b64: Optional[str] = None
    test_vl06i_bytes_b64: Optional[str] = None
    test_mb51_filename: Optional[str] = None
    test_vl06i_filename: Optional[str] = None

class LowStockTransferRequest(BaseModel):
    items: List[Dict[str, Any]]
    branch_tab: str = "BODY_3001"
