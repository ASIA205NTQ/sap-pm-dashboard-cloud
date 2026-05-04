import hashlib
import io
import os
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from supabase import create_client
except Exception:  # pragma: no cover - app vẫn chạy local nếu chưa cài/chưa cấu hình Supabase
    create_client = None

st.set_page_config(
    page_title="SAP NO/MO Weekly Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")
DEFAULT_SUPABASE_BUCKET = "sap-files"
LOGO_PATH = Path("assets/petrolimex_logo_header.png")

REQUIRED_NO_COLUMNS = [
    "Mã sự cố",
    "Tên sự cố",
    "Tên T.thái",
    "Ngày tạo sự cố",
    "Lệnh sửa chữa",
    "Tên KV chức năng",
    "Tên nhóm T/bị",
]

REQUIRED_MO_COLUMNS = [
    "Lệnh sửa chữa",
    "Tên T.thái",
    "Lệnh S/ chữa từ",
    "Mã sự cố",
    "Mã kế hoạch",
    "Tên KV chức năng",
    "Tên nhóm T/bị",
]

OPEN_NO_STATUSES = {
    "CREATED",
    "APPROVED",
    "IN_PROGRESS",
    "COMPLETED_NOT_CLOSED",
}

STATUS_ORDER = [
    "Khởi tạo",
    "Phê duyệt",
    "Thực hiện",
    "Hoàn thành",
    "Đã đóng",
    "Từ chối",
    "Khác",
]

st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.0rem; padding-bottom: 2rem;}
    .logo-wrap {display:flex; justify-content:center; align-items:center; margin-top: 0.3rem; margin-bottom: 0.4rem;}
    .dashboard-title {text-align:center; font-size:2.05rem; font-weight:900; color:#06145f; margin-bottom:0.15rem;}
    .dashboard-subtitle {text-align:center; color:#475569; font-size:1.03rem; margin-bottom:1.2rem;}
    .section-title {font-size:1.62rem; font-weight:900; color:#0f172a; margin-top:0.3rem; margin-bottom:0.25rem;}
    .section-note {color:#64748b; font-size:0.98rem;}
    .metric-title {font-size:1.05rem; font-weight:900; color:#0f172a; margin-bottom:0.25rem;}
    .metric-value {font-size:2.15rem; font-weight:950; line-height:1.08; color:#06145f; margin-bottom:0.1rem;}
    .metric-delta {font-size:0.95rem; font-weight:800; color:#475569; min-height:1.35rem;}
    .metric-card-note {font-size:0.82rem; color:#64748b; margin-top:0.15rem;}
    .warning-box {border:1px solid #f59e0b; background:#fffbeb; padding:12px 14px; border-radius:12px; color:#78350f; margin-bottom:0.8rem;}
    .ok-box {border:1px solid #10b981; background:#ecfdf5; padding:12px 14px; border-radius:12px; color:#064e3b; margin-bottom:0.8rem;}
    .saved-box {border:1px solid #2563eb; background:#eff6ff; padding:12px 14px; border-radius:12px; color:#172554; margin-bottom:0.8rem;}
    div[data-testid="stDownloadButton"] button {padding:0.28rem 0.55rem; min-height:2.0rem; font-size:0.86rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


@dataclass
class MemoryUpload:
    name: str
    raw: bytes

    def getvalue(self) -> bytes:
        return self.raw


def now_vn() -> datetime:
    return datetime.now(tz=VN_TZ)


def safe_filename(name: str) -> str:
    base = os.path.basename(name or "file")
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_")
    return base or "file"


def format_vn_date(value) -> str:
    if value is None or value == "":
        return "chưa rõ"
    if isinstance(value, datetime):
        return value.astimezone(VN_TZ).strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    try:
        return pd.to_datetime(value).date().strftime("%d/%m/%Y")
    except Exception:
        return str(value)


def format_vn_datetime(value) -> str:
    if value is None or value == "":
        return "chưa rõ"
    try:
        ts = pd.to_datetime(value, utc=True)
        if pd.isna(ts):
            return str(value)
        return ts.tz_convert("Asia/Ho_Chi_Minh").strftime("%d/%m/%Y %H:%M")
    except Exception:
        try:
            if isinstance(value, datetime):
                return value.astimezone(VN_TZ).strftime("%d/%m/%Y %H:%M")
        except Exception:
            pass
    return str(value)


def parse_report_date(value) -> date:
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return date.today()


def clean_text(value):
    if pd.isna(value):
        return np.nan
    text = str(value).replace("\xa0", " ").replace("Ð", "Đ").replace("ð", "đ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.lower() in {"nan", "none", "nat", ""}:
        return np.nan
    return text


def clean_id(value):
    value = clean_text(value)
    if pd.isna(value):
        return ""
    if re.fullmatch(r"\d+\.0", value):
        return value[:-2]
    return value


def parse_date_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=False).dt.date


def hash_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def file_hash(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    return hash_bytes(uploaded_file.getvalue())


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buffer = io.BytesIO()
    safe_sheet = re.sub(r"[\\/*?:\[\]]", "_", sheet_name)[:31] or "Data"
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=safe_sheet, index=False)
    return buffer.getvalue()


def read_mht_table(uploaded_file) -> pd.DataFrame:
    raw = uploaded_file.getvalue()
    text = raw.decode("utf-8", errors="ignore")
    tables = pd.read_html(io.StringIO(text))
    if not tables:
        raise ValueError("Không tìm thấy bảng HTML trong file.")
    df = tables[0]
    if df.empty:
        raise ValueError("Bảng dữ liệu rỗng.")
    header = [clean_text(x) for x in df.iloc[0].tolist()]
    df = df.iloc[1:].copy()
    df.columns = header
    df = df.loc[:, [not pd.isna(c) for c in df.columns]]
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].map(clean_text)
    return df.reset_index(drop=True)


def normalize_no_status(status: str) -> str:
    status = clean_text(status)
    if pd.isna(status):
        return "Khác"
    s = str(status).lower()
    if "đã đóng" in s or "da dong" in s or "đóng" in s:
        return "Đã đóng"
    if "từ chối" in s or "tu choi" in s:
        return "Từ chối"
    if "hoàn thành" in s or "hoan thanh" in s:
        return "Hoàn thành"
    if "thực hiện" in s or "thuc hien" in s:
        return "Thực hiện"
    if "phê duyệt" in s or "phe duyet" in s:
        return "Phê duyệt"
    if "khởi tạo" in s or "khoi tao" in s:
        return "Khởi tạo"
    return str(status)


def no_status_group(status: str) -> str:
    status = normalize_no_status(status)
    mapping = {
        "Đã đóng": "CLOSED",
        "Từ chối": "REJECTED",
        "Hoàn thành": "COMPLETED_NOT_CLOSED",
        "Thực hiện": "IN_PROGRESS",
        "Phê duyệt": "APPROVED",
        "Khởi tạo": "CREATED",
    }
    return mapping.get(status, "OTHER")


def status_group_label(group: str) -> str:
    return {
        "CLOSED": "Đã đóng",
        "REJECTED": "Từ chối",
        "COMPLETED_NOT_CLOSED": "Hoàn thành chưa đóng",
        "IN_PROGRESS": "Đang thực hiện",
        "APPROVED": "Phê duyệt",
        "CREATED": "Khởi tạo",
        "OTHER": "Khác",
    }.get(group, group)


def normalize_mo_status(status: str) -> str:
    status = clean_text(status)
    if pd.isna(status):
        return "Khác"
    s = str(status).lower()
    if "hoàn thành kỹ thuật" in s or "hoan thanh ky thuat" in s:
        return "Hoàn thành kỹ thuật"
    return str(status)


def extract_store_name(row: pd.Series) -> str:
    for col in ["Tên KV chức năng", "Tên PB S/dụng", "Tên đơn vị"]:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            return str(row[col]).strip()
    return "Không rõ cửa hàng"


def normalize_no(df: pd.DataFrame, report_date: Optional[date] = None) -> pd.DataFrame:
    missing = [c for c in REQUIRED_NO_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"File NO thiếu cột bắt buộc: {', '.join(missing)}")
    out = df.copy()
    out["Mã sự cố"] = out["Mã sự cố"].map(clean_id)
    out = out[out["Mã sự cố"].astype(str).str.len() > 0].copy()
    out["Lệnh sửa chữa"] = out["Lệnh sửa chữa"].map(clean_id) if "Lệnh sửa chữa" in out else ""
    out["Ngày tạo sự cố_dt"] = parse_date_series(out["Ngày tạo sự cố"])
    out["Trạng thái chuẩn"] = out["Tên T.thái"].map(normalize_no_status)
    out["Nhóm trạng thái"] = out["Tên T.thái"].map(no_status_group)
    out["Nhóm trạng thái hiển thị"] = out["Nhóm trạng thái"].map(status_group_label)
    out["Còn mở"] = out["Nhóm trạng thái"].isin(OPEN_NO_STATUSES)
    out["Có lệnh sửa chữa"] = out["Lệnh sửa chữa"].astype(str).str.len() > 0
    out["Cửa hàng"] = out.apply(extract_store_name, axis=1)
    out["Nhóm công việc"] = out.get("Tên nhóm T/bị", pd.Series(index=out.index, dtype="object")).fillna("Không rõ")
    out["Tổ đội"] = out.get("Tên tổ đội S/ chữa", pd.Series(index=out.index, dtype="object")).fillna("Không rõ")
    out["Người phụ trách"] = out.get("Tên nhân viên phụ trách KT", pd.Series(index=out.index, dtype="object")).fillna("Chưa phân công")
    if report_date is None:
        max_date = out["Ngày tạo sự cố_dt"].dropna().max()
        report_date = max_date if pd.notna(max_date) else date.today()
    out["Tuổi tồn ngày"] = out["Ngày tạo sự cố_dt"].map(lambda d: (report_date - d).days if pd.notna(d) else np.nan)
    out["Nhóm tuổi tồn"] = pd.cut(
        out["Tuổi tồn ngày"],
        bins=[-1, 7, 14, 30, 99999],
        labels=["0-7 ngày", "8-14 ngày", "15-30 ngày", ">30 ngày"],
    ).astype("object").fillna("Không rõ")
    return out.reset_index(drop=True)


def normalize_mo(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in REQUIRED_MO_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"File MO thiếu cột bắt buộc: {', '.join(missing)}")
    out = df.copy()
    out["Lệnh sửa chữa"] = out["Lệnh sửa chữa"].map(clean_id)
    out["Mã sự cố"] = out["Mã sự cố"].map(clean_id) if "Mã sự cố" in out else ""
    out["Mã kế hoạch"] = out["Mã kế hoạch"].map(clean_id) if "Mã kế hoạch" in out else ""
    out = out[out["Lệnh sửa chữa"].astype(str).str.len() > 0].copy()
    out["Trạng thái MO chuẩn"] = out["Tên T.thái"].map(normalize_mo_status)
    out["Nguồn lệnh"] = out["Lệnh S/ chữa từ"].fillna("Không rõ").astype(str)
    out["Cửa hàng"] = out.apply(extract_store_name, axis=1)
    out["Nhóm công việc"] = out.get("Tên nhóm T/bị", pd.Series(index=out.index, dtype="object")).fillna("Không rõ")
    out["Tổ đội"] = out.get("Tên tổ đội S/ chữa", pd.Series(index=out.index, dtype="object")).fillna("Không rõ")
    out["MO từ sự cố"] = out["Mã sự cố"].astype(str).str.len() > 0
    out["MO từ kế hoạch"] = out["Mã kế hoạch"].astype(str).str.len() > 0
    return out.reset_index(drop=True)


def link_no_mo(no_df: pd.DataFrame, mo_df: pd.DataFrame) -> pd.DataFrame:
    mo_by_incident = mo_df[["Mã sự cố", "Lệnh sửa chữa", "Trạng thái MO chuẩn", "Nguồn lệnh"]].copy()
    mo_by_incident = mo_by_incident[mo_by_incident["Mã sự cố"].astype(str).str.len() > 0]
    mo_by_incident = mo_by_incident.drop_duplicates(subset=["Mã sự cố", "Lệnh sửa chữa"])
    linked = no_df.merge(
        mo_by_incident,
        on=["Mã sự cố", "Lệnh sửa chữa"],
        how="left",
        suffixes=("", "_MO"),
    )
    linked["Có MO khớp"] = linked["Trạng thái MO chuẩn"].notna()
    linked["Phân loại liên kết"] = np.select(
        [
            linked["Có MO khớp"],
            linked["Có lệnh sửa chữa"] & ~linked["Có MO khớp"],
            ~linked["Có lệnh sửa chữa"],
        ],
        [
            "NO có MO khớp",
            "NO có lệnh nhưng thiếu trong MO",
            "NO chưa có lệnh sửa chữa",
        ],
        default="Khác",
    )
    return linked


def calculate_metrics(no_df: pd.DataFrame, mo_df: pd.DataFrame, linked_df: pd.DataFrame) -> Dict[str, int]:
    return {
        "Tổng NO": len(no_df),
        "NO còn mở": int(no_df["Còn mở"].sum()),
        "NO đã đóng": int((no_df["Nhóm trạng thái"] == "CLOSED").sum()),
        "NO khởi tạo": int((no_df["Nhóm trạng thái"] == "CREATED").sum()),
        "NO hoàn thành": int((no_df["Nhóm trạng thái"] == "COMPLETED_NOT_CLOSED").sum()),
        "NO từ chối": int((no_df["Nhóm trạng thái"] == "REJECTED").sum()),
        "NO chưa có lệnh": int((~no_df["Có lệnh sửa chữa"]).sum()),
        "NO có lệnh nhưng thiếu MO": int((linked_df["Phân loại liên kết"] == "NO có lệnh nhưng thiếu trong MO").sum()),
        "MO tổng": len(mo_df),
        "MO từ sự cố": int(mo_df["MO từ sự cố"].sum()),
        "MO từ kế hoạch": int(mo_df["MO từ kế hoạch"].sum()),
    }


def build_metric_details(snapshot: Dict[str, object]) -> Dict[str, pd.DataFrame]:
    no = snapshot["no"]
    mo = snapshot["mo"]
    linked = snapshot["linked"]
    return {
        "Tổng NO": no,
        "NO còn mở": no[no["Còn mở"]],
        "NO đã đóng": no[no["Nhóm trạng thái"] == "CLOSED"],
        "NO khởi tạo": no[no["Nhóm trạng thái"] == "CREATED"],
        "NO hoàn thành": no[no["Nhóm trạng thái"] == "COMPLETED_NOT_CLOSED"],
        "NO từ chối": no[no["Nhóm trạng thái"] == "REJECTED"],
        "NO chưa có lệnh": no[~no["Có lệnh sửa chữa"]],
        "NO có lệnh nhưng thiếu MO": linked[linked["Phân loại liên kết"] == "NO có lệnh nhưng thiếu trong MO"],
        "MO tổng": mo,
        "MO từ sự cố": mo[mo["MO từ sự cố"]],
        "MO từ kế hoạch": mo[mo["MO từ kế hoạch"]],
    }


def build_investigation(linked_df: pd.DataFrame) -> pd.DataFrame:
    rows = linked_df.copy()
    reasons: List[List[str]] = []
    for _, r in rows.iterrows():
        item_reasons = []
        age = r.get("Tuổi tồn ngày")
        status_group = r.get("Nhóm trạng thái")
        if r.get("Còn mở") and not r.get("Có lệnh sửa chữa"):
            item_reasons.append("NO còn mở nhưng chưa sinh lệnh MO")
        if r.get("Còn mở") and r.get("Có lệnh sửa chữa") and not r.get("Có MO khớp"):
            item_reasons.append("NO có lệnh nhưng không thấy trong file MO")
        if status_group == "COMPLETED_NOT_CLOSED":
            item_reasons.append("NO hoàn thành nhưng chưa đóng")
        if r.get("Có MO khớp") and status_group in OPEN_NO_STATUSES:
            item_reasons.append("Đã có MO nhưng NO vẫn còn mở")
        if pd.notna(age) and age > 30 and status_group in OPEN_NO_STATUSES:
            item_reasons.append("NO tồn trên 30 ngày")
        elif pd.notna(age) and age > 14 and status_group in OPEN_NO_STATUSES:
            item_reasons.append("NO tồn trên 14 ngày")
        elif pd.notna(age) and age > 7 and status_group == "CREATED":
            item_reasons.append("Khởi tạo quá 7 ngày")
        if not item_reasons and r.get("Còn mở"):
            item_reasons.append("NO còn mở cần theo dõi")
        reasons.append(item_reasons)
    rows["Cờ điều tra"] = ["; ".join(x) for x in reasons]
    rows = rows[(rows["Còn mở"]) | (rows["Cờ điều tra"].astype(str).str.len() > 0)].copy()
    preferred = [
        "Mã sự cố",
        "Tên sự cố",
        "Trạng thái chuẩn",
        "Cửa hàng",
        "Nhóm công việc",
        "Tổ đội",
        "Ngày tạo sự cố",
        "Tuổi tồn ngày",
        "Nhóm tuổi tồn",
        "Lệnh sửa chữa",
        "Có MO khớp",
        "Trạng thái MO chuẩn",
        "Người phụ trách",
        "Cờ điều tra",
    ]
    return rows[[c for c in preferred if c in rows.columns]].sort_values(
        by=["Tuổi tồn ngày", "Trạng thái chuẩn"], ascending=[False, True], na_position="last"
    )


def compare_snapshots(prev_no: pd.DataFrame, curr_no: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    prev = prev_no[["Mã sự cố", "Trạng thái chuẩn", "Nhóm trạng thái", "Còn mở", "Cửa hàng", "Nhóm công việc"]].copy()
    curr = curr_no[["Mã sự cố", "Trạng thái chuẩn", "Nhóm trạng thái", "Còn mở", "Cửa hàng", "Nhóm công việc"]].copy()
    joined = prev.merge(curr, on="Mã sự cố", how="outer", suffixes=(" tuần trước", " tuần này"), indicator=True)
    joined["Loại biến động"] = np.select(
        [
            joined["_merge"] == "left_only",
            joined["_merge"] == "right_only",
            (joined["Trạng thái chuẩn tuần trước"] != joined["Trạng thái chuẩn tuần này"]),
        ],
        ["Mất khỏi file tuần này", "Mới phát sinh", "Chuyển trạng thái"],
        default="Không đổi",
    )
    transition = joined[joined["_merge"] == "both"].groupby(
        ["Trạng thái chuẩn tuần trước", "Trạng thái chuẩn tuần này"], dropna=False
    ).size().reset_index(name="Số lượng")
    status_prev = prev_no.groupby("Trạng thái chuẩn").size().rename("Tuần trước")
    status_curr = curr_no.groupby("Trạng thái chuẩn").size().rename("Tuần này")
    status_delta = pd.concat([status_prev, status_curr], axis=1).fillna(0).astype(int).reset_index(names="Trạng thái")
    status_delta["Delta"] = status_delta["Tuần này"] - status_delta["Tuần trước"]
    store_prev = prev_no[prev_no["Còn mở"]].groupby("Cửa hàng").size().rename("NO mở tuần trước")
    store_curr = curr_no[curr_no["Còn mở"]].groupby("Cửa hàng").size().rename("NO mở tuần này")
    store_delta = pd.concat([store_prev, store_curr], axis=1).fillna(0).astype(int).reset_index()
    store_delta["Delta"] = store_delta["NO mở tuần này"] - store_delta["NO mở tuần trước"]
    store_delta = store_delta.sort_values("Delta", ascending=False)
    group_prev = prev_no[prev_no["Còn mở"]].groupby("Nhóm công việc").size().rename("NO mở tuần trước")
    group_curr = curr_no[curr_no["Còn mở"]].groupby("Nhóm công việc").size().rename("NO mở tuần này")
    group_delta = pd.concat([group_prev, group_curr], axis=1).fillna(0).astype(int).reset_index()
    group_delta["Delta"] = group_delta["NO mở tuần này"] - group_delta["NO mở tuần trước"]
    group_delta = group_delta.sort_values("Delta", ascending=False)
    return joined, transition, status_delta, store_delta.merge(group_delta, how="outer", left_index=True, right_index=True)


def prepare_snapshot(no_file, mo_file, report_date: Optional[date], metadata: Optional[Dict[str, object]] = None) -> Dict[str, object]:
    raw_no = read_mht_table(no_file)
    raw_mo = read_mht_table(mo_file)
    no = normalize_no(raw_no, report_date=report_date)
    mo = normalize_mo(raw_mo)
    linked = link_no_mo(no, mo)
    investigation = build_investigation(linked)
    metrics = calculate_metrics(no, mo, linked)
    return {
        "no": no,
        "mo": mo,
        "linked": linked,
        "investigation": investigation,
        "metrics": metrics,
        "no_hash": file_hash(no_file),
        "mo_hash": file_hash(mo_file),
        "no_name": no_file.name,
        "mo_name": mo_file.name,
        "report_date": report_date,
        "metadata": metadata or {},
    }


# =========================
# Supabase persistence
# =========================

def get_supabase_config() -> Dict[str, str]:
    config: Dict[str, str] = {}
    try:
        section = st.secrets.get("supabase", {})
        config["url"] = section.get("url", "")
        config["key"] = section.get("service_role_key", section.get("key", ""))
        config["bucket"] = section.get("bucket", DEFAULT_SUPABASE_BUCKET)
    except Exception:
        config = {"url": "", "key": "", "bucket": DEFAULT_SUPABASE_BUCKET}
    config["url"] = config.get("url") or os.getenv("SUPABASE_URL", "")
    config["key"] = config.get("key") or os.getenv("SUPABASE_SERVICE_ROLE_KEY", os.getenv("SUPABASE_KEY", ""))
    config["bucket"] = config.get("bucket") or os.getenv("SUPABASE_BUCKET", DEFAULT_SUPABASE_BUCKET)
    return config


def supabase_is_configured() -> bool:
    cfg = get_supabase_config()
    return create_client is not None and bool(cfg.get("url")) and bool(cfg.get("key"))


@st.cache_resource(show_spinner=False)
def get_supabase_client_cached(url: str, key: str):
    return create_client(url, key)


def get_supabase_client():
    cfg = get_supabase_config()
    if not supabase_is_configured():
        return None
    return get_supabase_client_cached(cfg["url"], cfg["key"])


def upload_file_to_supabase(uploaded_file, prefix: str) -> Dict[str, object]:
    client = get_supabase_client()
    cfg = get_supabase_config()
    if client is None:
        raise RuntimeError("Chưa cấu hình Supabase nên không thể lưu file lên cloud.")
    raw = uploaded_file.getvalue()
    stamp = now_vn().strftime("%Y%m%d_%H%M%S")
    path = f"snapshots/{stamp}/{prefix}_{safe_filename(uploaded_file.name)}"
    client.storage.from_(cfg["bucket"]).upload(
        path,
        raw,
        file_options={"content-type": "application/octet-stream"},
    )
    return {
        "path": path,
        "name": uploaded_file.name,
        "hash": hash_bytes(raw),
        "size_bytes": len(raw),
    }


def save_snapshot_to_supabase(curr_no_file, curr_mo_file, prev_no_file, prev_mo_file, report_date_input: date, label: str) -> Dict[str, object]:
    client = get_supabase_client()
    if client is None:
        raise RuntimeError("Chưa cấu hình Supabase.")

    uploaded_at = now_vn().isoformat()
    curr_no = upload_file_to_supabase(curr_no_file, "curr_no")
    curr_mo = upload_file_to_supabase(curr_mo_file, "curr_mo")
    prev_no = upload_file_to_supabase(prev_no_file, "prev_no") if prev_no_file else {}
    prev_mo = upload_file_to_supabase(prev_mo_file, "prev_mo") if prev_mo_file else {}

    row = {
        "uploaded_at": uploaded_at,
        "report_date": report_date_input.isoformat(),
        "label": label,
        "curr_no_name": curr_no.get("name"),
        "curr_mo_name": curr_mo.get("name"),
        "prev_no_name": prev_no.get("name"),
        "prev_mo_name": prev_mo.get("name"),
        "curr_no_path": curr_no.get("path"),
        "curr_mo_path": curr_mo.get("path"),
        "prev_no_path": prev_no.get("path"),
        "prev_mo_path": prev_mo.get("path"),
        "curr_no_hash": curr_no.get("hash"),
        "curr_mo_hash": curr_mo.get("hash"),
        "prev_no_hash": prev_no.get("hash"),
        "prev_mo_hash": prev_mo.get("hash"),
    }
    response = client.table("dashboard_snapshots").insert(row).execute()
    if not response.data:
        raise RuntimeError("Đã upload file nhưng không ghi được metadata snapshot vào Supabase.")
    return response.data[0]


def get_latest_snapshot_meta() -> Optional[Dict[str, object]]:
    client = get_supabase_client()
    if client is None:
        return None
    response = client.table("dashboard_snapshots").select("*").order("uploaded_at", desc=True).limit(1).execute()
    if response.data:
        return response.data[0]
    return None


def list_recent_snapshot_meta(limit: int = 10) -> List[Dict[str, object]]:
    client = get_supabase_client()
    if client is None:
        return []
    response = client.table("dashboard_snapshots").select("*").order("uploaded_at", desc=True).limit(limit).execute()
    return response.data or []


def download_supabase_file(path: str, name: str) -> MemoryUpload:
    client = get_supabase_client()
    cfg = get_supabase_config()
    if client is None:
        raise RuntimeError("Chưa cấu hình Supabase.")
    raw = client.storage.from_(cfg["bucket"]).download(path)
    if isinstance(raw, str):
        raw = raw.encode("utf-8")
    return MemoryUpload(name=name or os.path.basename(path), raw=bytes(raw))


def load_snapshot_files_from_supabase(meta: Dict[str, object]) -> Tuple[MemoryUpload, MemoryUpload, Optional[MemoryUpload], Optional[MemoryUpload]]:
    curr_no = download_supabase_file(meta["curr_no_path"], meta.get("curr_no_name") or "NO.MHT")
    curr_mo = download_supabase_file(meta["curr_mo_path"], meta.get("curr_mo_name") or "MO.MHT")
    prev_no = None
    prev_mo = None
    if meta.get("prev_no_path") and meta.get("prev_mo_path"):
        prev_no = download_supabase_file(meta["prev_no_path"], meta.get("prev_no_name") or "NO_prev.MHT")
        prev_mo = download_supabase_file(meta["prev_mo_path"], meta.get("prev_mo_name") or "MO_prev.MHT")
    return curr_no, curr_mo, prev_no, prev_mo


# =========================
# UI helpers
# =========================

def render_header():
    if LOGO_PATH.exists():
        st.markdown("<div class='logo-wrap'>", unsafe_allow_html=True)
        _, center, _ = st.columns([1.2, 2.4, 1.2])
        with center:
            st.image(str(LOGO_PATH), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='dashboard-title'>SAP NO/MO Weekly Dashboard</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='dashboard-subtitle'>Báo cáo tổng quan, biến động tuần và danh sách điều tra lệnh NO/MO</div>",
        unsafe_allow_html=True,
    )


def metric_card(col, key: str, value: int, detail_df: pd.DataFrame, delta: Optional[str] = None):
    with col:
        with st.container(border=True):
            st.markdown(f"<div class='metric-title'>{key}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{value:,}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-delta'>{'Delta: ' + delta if delta is not None else '&nbsp;'}</div>", unsafe_allow_html=True)
            file_name = safe_filename(key).lower() + ".xlsx"
            st.download_button(
                "⬇ Tải chi tiết",
                data=df_to_excel_bytes(detail_df, sheet_name=key),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_metric_{key}",
                use_container_width=True,
            )
            st.markdown("<div class='metric-card-note'>Tải danh sách dòng cấu thành chỉ số này</div>", unsafe_allow_html=True)


def kpi_grid(snapshot: Dict[str, object], prev_snapshot: Optional[Dict[str, object]] = None):
    metrics = snapshot["metrics"]
    prev_metrics = prev_snapshot["metrics"] if prev_snapshot else None
    details = build_metric_details(snapshot)
    keys = [
        "Tổng NO",
        "NO còn mở",
        "NO đã đóng",
        "NO khởi tạo",
        "NO hoàn thành",
        "NO chưa có lệnh",
        "NO có lệnh nhưng thiếu MO",
        "MO tổng",
        "MO từ sự cố",
        "MO từ kế hoạch",
    ]
    for start in range(0, len(keys), 5):
        cols = st.columns(5)
        for i, key in enumerate(keys[start:start + 5]):
            delta = None
            if prev_metrics and key in prev_metrics:
                d = metrics.get(key, 0) - prev_metrics.get(key, 0)
                delta = f"{d:+,}" if d != 0 else "0"
            metric_card(cols[i], key, int(metrics.get(key, 0)), details.get(key, pd.DataFrame()), delta)


def bar_chart(df: pd.DataFrame, x: str, y: str, title: str, orientation="v", height=420):
    if df.empty:
        st.info("Không có dữ liệu để vẽ biểu đồ.")
        return
    if orientation == "h":
        fig = px.bar(df, x=y, y=x, orientation="h", title=title, text=y)
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=height)
    else:
        fig = px.bar(df, x=x, y=y, title=title, text=y)
        fig.update_layout(height=height)
    fig.update_traces(textposition="outside")
    fig.update_layout(margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig, use_container_width=True)


def download_excel_button(snapshot: Dict[str, object], comparison: Optional[Dict[str, pd.DataFrame]] = None):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        pd.DataFrame([snapshot["metrics"]]).to_excel(writer, sheet_name="KPI", index=False)
        snapshot["investigation"].to_excel(writer, sheet_name="Dieu_tra", index=False)
        snapshot["linked"].to_excel(writer, sheet_name="Lien_ket_NO_MO", index=False)
        snapshot["no"].to_excel(writer, sheet_name="NO_chuan_hoa", index=False)
        snapshot["mo"].to_excel(writer, sheet_name="MO_chuan_hoa", index=False)
        if comparison:
            for name, df in comparison.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
    st.download_button(
        "⬇️ Tải Excel báo cáo đầy đủ",
        data=buffer.getvalue(),
        file_name="bao_cao_NO_MO_hang_tuan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )


def show_data_quality(curr: Dict[str, object], prev: Optional[Dict[str, object]] = None):
    warnings = []
    no = curr["no"]
    mo = curr["mo"]
    if no["Mã sự cố"].duplicated().any():
        warnings.append(f"NO có {int(no['Mã sự cố'].duplicated().sum())} dòng trùng Mã sự cố.")
    if mo["Lệnh sửa chữa"].duplicated().any():
        warnings.append(f"MO có {int(mo['Lệnh sửa chữa'].duplicated().sum())} dòng trùng Lệnh sửa chữa.")
    if prev and curr["mo_hash"] == prev["mo_hash"]:
        warnings.append("File MO tuần này giống hệt file MO tuần trước. Cần kiểm tra lại export ZTC MO nếu kỳ báo cáo đã thay đổi.")
    missing_mo = int((curr["linked"]["Phân loại liên kết"] == "NO có lệnh nhưng thiếu trong MO").sum())
    if missing_mo:
        warnings.append(f"Có {missing_mo} NO có Lệnh sửa chữa nhưng không tìm thấy MO khớp trong file MO.")
    if warnings:
        st.markdown("<div class='warning-box'><b>Cảnh báo dữ liệu:</b><br>" + "<br>".join(f"• {w}" for w in warnings) + "</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='ok-box'><b>Dữ liệu hợp lệ:</b> chưa phát hiện cảnh báo lớn.</div>", unsafe_allow_html=True)


def render_saved_meta_box(snapshot: Dict[str, object]):
    meta = snapshot.get("metadata", {}) or {}
    report_date = snapshot.get("report_date") or meta.get("report_date")
    uploaded_at = meta.get("uploaded_at")
    label = meta.get("label") or "Bộ dữ liệu đang xem"
    if meta:
        st.markdown(
            f"""
            <div class='saved-box'>
                <b>{label}</b><br>
                • Tính đến ngày: <b>{format_vn_date(report_date)}</b><br>
                • Cập nhật lúc: <b>{format_vn_datetime(uploaded_at)}</b><br>
                • File hiện tại: <b>{meta.get('curr_no_name', snapshot.get('no_name'))}</b> + <b>{meta.get('curr_mo_name', snapshot.get('mo_name'))}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.caption(f"Tính đến ngày {format_vn_date(report_date)} | File hiện tại: {snapshot['no_name']} + {snapshot['mo_name']}")


def render_overview(snapshot: Dict[str, object], prev_snapshot: Optional[Dict[str, object]] = None):
    report_date = snapshot.get("report_date") or (snapshot.get("metadata", {}) or {}).get("report_date")
    st.markdown(f"<div class='section-title'>1. Tổng quan tuần hiện tại <span style='font-size:1.05rem; color:#475569;'>(tính đến ngày {format_vn_date(report_date)})</span></div>", unsafe_allow_html=True)
    kpi_grid(snapshot, prev_snapshot)
    st.divider()
    no = snapshot["no"]
    mo = snapshot["mo"]
    c1, c2 = st.columns(2)
    with c1:
        status_counts = no.groupby("Trạng thái chuẩn").size().reset_index(name="Số lượng")
        status_counts["Trạng thái chuẩn"] = pd.Categorical(status_counts["Trạng thái chuẩn"], categories=STATUS_ORDER, ordered=True)
        status_counts = status_counts.sort_values("Trạng thái chuẩn")
        bar_chart(status_counts, "Trạng thái chuẩn", "Số lượng", "NO theo trạng thái")
    with c2:
        source_counts = mo.groupby("Nguồn lệnh").size().reset_index(name="Số lượng")
        bar_chart(source_counts, "Nguồn lệnh", "Số lượng", "MO theo nguồn lệnh")
    c3, c4 = st.columns(2)
    with c3:
        top_store = no[no["Còn mở"]].groupby("Cửa hàng").size().reset_index(name="NO mở").sort_values("NO mở", ascending=False).head(12)
        bar_chart(top_store, "Cửa hàng", "NO mở", "Top cửa hàng có NO mở", orientation="h")
    with c4:
        top_group = no[no["Còn mở"]].groupby("Nhóm công việc").size().reset_index(name="NO mở").sort_values("NO mở", ascending=False).head(12)
        bar_chart(top_group, "Nhóm công việc", "NO mở", "Top nhóm công việc có NO mở", orientation="h")
    c5, c6 = st.columns(2)
    with c5:
        aging = no[no["Còn mở"]].groupby("Nhóm tuổi tồn").size().reset_index(name="Số lượng")
        bar_chart(aging, "Nhóm tuổi tồn", "Số lượng", "Aging NO còn mở")
    with c6:
        link_counts = snapshot["linked"].groupby("Phân loại liên kết").size().reset_index(name="Số lượng")
        bar_chart(link_counts, "Phân loại liên kết", "Số lượng", "Tình trạng liên kết NO-MO")


def render_compare(prev_snapshot: Dict[str, object], curr_snapshot: Dict[str, object]):
    st.markdown("<div class='section-title'>2. So sánh với tuần trước</div>", unsafe_allow_html=True)
    joined, transition, status_delta, mixed_delta = compare_snapshots(prev_snapshot["no"], curr_snapshot["no"])
    new_count = int((joined["Loại biến động"] == "Mới phát sinh").sum())
    changed_count = int((joined["Loại biến động"] == "Chuyển trạng thái").sum())
    unchanged_open = joined[
        (joined["Loại biến động"] == "Không đổi")
        & (joined["Còn mở tuần này"].fillna(False))
    ]
    closed_this_week = joined[
        (joined["Nhóm trạng thái tuần trước"] != "CLOSED")
        & (joined["Nhóm trạng thái tuần này"] == "CLOSED")
    ]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("NO mới phát sinh", f"{new_count:,}")
    c2.metric("NO chuyển trạng thái", f"{changed_count:,}")
    c3.metric("NO đóng trong tuần", f"{len(closed_this_week):,}")
    c4.metric("NO mở không đổi", f"{len(unchanged_open):,}")
    c5, c6 = st.columns(2)
    with c5:
        st.markdown("**Delta theo trạng thái**")
        st.dataframe(status_delta.sort_values("Delta", ascending=False), use_container_width=True, hide_index=True)
    with c6:
        st.markdown("**Ma trận chuyển trạng thái**")
        st.dataframe(transition.sort_values("Số lượng", ascending=False), use_container_width=True, hide_index=True)
    store_prev = prev_snapshot["no"][prev_snapshot["no"]["Còn mở"]].groupby("Cửa hàng").size().rename("NO mở tuần trước")
    store_curr = curr_snapshot["no"][curr_snapshot["no"]["Còn mở"]].groupby("Cửa hàng").size().rename("NO mở tuần này")
    store_delta = pd.concat([store_prev, store_curr], axis=1).fillna(0).astype(int).reset_index()
    store_delta["Delta"] = store_delta["NO mở tuần này"] - store_delta["NO mở tuần trước"]
    group_prev = prev_snapshot["no"][prev_snapshot["no"]["Còn mở"]].groupby("Nhóm công việc").size().rename("NO mở tuần trước")
    group_curr = curr_snapshot["no"][curr_snapshot["no"]["Còn mở"]].groupby("Nhóm công việc").size().rename("NO mở tuần này")
    group_delta = pd.concat([group_prev, group_curr], axis=1).fillna(0).astype(int).reset_index()
    group_delta["Delta"] = group_delta["NO mở tuần này"] - group_delta["NO mở tuần trước"]
    c7, c8 = st.columns(2)
    with c7:
        bar_chart(store_delta.sort_values("Delta", ascending=False).head(12), "Cửa hàng", "Delta", "Top cửa hàng tăng NO mở", orientation="h")
        st.dataframe(store_delta.sort_values("Delta", ascending=False), use_container_width=True, hide_index=True)
    with c8:
        bar_chart(group_delta.sort_values("Delta", ascending=False).head(12), "Nhóm công việc", "Delta", "Top nhóm công việc tăng NO mở", orientation="h")
        st.dataframe(group_delta.sort_values("Delta", ascending=False), use_container_width=True, hide_index=True)
    with st.expander("Xem danh sách chi tiết biến động từng NO"):
        display_cols = [
            "Mã sự cố",
            "Trạng thái chuẩn tuần trước",
            "Trạng thái chuẩn tuần này",
            "Cửa hàng tuần này",
            "Nhóm công việc tuần này",
            "Loại biến động",
        ]
        st.dataframe(joined[[c for c in display_cols if c in joined.columns]], use_container_width=True, hide_index=True)
    return {
        "Bien_dong_NO": joined,
        "Ma_tran_trang_thai": transition,
        "Delta_trang_thai": status_delta,
        "Delta_cua_hang": store_delta,
        "Delta_nhom_cong_viec": group_delta,
    }


def render_investigation(snapshot: Dict[str, object]):
    st.markdown("<div class='section-title'>3. Điều tra chi tiết</div>", unsafe_allow_html=True)
    inv = snapshot["investigation"].copy()
    if inv.empty:
        st.success("Không có dòng điều tra đáng chú ý.")
        return
    c1, c2, c3, c4 = st.columns(4)
    store_filter = c1.multiselect("Cửa hàng", sorted(inv["Cửa hàng"].dropna().unique().tolist())) if "Cửa hàng" in inv else []
    status_filter = c2.multiselect("Trạng thái", sorted(inv["Trạng thái chuẩn"].dropna().unique().tolist())) if "Trạng thái chuẩn" in inv else []
    group_filter = c3.multiselect("Nhóm công việc", sorted(inv["Nhóm công việc"].dropna().unique().tolist())) if "Nhóm công việc" in inv else []
    min_age = c4.number_input("Tuổi tồn tối thiểu", min_value=0, max_value=365, value=0, step=1)
    preset = st.radio(
        "Bộ lọc nhanh",
        [
            "Tất cả",
            "NO chưa có MO",
            "NO có MO nhưng chưa đóng",
            "Hoàn thành chưa đóng",
            "Tồn trên 14 ngày",
            "Tồn trên 30 ngày",
        ],
        horizontal=True,
    )
    filtered = inv.copy()
    if store_filter:
        filtered = filtered[filtered["Cửa hàng"].isin(store_filter)]
    if status_filter:
        filtered = filtered[filtered["Trạng thái chuẩn"].isin(status_filter)]
    if group_filter:
        filtered = filtered[filtered["Nhóm công việc"].isin(group_filter)]
    if "Tuổi tồn ngày" in filtered:
        filtered = filtered[filtered["Tuổi tồn ngày"].fillna(0) >= min_age]
    if preset == "NO chưa có MO":
        filtered = filtered[filtered["Cờ điều tra"].str.contains("chưa sinh lệnh", case=False, na=False)]
    elif preset == "NO có MO nhưng chưa đóng":
        filtered = filtered[filtered["Cờ điều tra"].str.contains("Đã có MO", case=False, na=False)]
    elif preset == "Hoàn thành chưa đóng":
        filtered = filtered[filtered["Cờ điều tra"].str.contains("hoàn thành nhưng chưa đóng", case=False, na=False)]
    elif preset == "Tồn trên 14 ngày":
        filtered = filtered[filtered["Tuổi tồn ngày"].fillna(0) > 14]
    elif preset == "Tồn trên 30 ngày":
        filtered = filtered[filtered["Tuổi tồn ngày"].fillna(0) > 30]
    st.caption(f"Đang hiển thị {len(filtered):,} / {len(inv):,} dòng điều tra")
    st.dataframe(filtered, use_container_width=True, hide_index=True, height=560)
    st.download_button(
        "⬇ Tải bảng điều tra đang lọc",
        data=df_to_excel_bytes(filtered, "Dieu_tra_dang_loc"),
        file_name="bang_dieu_tra_dang_loc.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_raw(snapshot: Dict[str, object]):
    st.markdown("<div class='section-title'>4. Dữ liệu chuẩn hóa</div>", unsafe_allow_html=True)
    data_choice = st.selectbox("Chọn bảng", ["NO chuẩn hóa", "MO chuẩn hóa", "Liên kết NO-MO", "Điều tra"])
    mapping = {
        "NO chuẩn hóa": snapshot["no"],
        "MO chuẩn hóa": snapshot["mo"],
        "Liên kết NO-MO": snapshot["linked"],
        "Điều tra": snapshot["investigation"],
    }
    st.dataframe(mapping[data_choice], use_container_width=True, hide_index=True, height=620)


# =========================
# Main app
# =========================
render_header()

configured = supabase_is_configured()
latest_meta = None
recent_metas = []
if configured:
    try:
        latest_meta = get_latest_snapshot_meta()
        recent_metas = list_recent_snapshot_meta(10)
    except Exception as exc:
        st.warning(f"Đã cấu hình Supabase nhưng chưa đọc được snapshot đã lưu: {exc}")
else:
    st.warning("Chưa cấu hình Supabase. App vẫn chạy được bằng upload trực tiếp, nhưng chưa thể lưu dữ liệu để sếp mở link là thấy dashboard.")

with st.sidebar:
    st.header("Dữ liệu báo cáo")
    if configured:
        st.success("Đã kết nối Supabase để lưu dữ liệu cloud.")
        if latest_meta:
            st.caption(f"Bản mới nhất: {format_vn_datetime(latest_meta.get('uploaded_at'))}")
            st.caption(f"Tính đến: {format_vn_date(latest_meta.get('report_date'))}")
        mode = st.radio(
            "Chế độ",
            ["Xem dữ liệu đã lưu mới nhất", "Upload/Cập nhật dữ liệu mới"],
            index=0 if latest_meta else 1,
        )
        if recent_metas:
            with st.expander("Các lần cập nhật gần đây"):
                for m in recent_metas:
                    st.write(f"• {format_vn_datetime(m.get('uploaded_at'))} | tính đến {format_vn_date(m.get('report_date'))}")
    else:
        mode = "Upload/Cập nhật dữ liệu mới"

curr_snapshot = None
prev_snapshot = None
comparison_outputs = None

if mode == "Xem dữ liệu đã lưu mới nhất":
    if not latest_meta:
        st.info("Chưa có bộ dữ liệu nào được lưu. Hãy chuyển sang Upload/Cập nhật dữ liệu mới để lưu bộ đầu tiên.")
        st.stop()
    try:
        curr_no_file, curr_mo_file, prev_no_file, prev_mo_file = load_snapshot_files_from_supabase(latest_meta)
        report_date_input = parse_report_date(latest_meta.get("report_date"))
        curr_snapshot = prepare_snapshot(curr_no_file, curr_mo_file, report_date_input, metadata=latest_meta)
        if prev_no_file and prev_mo_file:
            prev_snapshot = prepare_snapshot(prev_no_file, prev_mo_file, report_date_input, metadata=latest_meta)
    except Exception as exc:
        st.error(f"Không tải/đọc được dữ liệu đã lưu trên cloud: {exc}")
        st.stop()
else:
    with st.sidebar:
        st.markdown("**Tuần hiện tại**")
        curr_no_file = st.file_uploader("NO tuần hiện tại", type=["mht", "html", "xls"], key="curr_no")
        curr_mo_file = st.file_uploader("MO tuần hiện tại", type=["mht", "html", "xls"], key="curr_mo")
        st.markdown("**Tuần trước, tùy chọn**")
        prev_no_file = st.file_uploader("NO tuần trước", type=["mht", "html", "xls"], key="prev_no")
        prev_mo_file = st.file_uploader("MO tuần trước", type=["mht", "html", "xls"], key="prev_mo")
        report_date_input = st.date_input("Ngày chốt báo cáo", value=date.today())
        label = st.text_input("Tên kỳ báo cáo", value=f"Báo cáo NO/MO tính đến {format_vn_date(report_date_input)}")
        save_to_cloud = st.checkbox("Lưu bộ file này lên cloud để sếp mở link là thấy dashboard", value=configured, disabled=not configured)
        run_btn = st.button("🚀 Phân tích & cập nhật dashboard", use_container_width=True)

    if not curr_no_file or not curr_mo_file:
        st.info("Hãy upload tối thiểu 2 file: NO tuần hiện tại và MO tuần hiện tại.")
        st.stop()
    if not run_btn:
        st.info("Upload file xong thì bấm 🚀 Phân tích & cập nhật dashboard để chạy báo cáo.")
        st.stop()

    try:
        curr_snapshot = prepare_snapshot(curr_no_file, curr_mo_file, report_date_input)
        if prev_no_file and prev_mo_file:
            prev_snapshot = prepare_snapshot(prev_no_file, prev_mo_file, report_date_input)
        if save_to_cloud:
            saved_meta = save_snapshot_to_supabase(curr_no_file, curr_mo_file, prev_no_file, prev_mo_file, report_date_input, label)
            curr_snapshot["metadata"] = saved_meta
            if prev_snapshot:
                prev_snapshot["metadata"] = saved_meta
            st.success(f"Đã lưu dữ liệu lên cloud. Sếp mở link sẽ thấy bản cập nhật lúc {format_vn_datetime(saved_meta.get('uploaded_at'))}.")
    except Exception as exc:
        st.error(f"Không đọc/lưu được file: {exc}")
        st.stop()

render_saved_meta_box(curr_snapshot)
show_data_quality(curr_snapshot, prev_snapshot)

tab1, tab2, tab3, tab4 = st.tabs(["Tổng quan", "So sánh tuần", "Điều tra", "Dữ liệu gốc"])
with tab1:
    render_overview(curr_snapshot, prev_snapshot)
with tab2:
    if prev_snapshot:
        comparison_outputs = render_compare(prev_snapshot, curr_snapshot)
    else:
        st.info("Bộ dữ liệu này chưa có cặp NO/MO tuần trước, nên chưa có dashboard so sánh và delta.")
with tab3:
    render_investigation(curr_snapshot)
with tab4:
    render_raw(curr_snapshot)

st.divider()
download_excel_button(curr_snapshot, comparison_outputs)
