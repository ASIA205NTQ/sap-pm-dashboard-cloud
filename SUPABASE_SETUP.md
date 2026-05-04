# Thiết lập lưu dữ liệu cloud bằng Supabase cho Streamlit

## 1. Tạo Supabase project

1. Vào https://supabase.com
2. Tạo project mới, ví dụ: `sap-pm-dashboard`
3. Chờ project khởi tạo xong.

## 2. Tạo Storage bucket

Vào **Storage** → **New bucket**:

- Bucket name: `sap-files`
- Public bucket: tắt / false

## 3. Tạo bảng metadata

Vào **SQL Editor** → **New query**, chạy SQL sau:

```sql
create extension if not exists "pgcrypto";

create table if not exists public.dashboard_snapshots (
  id uuid primary key default gen_random_uuid(),
  uploaded_at timestamptz not null default now(),
  report_date date not null,
  label text,

  curr_no_name text,
  curr_mo_name text,
  prev_no_name text,
  prev_mo_name text,

  curr_no_path text not null,
  curr_mo_path text not null,
  prev_no_path text,
  prev_mo_path text,

  curr_no_hash text,
  curr_mo_hash text,
  prev_no_hash text,
  prev_mo_hash text,

  created_at timestamptz not null default now()
);

create index if not exists idx_dashboard_snapshots_uploaded_at
on public.dashboard_snapshots (uploaded_at desc);
```

## 4. Lấy URL và service role key

Vào **Project Settings** → **API**:

- Copy `Project URL`
- Copy `service_role` key

Không commit key này lên GitHub.

## 5. Thêm secret vào Streamlit Cloud

Vào app trên Streamlit Cloud:

**Manage app** → **Settings** → **Secrets**

Dán nội dung:

```toml
[supabase]
url = "https://YOUR_PROJECT_ID.supabase.co"
service_role_key = "YOUR_SUPABASE_SERVICE_ROLE_KEY"
bucket = "sap-files"
```

Bấm Save, sau đó reboot/redeploy app.

## 6. Dùng app

1. Mở app.
2. Sidebar sẽ báo đã kết nối Supabase.
3. Chọn `Upload/Cập nhật dữ liệu mới`.
4. Upload 2 file hiện tại hoặc 4 file nếu muốn so sánh tuần.
5. Bấm `Phân tích & cập nhật dashboard`.
6. Nếu bật checkbox lưu cloud, bộ file sẽ được lưu lên Supabase Storage và metadata lưu trong bảng `dashboard_snapshots`.
7. Lần sau sếp mở link, app mặc định chọn `Xem dữ liệu đã lưu mới nhất` và hiện dashboard mới nhất.

## 7. Chạy local với Supabase

Tạo file `.streamlit/secrets.toml` trên máy local theo format ở trên.

File `.streamlit/secrets.toml` đã được `.gitignore`, không được commit lên GitHub.
