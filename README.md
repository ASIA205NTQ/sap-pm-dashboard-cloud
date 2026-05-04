# SAP PM NO/MO Weekly Dashboard

Streamlit dashboard để upload, chuẩn hóa, so sánh và điều tra dữ liệu SAP ZTC NO/ZTC MO.

## Chạy local

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Deploy Streamlit Cloud

- Repository: repo GitHub chứa project này
- Branch: `main`
- Main file path: `app.py`

## Lưu dữ liệu cloud

Bản này hỗ trợ lưu bộ file đã upload lên Supabase để sếp mở link là thấy dashboard mới nhất. Xem file `SUPABASE_SETUP.md` để thiết lập Supabase và Streamlit Secrets.

## Không commit dữ liệu SAP thật

`.gitignore` đã chặn các file `.MHT`, `.xlsx`, `.xls`, `.csv`, `.venv`, cache và `.streamlit/secrets.toml`.
