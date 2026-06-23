# 🎯 AI Lead Scoring Center

Ung dung Streamlit cham diem va phan loai khach hang tiem nang nganh **Bat Dong San** bang AI ket hop kiem duyet thu cong.

## ✨ Tinh nang

- **Tai du lieu tu dong** tu Google Sheets hoac file CSV/Excel cuc bo
- **Cham diem AI** theo tieu chi VIP / Binh thuong / Rac (0–100 diem)
- **Kiem duyet thu cong** (Human-in-the-loop) truc tiep tren bang
- **Xuat bao cao Excel** chuyen nghiep voi Dashboard + Bieu do
- **Ghi chu kiem duyet** va dong bo diem cuoi cung

## 🚀 Chay ung dung

### Cai dat phu thuoc

```bash
pip install -r requirements.txt
```

### Khoi dong Streamlit

```bash
streamlit run app_lead_scoring.py
```

Sau do mo trinh duyet tai: `http://localhost:8501`

## 📁 Cau truc du an

```
.
├── app_lead_scoring.py       # Streamlit app chinh
├── score_leads.py            # Logic cham diem va xuat Excel
├── requirements.txt          # Thu vien Python can thiet
├── tieu_chi_cham_diem.txt    # Tieu chi cham diem goc
├── skills/
│   └── lead-scoring/
│       ├── SKILL.md
│       └── scripts/
│           └── score_leads.py  # Ban du phong trong skill
└── .devcontainer/
    └── devcontainer.json     # Cau hinh GitHub Codespaces
```

## 🏗️ Tieu chi cham diem

| Phan loai | Diem | Dieu kien |
|-----------|------|-----------|
| **VIP**        | 100 | Co tu khoa VIP (ngan sach >= 20 ty, loai hinh cao cap, vi tri dac dia, v.v.) |
| **Binh thuong** | 50  | Khach hang tieu chuan, khong co dau hieu VIP hay Rac |
| **Rac**         | 0   | Co tu khoa tieu cuc (nham so, khong nhu cau, spam, v.v.) |

## 🌐 Deploy tren Streamlit Cloud

1. Dang nhap vao [streamlit.io](https://streamlit.io)
2. Ket noi GitHub repository nay
3. Chon file `app_lead_scoring.py`
4. Deploy!

## 📊 Nguon du lieu mac dinh

Google Sheet: [Link du lieu](https://docs.google.com/spreadsheets/d/1EaHaNMNmqz2Yy-3DpaNktbi4Ii0vRvaz4lwjpL71zYg/edit?usp=sharing)
