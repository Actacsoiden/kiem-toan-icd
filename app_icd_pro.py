# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import re
import unicodedata

# =====================================================================
# THIẾT LẬP MẬT KHẨU BẢO MẬT TRÊN ĐÁM MÂY
# =====================================================================
SECRET_PASSWORD = "admin" 

st.set_page_config(page_title="Cổng Kiểm Toán ICD-10 Pro", page_icon="🏥", layout="wide")

# =====================================================================
# 1. KHỞI TẠO VÀ ĐỒNG BỘ BỘ NHỚ TRẠNG THÁI (STATE ENGINE)
# =====================================================================
if 'current_view' not in st.session_state: st.session_state.current_view = "Kiểm toán BHYT"
if 'active_code' not in st.session_state: st.session_state.active_code = ""
if 'audit_input' not in st.session_state: st.session_state.audit_input = ""
if 'audited_codes' not in st.session_state: st.session_state.audited_codes = set()
if 'is_unlocked' not in st.session_state: st.session_state.is_unlocked = False

# Đón và nạp mã từ tham số URL Internet trước khi giao diện được dựng lên
if "code" in st.query_params:
    received_code = str(st.query_params["code"]).strip().upper()
    if received_code:
        st.session_state.audit_input = received_code
        st.session_state.active_code = received_code
        st.session_state.current_view = "Kiểm toán BHYT"
        st.query_params.clear() # Giải phóng thanh địa chỉ chống kẹt lệnh

# =====================================================================
# 2. CẤU HÌNH GIAO DIỆN HIỂN THỊ (UI & CSS)
# =====================================================================
st.markdown("""
    <meta name="google" content="notranslate">
    <style>
    body { font-family: Arial, Helvetica, sans-serif !important; }
    .main .block-container { padding-top: 2rem; }
    .info-card { background-color: #F8F9FA; padding: 20px; border-radius: 10px; border-left: 6px solid #1A73E8; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .alert-red { background-color: #FDEDED; border-left: 6px solid #D32F2F; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .alert-yellow { background-color: #FFF8E1; border-left: 6px solid #FFA000; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .alert-green { background-color: #E8F5E9; border-left: 6px solid #388E3C; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .alert-purple { background-color: #F3E5F5; border-left: 6px solid #8E24AA; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .alert-pink { background-color: #FCE4EC; border-left: 6px solid #E91E63; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .alert-blue { background-color: #E3F2FD; border-left: 6px solid #2196F3; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .who-guide { background-color: #FAFAFA; border: 1px solid #CFD8DC; padding: 20px; border-radius: 8px; border-left: 8px solid #455A64; margin-top: 20px; font-size: 1em; line-height: 1.6;}
    .recommend-box { background-color: #E0F7FA; padding: 15px; border-radius: 8px; border-left: 6px solid #00BCD4; margin-top: 25px; }
    .paywall-box { background-color: #FFF3E0; padding: 30px; border-radius: 12px; border: 2px solid #FF9800; text-align: center; margin-top: 20px;}
    ul.audit-list { margin-top: 10px; margin-bottom: 5px; padding-left: 20px; }
    ul.audit-list li { margin-bottom: 8px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

def navigate_to_audit(code):
    st.session_state.audit_input = code
    st.session_state.active_code = code
    st.session_state.current_view = "Kiểm toán BHYT"

def remove_vietnamese_accents(s):
    if pd.isna(s) or not isinstance(s, str): return ""
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').replace('đ', 'd').replace('Đ', 'D').lower()

def safe_get_str(row, col_name):
    if col_name in row and pd.notna(row[col_name]):
        return str(row[col_name]).strip()
    return ""

@st.cache_data(show_spinner="⚙️ Đang đồng bộ dữ liệu kiểm toán từ máy chủ...")
def init_enterprise_engine(file_source):
    try:
        df_raw = pd.read_excel(file_source, sheet_name="Bảng mã ICD10", header=2, engine='openpyxl', dtype=str)
        if str(df_raw.iloc[0]['MÃ BỆNH']).strip() == '18' or str(df_raw.iloc[0]['STT']).strip() == '1':
            df = df_raw.iloc[1:].reset_index(drop=True)
        else:
            df = df_raw.copy()
        df = df.dropna(subset=['MÃ BỆNH'])
        df = df[df['MÃ BỆNH'].str.strip() != '']
        
        cols_to_ffill = ['STT CHƯƠNG', 'TÊN CHƯƠNG', 'MÃ KHỐI', 'TÊN KHỐI', 'MÃ NHÓM BỆNH 3 KÝ TỰ', 'TÊN NHÓM BỆNH 3 KÝ TỰ']
        cols_exist = [c for c in cols_to_ffill if c in df.columns]
        if cols_exist: df[cols_exist] = df[cols_exist].ffill()
        
        df['MÃ_KEY'] = df['MÃ BỆNH'].astype(str).str.strip().str.upper()
        df['NHÓM_KEY'] = df['MÃ NHÓM BỆNH 3 KÝ TỰ'].astype(str).str.strip().str.upper()
        df['TÊN_SEARCH_KEY'] = df['TÊN BỆNH'].apply(remove_vietnamese_accents)
        
        exact_map = df.set_index('MÃ_KEY', drop=False).to_dict(orient='index')
        return df, exact_map
    except Exception as e:
        st.error(f"Lỗi cấu trúc danh mục tệp Excel: {e}")
        return None, None

def get_code_role_and_priority(row):
    if safe_get_str(row, 'CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG'): return 4, "⚰️ CHỈ MÃ HÓA TỬ VONG"
    if safe_get_str(row, 'MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH'): return 3, "🛑 CHỈ LÀM BỆNH KÈM THEO"
    if safe_get_str(row, 'MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH'): return 2, "⚠️ KHÔNG KHUYẾN KHÍCH BỆNH CHÍNH"
    return 1, "🟢 VỪA LÀM CHÍNH, VỪA LÀM PHỤ"

def parse_who_clinical_data(row):
    html_parts = []
    inc = safe_get_str(row, 'BAO GỒM')
    exc = safe_get_str(row, 'LOẠI TRỪ')
    note = safe_get_str(row, 'GHI CHÚ')
    who_guide = safe_get_str(row, 'HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019')
    
    if inc: html_parts.append(f'<div style="margin-bottom:12px;"><span style="color:#2E7D32; font-size: 1.05em;">🟢 <b>BAO GỒM (Includes):</b></span><br><span style="color:#333; padding-left:5px;">{inc}</span></div>')
    if exc: html_parts.append(f'<div style="margin-bottom:12px;"><span style="color:#C62828; font-size: 1.05em;">🔴 <b>LOẠI TRỪ (Excludes):</b></span><br><span style="color:#333; padding-left:5px;">{exc}</span></div>')
    if note: html_parts.append(f'<div style="margin-bottom:12px;"><span style="color:#0277BD; font-size: 1.05em;">📝 <b>GHI CHÚ LÂM SÀNG:</b></span><br><span style="color:#333; padding-left:5px;">{note}</span></div>')
    
    if who_guide:
        t = who_guide.replace("\n", "<br>")
        if not re.search(r'(?i)(Bao gồm|Loại trừ|Ghi chú|Lưu ý|Mã hóa kép)', t):
            if not html_parts: html_parts.append(f'<div style="margin-bottom:12px;"><span style="color:#2E7D32; font-size: 1.05em;">🟢 <b>ĐỊNH NGHĨA LÂM SÀNG:</b></span><br><span style="color:#333; padding-left:5px;">{t}</span></div>')
        else:
            t = re.sub(r'(?i)(Bao gồm:?)', r'<div style="margin-top:10px; margin-bottom:4px;"><span style="color:#2E7D32; font-size: 1.05em;">🟢 <b>BAO GỒM (Includes):</b></span></div>', t)
            t = re.sub(r'(?i)(-?\s*Loại trừ:?)', r'<div style="margin-top:10px; margin-bottom:4px;"><span style="color:#C62828; font-size: 1.05em;">🔴 <b>LOẠI TRỪ (Excludes):</b></span></div>', t)
            t = re.sub(r'(?i)(Ghi chú:?)', r'<div style="margin-top:10px; margin-bottom:4px;"><span style="color:#0277BD; font-size: 1.05em;">📝 <b>GHI CHÚ LÂM SÀNG:</b></span></div>', t)
            t = re.sub(r'(?i)(Lưu ý:?)', r'<div style="margin-top:10px; margin-bottom:4px;"><span style="color:#E65100; font-size: 1.05em;">📌 <b>LƯU Ý QUAN TRỌNG:</b></span></div>', t)
            t = re.sub(r'(?i)(Sử dụng mã bổ sung|Dùng thêm mã|Mã hóa kép|Mã hóa thêm:?)', r'<div style="margin-top:10px; margin-bottom:4px;"><span style="color:#6A1B9A; font-size: 1.05em;">➕ <b>QUY TẮC MÃ HÓA KÉP:</b></span></div>', t)
            html_parts.append(f'<div style="color:#333; line-height:1.6;">{t}</div>')
    return "".join(html_parts) if html_parts else "<i>(Không có chỉ dẫn lâm sàng đặc thù)</i>"

def format_search_who_table(row):
    parts = []
    inc = safe_get_str(row, 'BAO GỒM')
    exc = safe_get_str(row, 'LOẠI TRỪ')
    who_guide = safe_get_str(row, 'HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019')
    
    if inc: parts.append(f"🟢 Bao gồm: {inc}")
    if exc: parts.append(f"🔴 Loại trừ: {exc}")
    if who_guide:
        if not re.search(r'(?i)(Bao|Loại)', who_guide) and not parts: parts.append(f"🟢 Định nghĩa: {who_guide}")
        else:
            t = re.sub(r'(?i)(Bao gồm:?)', '🟢 Bao gồm: ', who_guide)
            t = re.sub(r'(?i)(-?\s*Loại trừ:?)', '🔴 Loại trừ: ', t)
            parts.append(t)
    return " | ".join(parts)

# Đường dẫn tệp danh mục Excel trên không gian lưu trữ đám mây
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
data_source = os.path.join(CURRENT_DIR, "icd.xlsx")

with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/medical-doctor.png", width=80)
    st.title("Chuyên gia ICD-10")
    if not os.path.exists(data_source):
        st.error("❌ Chưa tìm thấy cấu trúc tệp dữ liệu 'icd.xlsx' trên GitHub.")
        st.stop()
    df_master, exact_map = init_enterprise_engine(data_source)
    if df_master is not None: st.success(f"✅ Máy chủ trực tuyến: **{len(df_master):,}** mã bệnh.")
    else: st.stop()
    st.markdown("---")
    menu_options = ["Kiểm toán BHYT", "Từ điển Tìm kiếm"]
    st.session_state.current_view = st.radio("📍 **BẢNG ĐIỀU KHIỂN:**", menu_options, index=menu_options.index(st.session_state.current_view))
    st.markdown("---")
    if st.session_state.is_unlocked: st.success("🔓 CẤP QUYỀN TRUY CẬP: TOÀN VIỆN")
    else: st.info(f"🎁 Lượt dùng thử thiết bị: **{len(st.session_state.audited_codes)}/3 mã**")

# =====================================================================
# PHÂN HỆ KIỂM TOÁN LÂM SÀNG PHÁP LÝ BHYT
# =====================================================================
if st.session_state.current_view == "Kiểm toán BHYT":
    st.header("🔍 Thẩm định Pháp lý & Định hướng Bệnh án")
    
    # Ép điều khiển luồng hiển thị thuần túy qua State trung tâm
    raw_code = st.text_input("Nhập mã ICD-10 cần kiểm toán:", key="audit_input")
    search_code = re.sub(r'[^a-zA-Z0-9.]', '', raw_code).strip().upper()
    st.session_state.active_code = search_code 
    
    if search_code:
        record = exact_map.get(search_code)
        if not record: st.error(f"❌ Mã '{search_code}' không tồn tại trong danh mục Bộ Y tế.")
        else:
            show_paywall = False
            if not st.session_state.is_unlocked:
                if search_code not in st.session_state.audited_codes:
                    if len(st.session_state.audited_codes) >= 3: show_paywall = True
                    else: st.session_state.audited_codes.add(search_code)
            
            if show_paywall:
                st.markdown(f"""<div class="paywall-box"><h2 style="color:#E65100; margin-top:0;">🔒 HỆ THỐNG KIỂM TOÁN ĐANG BỊ KHÓA</h2><p style="font-size:16px;">Vui lòng nhập Mật khẩu để kích hoạt trạng thái liên thông không giới hạn.</p></div>""", unsafe_allow_html=True)
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    pwd_input = st.text_input("🔑 Nhập Mật Khẩu Bản Quyền:", type="password")
                    if st.button("Mở Khóa Phần Mềm", use_container_width=True, type="primary"):
                        if pwd_input == SECRET_PASSWORD:
                            st.session_state.is_unlocked = True
                            st.session_state.audited_codes.add(search_code)
                            st.success("🎉 Hệ thống đã được kích hoạt trực tuyến thành công!")
                            st.rerun()
                        else: st.error("❌ Mật khẩu bảo mật không chính xác.")
            else:
                st.markdown(f"""<div class="info-card"><h2 style='color:#1A73E8; margin:0;'>{record['MÃ BỆNH']} — {record['TÊN BỆNH']}</h2></div>""", unsafe_allow_html=True)
                
                leaf_mask = df_master['MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'].isna()
                is_macro_locked = pd.notna(record.get('MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'))
                
                st.markdown("### ⚖️ Kết quả Giám định & Căn cứ Pháp lý xuất toán:")
                if is_macro_locked:
                    st.markdown("""<div class="alert-red"><h4 style="margin:0;">🛑 TỪ CHỐI THANH TOÁN: VI PHẠM MÃ VĨ MÔ (CỘT 26)</h4><ul class="audit-list"><li><b>Lý do:</b> Danh mục yêu cầu phân tách mức độ chi tiết cao hơn. Hãy chọn mã thay thế dưới đây:</li></ul></div>""", unsafe_allow_html=True)
                    sub_codes_df = df_master[(df_master['NHÓM_KEY'] == record['NHÓM_KEY']) & leaf_mask].copy()
                    if not sub_codes_df.empty:
                        sub_codes_df[['Mức Ưu Tiên', 'QUYỀN HẠN SỬ DỤNG']] = sub_codes_df.apply(lambda r: pd.Series(get_code_role_and_priority(r)), axis=1)
                        sub_codes_df = sub_codes_df.sort_values(by=['Mức Ưu Tiên', 'MÃ_KEY'])
                        display_subs = sub_codes_df[['MÃ BỆNH', 'QUYỀN HẠN SỬ DỤNG', 'TÊN BỆNH']].rename(columns={'MÃ BỆNH': 'Mã thay thế', 'TÊN BỆNH': 'Mô tả Lâm sàng'})
                        st.dataframe(display_subs, use_container_width=True, hide_index=True)
                else:
                    has_role_error = False
                    if safe_get_str(record, 'CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG'):
                        has_role_error = True; st.markdown("""<div class="alert-purple"><h4 style="margin:0;">⚰️ VI PHẠM CỘT 27: CHỈ SỬ DỤNG CHO NGUYÊN NHÂN TỬ VONG</h4></div>""", unsafe_allow_html=True)
                    if safe_get_str(record, 'MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH'):
                        has_role_error = True; st.markdown("""<div class="alert-red"><h4 style="margin:0;">🛑 VI PHẠM CỘT 24: CẤM SỬ DỤNG LÀM BỆNH CHÍNH</h4></div>""", unsafe_allow_html=True)
                    if safe_get_str(record, 'MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH'):
                        has_role_error = True; st.markdown("""<div class="alert-yellow"><h4 style="margin:0;">⚠️ CẢNH BÁO CỘT 25: KHÔNG KHUYẾN KHÍCH LÀM BỆNH CHÍNH</h4></div>""", unsafe_allow_html=True)
                    
                    if not has_role_error: 
                        st.markdown("""<div class="alert-green"><h4 style="margin:0;">🟢 QUYỀN HẠN CHẨN ĐOÁN: ĐẠT CHUẨN VAI TRÒ CHÍNH + PHỤ</h4></div>""", unsafe_allow_html=True)

                    if safe_get_str(record, 'CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NỮ GIỚI'):
                        st.markdown("""<div class="alert-pink"><h4 style="margin:0;">♀️ ĐIỀU KIỆN GIỚI TÍNH: CHỈ ĐƯỢC ÁP DỤNG CHO BỆNH NHÂN NỮ (Cột 28)</h4></div>""", unsafe_allow_html=True)
                    if safe_get_str(record, 'CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NAM GIỚI'):
                        st.markdown("""<div class="alert-blue"><h4 style="margin:0;">♂️ ĐIỀU KIỆN GIỚI TÍNH: CHỈ ĐƯỢC ÁP DỤNG CHO BỆNH NHÂN NAM (Cột 29)</h4></div>""", unsafe_allow_html=True)

                parsed_who_text = parse_who_clinical_data(record)
                st.markdown(f"""
                <div class="who-guide">
                    <h3 style="margin-top:0; color: #263238;">📚 CHỈ DẪN LÂM SÀNG & ĐỊNH VỊ PHÁP LÝ WHO</h3>
                    <b style="color:#1565C0;">1. Phân vị cấu trúc:</b> Chương {safe_get_str(record, 'STT CHƯƠNG')} | Khối {safe_get_str(record, 'MÃ KHỐI')} | Nhóm {safe_get_str(record, 'MÃ NHÓM BỆNH 3 KÝ TỰ')}<br>
                    <b style="color:#1565C0;">2. Diễn giải y khoa chuẩn hóa:</b><br>
                    <div style="padding-left: 20px; border-left: 3px solid #90A4AE;">{parsed_who_text}</div>
                </div>
                """, unsafe_allow_html=True)

                current_block = record.get('MÃ KHỐI')
                current_group = record.get('NHÓM_KEY')
                related_df = pd.DataFrame()
                
                if is_macro_locked:
                    if pd.notna(current_block):
                        related_df = df_master[(df_master['MÃ KHỐI'] == current_block) & (df_master['NHÓM_KEY'] != current_group) & leaf_mask].copy()
                        rec_title, rec_desc = f"💡 GỢI Ý CHẨN ĐOÁN PHÂN BIỆT (Khối {current_block})", "Các hệ cơ quan lân cận:"
                else:
                    related_df = df_master[(df_master['NHÓM_KEY'] == current_group) & (df_master['MÃ_KEY'] != search_code) & leaf_mask].copy()
                    rec_title, rec_desc = f"💡 GỢI Ý ĐỊNH VỊ GIẢI PHẪU CHI TIẾT (Nhóm {current_group})", "Các tổn thương đặc hiệu khác cùng nhóm cơ quan:"
                    if related_df.empty and pd.notna(current_block):
                        related_df = df_master[(df_master['MÃ KHỐI'] == current_block) & (df_master['NHÓM_KEY'] != current_group) & leaf_mask].copy()
                        rec_title, rec_desc = f"💡 GỢI Ý CHẨN ĐOÁN PHÂN BIỆT (Khối {current_block})", "Các hệ cơ quan lân cận:"

                if not related_df.empty:
                    suggested_df = related_df.head(15).copy()
                    st.markdown(f"""<div class="recommend-box"><h4 style="margin:0; color:#00838F;">{rec_title}</h4><p style="margin-top:5px; margin-bottom:10px;">{rec_desc}</p></div>""", unsafe_allow_html=True)
                    suggested_df[['Mức Ưu Tiên', 'QUYỀN HẠN SỬ DỤNG']] = suggested_df.apply(lambda r: pd.Series(get_code_role_and_priority(r)), axis=1)
                    suggested_df = suggested_df.sort_values(by=['Mức Ưu Tiên', 'MÃ_KEY'])
                    suggested_df['Chỉ dẫn lâm sàng (WHO)'] = suggested_df.apply(format_search_who_table, axis=1)
                    display_related = suggested_df[['MÃ BỆNH', 'QUYỀN HẠN SỬ DỤNG', 'TÊN BỆNH', 'Chỉ dẫn lâm sàng (WHO)']].rename(columns={'TÊN BỆNH': 'Tên bệnh lý'})
                    st.dataframe(display_related, use_container_width=True, hide_index=True)

# =====================================================================
# PHÂN HỆ TỪ ĐIỂN TÌM KIẾM AN TOÀN TOÀN VIỆN
# =====================================================================
elif st.session_state.current_view == "Từ điển Tìm kiếm":
    st.header("📖 Bộ Lọc Bệnh Lý Hợp Lệ (Vùng An Toàn)")
    st.markdown("*Lọc bỏ hoàn toàn các mã cấm dùng cho chẩn đoán xác định bệnh chính.*")
    search_query = st.text_input("Gõ từ khóa thuật ngữ lâm sàng (Không dấu):")
    if search_query:
        query_normalized = remove_vietnamese_accents(search_query)
        result_df = df_master[df_master['TÊN_SEARCH_KEY'].str.contains(query_normalized, na=False)].copy()
        if result_df.empty: st.warning("Không tìm thấy thuật ngữ phù hợp.")
        else:
            valid_mask = result_df['MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'].isna()
            if 'CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG' in result_df.columns: valid_mask &= result_df['CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG'].isna()
            if 'MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH' in result_df.columns: valid_mask &= result_df['MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH'].isna()
            if 'MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH' in result_df.columns: valid_mask &= result_df['MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH'].isna()
            result_df = result_df[valid_mask].copy()

            if result_df.empty: st.warning("⚠️ Không có mã sạch nào đạt điều kiện làm bệnh chính.")
            else:
                st.success(f"🔍 Hệ thống đã bóc tách được **{len(result_df)}** chẩn đoán hợp lệ:")
                def get_status(row):
                    gender_flags = []
                    if safe_get_str(row, 'CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NỮ GIỚI'): gender_flags.append("♀️ Chỉ Nữ")
                    if safe_get_str(row, 'CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NAM GIỚI'): gender_flags.append("♂️ Chỉ Nam")
                    if gender_flags: return f"🟢 Chính + Phụ | " + " + ".join(gender_flags)
                    return "🟢 Chính + Phụ (Hợp Lệ Toàn Vẹn)"
                result_df['TRẠNG THÁI KIỂM TOÁN'] = result_df.apply(get_status, axis=1)
                result_df['Chỉ dẫn lâm sàng (WHO)'] = result_df.apply(format_search_who_table, axis=1)
                st.dataframe(result_df[['MÃ BỆNH', 'TRẠNG THÁI KIỂM TOÁN', 'TÊN BỆNH', 'Chỉ dẫn lâm sàng (WHO)']], use_container_width=True, hide_index=True)
                col_sel, col_btn = st.columns([3, 1])
                with col_sel: selected_code = st.selectbox("Chọn mã để nạp vào hệ thống giám định:", result_df['MÃ BỆNH'].tolist())
                with col_btn:
                    st.write("") 
                    if st.button("Kiểm toán mã này", use_container_width=True, type="primary"):
                        navigate_to_audit(selected_code)
                        st.rerun()