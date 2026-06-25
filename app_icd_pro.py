# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import re
import unicodedata

# =====================================================================
# THIẾT LẬP MẬT KHẨU BẢN QUYỀN
# =====================================================================
SECRET_PASSWORD = "admin" 

# =====================================================================
# 1. CẤU HÌNH GIAO DIỆN & STATE MANAGEMENT
# =====================================================================
st.set_page_config(page_title="Cổng Kiểm Toán ICD-10 Pro", page_icon="🏥", layout="wide")

# --- ĐỌC TÍN HIỆU TỪ CẦU NỐI F9 ---
code_file = "shared_code.txt"
if os.path.exists(code_file):
    with open(code_file, "r", encoding="utf-8") as f:
        new_code = f.read().strip()
    if new_code:
        st.session_state.active_code = new_code
        st.session_state.current_view = "Kiểm toán BHYT"
        if "audit_input" in st.session_state:
            st.session_state.audit_input = new_code
        open(code_file, "w").close()

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
    .who-guide { background-color: #F5F5F5; padding: 15px; border-radius: 8px; border-left: 6px solid #607D8B; margin-top: 15px; font-size: 0.95em;}
    .recommend-box { background-color: #E0F7FA; padding: 15px; border-radius: 8px; border-left: 6px solid #00BCD4; margin-top: 25px; }
    .paywall-box { background-color: #FFF3E0; padding: 30px; border-radius: 12px; border: 2px solid #FF9800; text-align: center; margin-top: 20px;}
    ul.audit-list { margin-top: 10px; margin-bottom: 5px; padding-left: 20px; }
    ul.audit-list li { margin-bottom: 8px; line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

if 'current_view' not in st.session_state:
    st.session_state.current_view = "Kiểm toán BHYT"
if 'active_code' not in st.session_state:
    st.session_state.active_code = ""
if 'audited_codes' not in st.session_state:
    st.session_state.audited_codes = set()
if 'is_unlocked' not in st.session_state:
    st.session_state.is_unlocked = False

def navigate_to_audit(code):
    st.session_state.active_code = code
    st.session_state.current_view = "Kiểm toán BHYT"

# =====================================================================
# 2. XỬ LÝ DỮ LIỆU & TIẾNG VIỆT
# =====================================================================
def remove_vietnamese_accents(s):
    if pd.isna(s) or not isinstance(s, str): return ""
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.replace('đ', 'd').replace('Đ', 'D').lower()

@st.cache_data(show_spinner="⚙️ Đang cấu trúc hệ thống dữ liệu Y khoa...")
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
        if cols_exist:
            df[cols_exist] = df[cols_exist].ffill()
            
        df['MÃ_KEY'] = df['MÃ BỆNH'].astype(str).str.strip().str.upper()
        df['NHÓM_KEY'] = df['MÃ NHÓM BỆNH 3 KÝ TỰ'].astype(str).str.strip().str.upper()
        df['TÊN_SEARCH_KEY'] = df['TÊN BỆNH'].apply(remove_vietnamese_accents)
        
        exact_map = df.set_index('MÃ_KEY', drop=False).to_dict(orient='index')
        return df, exact_map
    except Exception as e:
        st.error(f"Lỗi khởi tạo dữ liệu: {e}")
        return None, None

def get_code_role_and_priority(row):
    if 'CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG' in row and pd.notna(row['CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG']):
        return 4, "⚰️ CHỈ MÃ HÓA TỬ VONG"
    if pd.notna(row.get('MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH')):
        return 3, "🛑 CHỈ LÀM BỆNH KÈM THEO"
    if pd.notna(row.get('MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH')):
        return 2, "⚠️ KHÔNG KHUYẾN KHÍCH BỆNH CHÍNH"
    return 1, "🟢 VỪA LÀM CHÍNH, VỪA LÀM PHỤ"

# =====================================================================
# 3. QUẢN LÝ TỆP DỮ LIỆU BÊN TRÁI
# =====================================================================
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/medical-doctor.png", width=80)
    st.title("Chuyên gia ICD-10")
    st.markdown("---")
    
    default_filename = "icd.xlsx"
    absolute_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), default_filename)
    data_source = absolute_path if os.path.exists(absolute_path) else (default_filename if os.path.exists(default_filename) else None)
    
    if not data_source:
        uploaded_file = st.file_uploader("📂 Kéo thả file 'icd.xlsx' vào đây:", type=["xlsx"])
        if uploaded_file: data_source = uploaded_file

    if data_source:
        df_master, exact_map = init_enterprise_engine(data_source)
        if df_master is not None:
            st.success(f"✅ Đã tải: **{len(df_master):,}** mã bệnh.")
        else:
            st.stop()
    else:
        st.stop()

    st.markdown("---")
    menu_options = ["Kiểm toán BHYT", "Từ điển Tìm kiếm"]
    st.session_state.current_view = st.radio("📍 **BẢNG ĐIỀU KHIỂN:**", menu_options, index=menu_options.index(st.session_state.current_view))
    
    st.markdown("---")
    if st.session_state.is_unlocked:
        st.success("🔓 Phiên bản: PRO (Đã kết nối HIS)")
    else:
        st.info(f"🎁 Lượt tra cứu miễn phí: **{len(st.session_state.audited_codes)}/3 mã**")

# =====================================================================
# 4. PHÂN HỆ KIỂM TOÁN LÂM SÀNG & BẢO MẬT
# =====================================================================
if st.session_state.current_view == "Kiểm toán BHYT":
    st.header("🔍 Thẩm định Pháp lý & Định hướng Bệnh án")
    
    raw_code = st.text_input("Nhập mã ICD-10 (Gõ tay hoặc Ấn F9 từ phần mềm khác):", value=st.session_state.active_code, key="audit_input")
    search_code = re.sub(r'[^a-zA-Z0-9.]', '', raw_code).strip().upper()
    st.session_state.active_code = search_code 
    
    if search_code:
        record = exact_map.get(search_code)
        
        if not record:
            st.error(f"❌ Mã '{search_code}' không hợp lệ hoặc không có trong Danh mục.")
        else:
            show_paywall = False
            if not st.session_state.is_unlocked:
                if search_code not in st.session_state.audited_codes:
                    if len(st.session_state.audited_codes) >= 3:
                        show_paywall = True
                    else:
                        st.session_state.audited_codes.add(search_code)
            
            if show_paywall:
                st.markdown(f"""
                <div class="paywall-box">
                    <h2 style="color:#E65100; margin-top:0;">🔒 YÊU CẦU CẤP QUYỀN TRUY CẬP HỆ THỐNG</h2>
                    <p style="font-size:16px;">Bạn đang tra cứu mã thứ 4 (<b>{search_code}</b>).<br>
                    Vui lòng nhập Mật khẩu để kích hoạt tính năng tra cứu không giới hạn & Kết nối liên thông HIS.</p>
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    pwd_input = st.text_input("🔑 Nhập Mật Khẩu Bản Quyền:", type="password")
                    if st.button("Mở Khóa Phần Mềm", use_container_width=True, type="primary"):
                        if pwd_input == SECRET_PASSWORD:
                            st.session_state.is_unlocked = True
                            st.session_state.audited_codes.add(search_code)
                            st.success("🎉 Mở khóa thành công! Hệ thống hoạt động 100%.")
                            st.rerun()
                        else:
                            st.error("❌ Sai Mật khẩu! Vui lòng thử lại.")
            
            else:
                st.markdown(f"""
                <div class="info-card">
                    <h2 style='color:#1A73E8; margin:0;'>{record['MÃ BỆNH']} — {record['TÊN BỆNH']}</h2>
                    <hr style='margin: 10px 0;'>
                    <b>Nhóm:</b> {record['MÃ NHÓM BỆNH 3 KÝ TỰ']} - {record['TÊN NHÓM BỆNH 3 KÝ TỰ']}<br>
                    <b>Thuộc Khối:</b> {record.get('MÃ KHỐI', 'N/A')} - {record.get('TÊN KHỐI', 'N/A')}<br>
                    <b>Thuộc Chương:</b> {record['STT CHƯƠNG']} - {record['TÊN CHƯƠNG']}
                </div>
                """, unsafe_allow_html=True)
                
                leaf_mask = df_master['MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'].isna()
                is_macro_locked = pd.notna(record.get('MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'))
                
                st.markdown("### ⚖️ Kết quả Giám định & Căn cứ Pháp lý:")
                
                if is_macro_locked:
                    st.markdown("""
                    <div class="alert-red">
                        <h4 style="margin:0;">🛑 TỪ CHỐI THANH TOÁN: MÃ VĨ MÔ CHƯA ĐỦ CHI TIẾT</h4>
                        <ul class="audit-list">
                            <li><b>Lý do:</b> Vi phạm Cột 26. Hệ thống yêu cầu mức độ chi tiết cao hơn.</li>
                            <li><b>Hướng giải quyết:</b> Hãy chọn một trong các mã con hợp lệ dưới đây <i>(Ưu tiên làm Bệnh chính xếp trên cùng)</i>.</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    sub_codes_df = df_master[(df_master['NHÓM_KEY'] == record['NHÓM_KEY']) & leaf_mask].copy()
                    if not sub_codes_df.empty:
                        sub_codes_df[['Mức Ưu Tiên', 'QUYỀN HẠN SỬ DỤNG']] = sub_codes_df.apply(lambda r: pd.Series(get_code_role_and_priority(r)), axis=1)
                        sub_codes_df = sub_codes_df.sort_values(by=['Mức Ưu Tiên', 'MÃ_KEY'])
                        display_subs = sub_codes_df[['MÃ BỆNH', 'QUYỀN HẠN SỬ DỤNG', 'TÊN BỆNH']].rename(columns={'MÃ BỆNH': 'Mã thay thế', 'TÊN BỆNH': 'Mô tả Lâm sàng'})
                        st.dataframe(display_subs, use_container_width=True, hide_index=True)
                else:
                    has_error = False
                    if pd.notna(record.get('CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG')):
                        has_error = True
                        st.markdown("""<div class="alert-purple"><h4 style="margin:0;">⚰️ CHỈ SỬ DỤNG CHO TỬ VONG (Cột 27)</h4></div>""", unsafe_allow_html=True)
                    if pd.notna(record.get('MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH')):
                        has_error = True
                        st.markdown("""<div class="alert-red"><h4 style="margin:0;">🛑 CẤM LÀM BỆNH CHÍNH (Cột 24)</h4><ul class="audit-list"><li>Chỉ được đặt mã này ở ô <b>[Bệnh kèm theo]</b>.</li></ul></div>""", unsafe_allow_html=True)
                    if pd.notna(record.get('MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH')):
                        has_error = True
                        st.markdown("""<div class="alert-yellow"><h4 style="margin:0;">⚠️ KHÔNG KHUYẾN KHÍCH LÀM BỆNH CHÍNH (Cột 25)</h4></div>""", unsafe_allow_html=True)
                    if pd.notna(record.get('CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NỮ GIỚI')):
                        has_error = True
                        st.markdown("""<div class="alert-pink"><h4 style="margin:0;">♀️ MÃ BỆNH LÝ NỮ GIỚI (Cột 28)</h4></div>""", unsafe_allow_html=True)
                    if pd.notna(record.get('CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NAM GIỚI')):
                        has_error = True
                        st.markdown("""<div class="alert-blue"><h4 style="margin:0;">♂️ MÃ BỆNH LÝ NAM GIỚI (Cột 29)</h4></div>""", unsafe_allow_html=True)
                    
                    if not has_error:
                        st.markdown("""<div class="alert-green"><h4 style="margin:0;">🟢 ĐẠT CHUẨN TOÀN VẸN</h4><ul class="audit-list"><li>Được phép dùng tự do cho cả ô <b>Bệnh chính</b> hoặc <b>Bệnh kèm theo</b>.</li></ul></div>""", unsafe_allow_html=True)

                who_guide = record.get('HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019', None)
                if pd.notna(who_guide):
                    st.markdown(f"""
                    <div class="who-guide">
                        <b>📖 CHỈ DẪN LÂM SÀNG TỪ WHO TẠI MÃ NÀY:</b><br>
                        {str(who_guide).replace("Bao gồm:", "<br>🟢 <b>Bao gồm:</b>").replace("- Loại trừ:", "<br>🔴 <b>Loại trừ:</b>")}
                    </div>
                    """, unsafe_allow_html=True)

                # =========================================================
                # THUẬT TOÁN GỢI Ý ĐỊNH VỊ GIẢI PHẪU KÈM CHỈ DẪN WHO
                # =========================================================
                current_block = record.get('MÃ KHỐI')
                current_group = record.get('NHÓM_KEY')
                current_code = search_code 
                
                related_df = pd.DataFrame()
                
                if is_macro_locked:
                    if pd.notna(current_block):
                        related_df = df_master[(df_master['MÃ KHỐI'] == current_block) & (df_master['NHÓM_KEY'] != current_group) & leaf_mask].copy()
                        rec_title = f"💡 GỢI Ý CHẨN ĐOÁN PHÂN BIỆT (Khối {current_block})"
                        rec_desc = f"Các mã thuộc hệ <b>{record.get('TÊN KHỐI')}</b> có thể bạn quan tâm:"
                else:
                    related_df = df_master[(df_master['NHÓM_KEY'] == current_group) & (df_master['MÃ_KEY'] != current_code) & leaf_mask].copy()
                    rec_title = f"💡 GỢI Ý ĐỊNH VỊ GIẢI PHẪU CHI TIẾT HƠN (Nhóm {current_group})"
                    rec_desc = f"Bạn đang chọn mã <b>{current_code}</b>. Dưới đây là các vị trí giải phẫu/mức độ cụ thể khác thuộc nhóm <b>{record.get('TÊN NHÓM BỆNH 3 KÝ TỰ')}</b> mà bạn có thể cân nhắc:"
                    
                    if related_df.empty and pd.notna(current_block):
                        related_df = df_master[(df_master['MÃ KHỐI'] == current_block) & (df_master['NHÓM_KEY'] != current_group) & leaf_mask].copy()
                        rec_title = f"💡 GỢI Ý CHẨN ĐOÁN PHÂN BIỆT (Khối {current_block})"
                        rec_desc = f"Các mã thuộc hệ <b>{record.get('TÊN KHỐI')}</b> có thể bạn quan tâm:"

                if not related_df.empty:
                    suggested_df = related_df.head(15).copy()
                    st.markdown(f"""
                    <div class="recommend-box">
                        <h4 style="margin:0; color:#00838F;">{rec_title}</h4>
                        <p style="margin-top:5px; margin-bottom:10px;">{rec_desc}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    suggested_df[['Mức Ưu Tiên', 'QUYỀN HẠN SỬ DỤNG']] = suggested_df.apply(lambda r: pd.Series(get_code_role_and_priority(r)), axis=1)
                    suggested_df = suggested_df.sort_values(by=['Mức Ưu Tiên', 'MÃ_KEY'])
                    
                    # Thêm Cột Hướng dẫn Lâm sàng WHO vào Bảng Gợi ý
                    if 'HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019' in suggested_df.columns:
                        def format_who_guide(text):
                            if pd.isna(text): return ""
                            return str(text).replace("Bao gồm:", "🟢 Bao gồm:").replace("- Loại trừ:", "🔴 Loại trừ:")
                        
                        suggested_df['Chỉ dẫn Lâm sàng (WHO)'] = suggested_df['HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019'].apply(format_who_guide)
                        display_related = suggested_df[['MÃ BỆNH', 'QUYỀN HẠN SỬ DỤNG', 'TÊN BỆNH', 'Chỉ dẫn Lâm sàng (WHO)']].rename(columns={'TÊN BỆNH': 'Tên bệnh lý'})
                    else:
                        display_related = suggested_df[['MÃ BỆNH', 'QUYỀN HẠN SỬ DỤNG', 'TÊN BỆNH']].rename(columns={'TÊN BỆNH': 'Tên bệnh lý'})
                        
                    st.dataframe(display_related, use_container_width=True, hide_index=True)

# =====================================================================
# 5. PHÂN HỆ TỪ ĐIỂN TÌM KIẾM SẠCH (CLEAN SEARCH ENGINE)
# =====================================================================
elif st.session_state.current_view == "Từ điển Tìm kiếm":
    st.header("📖 Bộ Lọc Bệnh Lý Hợp Lệ (Vùng An Toàn)")
    st.markdown("*Lưu ý: Bộ lọc này đã được cấu hình tự động ẩn toàn bộ các mã vĩ mô, mã tử vong, mã cấm làm bệnh chính và mã không khuyến khích. Chỉ hiển thị các chẩn đoán đạt chuẩn thanh toán 100%.*")
    search_query = st.text_input("Gõ từ khóa bệnh lý (Gõ KHÔNG DẤU. VD: ruot thua, sinh de...):")
    
    if search_query:
        query_normalized = remove_vietnamese_accents(search_query)
        result_df = df_master[df_master['TÊN_SEARCH_KEY'].str.contains(query_normalized, na=False)].copy()
        
        if result_df.empty:
            st.warning("Không tìm thấy thuật ngữ phù hợp.")
        else:
            # --- MÀNG LỌC TUYỆT ĐỐI: CHỈ LẤY CÁC MÃ ĐẠT CHUẨN ---
            # 1. Loại bỏ Mã Vĩ mô (Cột 26)
            valid_mask = result_df['MÃ KHÔNG ĐƯỢC SỬ DỤNG VÌ CÓ MÃ 4 HOẶC 5 KÝ TỰ CỤ THỂ HƠN'].isna()
            
            # 2. Loại bỏ Mã Tử vong (Cột 27)
            if 'CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG' in result_df.columns:
                valid_mask &= result_df['CHỈ SỬ DỤNG MÃ HÓA NGUYÊN NHÂN TỬ VONG'].isna()
                
            # 3. Loại bỏ Mã Cấm Bệnh Chính (Cột 24)
            if 'MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH' in result_df.columns:
                valid_mask &= result_df['MÃ KHÔNG ĐƯỢC DÙNG LÀ BỆNH CHÍNH'].isna()
                
            # 4. Loại bỏ Mã Không Khuyến Khích Bệnh Chính (Cột 25)
            if 'MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH' in result_df.columns:
                valid_mask &= result_df['MÃ KHÔNG KHUYẾN KHÍCH DÙNG LÀ BỆNH CHÍNH'].isna()
            
            # Áp dụng bộ lọc
            result_df = result_df[valid_mask]

            if result_df.empty:
                st.warning("⚠️ Không có mã 'Hợp lệ toàn vẹn' nào khớp với từ khóa của bạn (Các mã tìm được đều vi phạm luật xuất toán hoặc là mã vĩ mô). Vui lòng gõ từ khóa chi tiết hơn.")
            else:
                st.success(f"🔍 Đã lọc thành công. Dưới đây là **{len(result_df)}** mã hoàn toàn hợp lệ để sử dụng làm bệnh án:")
                
                def get_status(row):
                    # Chỉ còn lại Hợp lệ toàn vẹn và Hợp lệ theo Giới tính
                    flags = []
                    if pd.notna(row.get('CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NỮ GIỚI')): flags.append("♀️ Bệnh Nữ (Hợp lệ nếu đúng giới tính)")
                    if pd.notna(row.get('CÁC MÃ BỆNH CHỈ CÓ HOẶC CHỦ YẾU CÓ Ở NAM GIỚI')): flags.append("♂️ Bệnh Nam (Hợp lệ nếu đúng giới tính)")
                    return " + ".join(flags) if flags else "🟢 Hợp Lệ Toàn Vẹn"
                    
                result_df['TRẠNG THÁI KIỂM TOÁN'] = result_df.apply(get_status, axis=1)
                
                # Hiển thị thêm Chỉ dẫn WHO nếu có để bác sĩ dễ chọn
                if 'HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019' in result_df.columns:
                    def format_search_who(text):
                        if pd.isna(text): return ""
                        return str(text).replace("Bao gồm:", "🟢 Bao gồm:").replace("- Loại trừ:", "🔴 Loại trừ:")
                    result_df['Chỉ dẫn (WHO)'] = result_df['HƯỚNG DẪN MÃ HÓA BỔ SUNG CỦA WHO 2019'].apply(format_search_who)
                    st.dataframe(result_df[['MÃ BỆNH', 'TRẠNG THÁI KIỂM TOÁN', 'TÊN BỆNH', 'Chỉ dẫn (WHO)']], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(result_df[['MÃ BỆNH', 'TRẠNG THÁI KIỂM TOÁN', 'TÊN BỆNH']], use_container_width=True, hide_index=True)
                
                col_sel, col_btn = st.columns([3, 1])
                with col_sel:
                    selected_code = st.selectbox("Chọn mã để xem luồng phân tích rủi ro chi tiết (Tính 1 lượt):", result_df['MÃ BỆNH'].tolist())
                with col_btn:
                    st.write("") 
                    if st.button("Kiểm toán mã này", use_container_width=True, type="primary"):
                        navigate_to_audit(selected_code)
                        st.rerun()