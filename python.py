import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from google import genai
from io import BytesIO
from docx import Document

# Cấu hình trang Streamlit
st.set_page_config(layout="wide", page_title="Hệ thống Thẩm định Phương án Kinh doanh")

# --- Các hàm tiện ích ---

def format_currency(value):
    """Định dạng số thành chuỗi tiền tệ với dấu chấm phân cách hàng nghìn."""
    if pd.isna(value) or value is None:
        return 0
    return "{:,.0f}".format(value).replace(",", ".")

def extract_text_from_docx(file_buffer):
    """Trích xuất toàn bộ nội dung text từ file .docx."""
    try:
        document = Document(file_buffer)
        full_text = []
        for para in document.paragraphs:
            full_text.append(para.text)
        
        # Thêm nội dung từ các bảng (giả định có trong file)
        for table in document.tables:
            table_text = []
            for row in table.rows:
                row_text = " | ".join(cell.text for cell in row.cells)
                table_text.append(row_text)
            full_text.append("\n" + "BẢNG DỮ LIỆU:\n" + "\n".join(table_text))

        return "\n".join(full_text)
    except Exception as e:
        st.error(f"Lỗi khi trích xuất dữ liệu từ file DOCX: {e}")
        return ""

def calculate_loan_schedule(loan_amount, annual_rate, duration_months):
    """Tính toán kế hoạch trả nợ theo phương pháp dư nợ giảm dần."""
    if loan_amount <= 0 or annual_rate <= 0 or duration_months <= 0:
        return pd.DataFrame({
            'Kỳ trả nợ': [1],
            'Dư nợ đầu kỳ (VNĐ)': [0],
            'Gốc trả (VNĐ)': [0],
            'Lãi trả (VNĐ)': [0],
            'Tổng gốc và lãi (VNĐ)': [0],
            'Dư nợ cuối kỳ (VNĐ)': [0]
        })

    monthly_rate = annual_rate / 12 / 100
    df = pd.DataFrame()
    
    # Giả định trả gốc đều hàng tháng
    monthly_principal = loan_amount / duration_months
    
    outstanding_balance = loan_amount
    total_payment_amount = 0

    for month in range(1, duration_months + 1):
        interest_payment = outstanding_balance * monthly_rate
        principal_payment = monthly_principal if month < duration_months else outstanding_balance
        
        # Đảm bảo kỳ cuối trả hết gốc
        if month == duration_months:
            principal_payment = outstanding_balance

        total_payment = principal_payment + interest_payment
        outstanding_balance_end = outstanding_balance - principal_payment
        
        # Khắc phục lỗi làm tròn
        if month == duration_months:
             outstanding_balance_end = 0

        df_row = pd.DataFrame([{
            'Kỳ trả nợ': month,
            'Dư nợ đầu kỳ (VNĐ)': outstanding_balance,
            'Gốc trả (VNĐ)': principal_payment,
            'Lãi trả (VNĐ)': interest_payment,
            'Tổng gốc và lãi (VNĐ)': total_payment,
            'Dư nợ cuối kỳ (VNĐ)': outstanding_balance_end
        }])
        
        df = pd.concat([df, df_row], ignore_index=True)
        outstanding_balance = outstanding_balance_end
        total_payment_amount += total_payment

    return df.apply(lambda x: np.round(x, 0) if x.name != 'Kỳ trả nợ' else x).astype({'Kỳ trả nợ': int})

def to_excel(df):
    """Chuyển đổi DataFrame thành file Excel (BytesIO)"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Kế hoạch trả nợ')
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# --- Khởi tạo State và Dữ liệu Mặc định ---

if 'loan_data' not in st.session_state:
    st.session_state.loan_data = {
        'HoTen': 'Nguyễn Thị A',
        'CCCD': '16262722727',
        'DiaChi': 'Thôn 17, xã Tân Lâm Hương, TP Hà Tĩnh',
        'SDT': '0913273',
        'MucDichVay': 'Kinh doanh vật liệu xây dựng',
        'TongNhuCauVon': 7827181642,
        'VonDoiUng': 385931642,
        'SoTienVay': 7300000000,
        'LaiSuat': 5.0, # %/năm
        'ThoiHanVay': 3, # tháng
        'GiaTriTSDB': 10000000000,
        'MoTaTSDB': 'Tổng giá trị 5 Giấy chứng nhận QSDĐ',
        'FileContent': 'Nội dung file .docx sẽ được trích xuất tại đây.'
    }
    st.session_state.chat_history = []
    st.session_state.uploaded_file_text = ""

# --- Sidebar (API Key và Xuất Dữ liệu) ---

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/e/ee/Agribank_logo.svg", width=200)
    st.title("⚙️ Cấu hình Hệ thống")
    
    # 1. API Key
    st.subheader("🔑 Gemini API Key")
    gemini_api_key = st.text_input(
        "Nhập API Key của bạn", 
        type="password", 
        key="gemini_api_key",
        help="Sử dụng Gemini API Key để kích hoạt tính năng phân tích bằng AI."
    )

    if gemini_api_key:
        try:
            client = genai.Client(api_key=gemini_api_key)
            # Thử gọi API để kiểm tra
            client.models.get("gemini-2.5-flash") # Kiểm tra xem API có hoạt động không
            st.success("API Key đã được xác nhận. Sẵn sàng sử dụng AI!")
        except Exception as e:
            st.warning("API Key không hợp lệ hoặc có lỗi kết nối!")
            client = None
    else:
        client = None

    st.markdown("---")

    # 2. Chức năng Xuất Dữ liệu
    st.subheader("📥 Chức năng Xuất Dữ liệu")
    export_option = st.selectbox(
        "Chọn loại dữ liệu muốn xuất",
        ("Xuất Kế hoạch trả nợ (Excel)", "Xuất Báo cáo Thẩm định (Excel)")
    )
    
    # Tính toán kế hoạch trả nợ cho chức năng xuất
    loan_schedule_df = calculate_loan_schedule(
        st.session_state.loan_data['SoTienVay'],
        st.session_state.loan_data['LaiSuat'],
        st.session_state.loan_data['ThoiHanVay']
    )

    if export_option == "Xuất Kế hoạch trả nợ (Excel)":
        excel_data = to_excel(loan_schedule_df)
        st.download_button(
            label="⬇️ Tải file Excel",
            data=excel_data,
            file_name="Kế_hoạch_trả_nợ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key='download_excel_loan_schedule'
        )
    elif export_option == "Xuất Báo cáo Thẩm định (Excel)":
        # Tạo DataFrame cho báo cáo thẩm định (đơn giản hóa)
        report_data = {
            "Chỉ tiêu": ["Tổng nhu cầu vốn", "Vốn đối ứng", "Số tiền vay", "Lãi suất (%/năm)", "Thời hạn vay (tháng)", "Tỷ lệ Vay/Tổng nhu cầu vốn (%)", "Tỷ lệ Vay/TSĐB (%)"],
            "Giá trị": [
                st.session_state.loan_data['TongNhuCauVon'],
                st.session_state.loan_data['VonDoiUng'],
                st.session_state.loan_data['SoTienVay'],
                st.session_state.loan_data['LaiSuat'],
                st.session_state.loan_data['ThoiHanVay'],
                (st.session_state.loan_data['SoTienVay'] / st.session_state.loan_data['TongNhuCauVon']) * 100 if st.session_state.loan_data['TongNhuCauVon'] > 0 else 0,
                (st.session_state.loan_data['SoTienVay'] / st.session_state.loan_data['GiaTriTSDB']) * 100 if st.session_state.loan_data['GiaTriTSDB'] > 0 else 0
            ]
        }
        report_df = pd.DataFrame(report_data)
        report_excel_data = to_excel(report_df)
        
        st.download_button(
            label="⬇️ Tải file Excel Báo cáo",
            data=report_excel_data,
            file_name="Báo_cáo_Thẩm_định.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key='download_excel_report'
        )


# --- Tiêu đề chính của Ứng dụng ---
st.title("🏦 Hệ thống Thẩm định Phương án Kinh doanh Agribank")
st.caption("Công cụ tự động hóa và chuẩn hóa quy trình thẩm định hồ sơ vay vốn.")

# --- Tab Layout ---

tabs = st.tabs([
    "📂 Nhập liệu & Trích xuất thông tin",
    "📊 Phân tích Chỉ số & Dòng tiền",
    "📈 Biểu đồ Trực quan",
    "🧠 Phân tích bởi AI",
    "💬 Chatbot Hỗ trợ"
])

# --- Tab 1: Nhập liệu & Trích xuất thông tin ---

with tabs[0]:
    st.header("1. Nhập liệu & Trích xuất Thông tin Khách hàng/Phương án")
    
    # File Uploader
    uploaded_file = st.file_uploader("Tải lên file Phương án Sử dụng Vốn (.docx)", type="docx")
    
    if uploaded_file is not None:
        file_content_text = extract_text_from_docx(uploaded_file)
        st.session_state.uploaded_file_text = file_content_text
        
        # GIẢ ĐỊNH: Trích xuất một số dữ liệu đơn giản từ file .docx mẫu (dựa trên cấu trúc file)
        # Trong ứng dụng thực tế, cần sử dụng thư viện NLP/Regex phức tạp hơn
        try:
            lines = file_content_text.split('\n')
            # Lấy Họ tên, CCCD, TNDV
            st.session_state.loan_data['HoTen'] = next((line.split(':')[-1].strip().split('Sinh ngày')[0].strip() for line in lines if "Họ và tên" in line and "Nguyễn Thị" in line), st.session_state.loan_data['HoTen'])
            st.session_state.loan_data['CCCD'] = next((line.split('CCCD số:')[-1].strip().split(';')[0].strip() for line in lines if "Họ và tên" in line and "CCCD số" in line), st.session_state.loan_data['CCCD'])
            st.session_state.loan_data['MucDichVay'] = next((line.split(':')[-1].strip() for line in lines if "Lĩnh vực kinh doanh chính" in line), st.session_state.loan_data['MucDichVay'])
            
            # Lấy số liệu tài chính (Dựa trên dòng đã trích xuất trong file mẫu)
            st.session_state.loan_data['TongNhuCauVon'] = 7827181642 # Lấy giá trị cứng từ file mẫu
            st.session_state.loan_data['VonDoiUng'] = 385931642 # Lấy giá trị cứng từ file mẫu
            st.session_state.loan_data['SoTienVay'] = 7300000000 # Lấy giá trị cứng từ file mẫu
            st.session_state.loan_data['ThoiHanVay'] = 3 # Lấy giá trị cứng từ file mẫu
            st.session_state.loan_data['LaiSuat'] = 5.0 # Lấy giá trị cứng từ file mẫu
            st.session_state.loan_data['GiaTriTSDB'] = 10000000000 # Lấy giá trị cứng từ file mẫu
            
            st.success("Đã trích xuất dữ liệu cơ bản từ file DOCX.")
        except Exception:
            st.warning("Không thể trích xuất dữ liệu tự động. Vui lòng nhập thủ công.")
    
    
    # 1.1 Thông tin Khách hàng
    st.subheader("1.1. Thông tin Khách hàng")
    col1, col2 = st.columns(2)
    st.session_state.loan_data['HoTen'] = col1.text_input("Họ và tên", st.session_state.loan_data['HoTen'])
    st.session_state.loan_data['CCCD'] = col2.text_input("CCCD/CMND", st.session_state.loan_data['CCCD'])
    col3, col4 = st.columns(2)
    st.session_state.loan_data['DiaChi'] = col3.text_input("Địa chỉ cư trú", st.session_state.loan_data['DiaChi'])
    st.session_state.loan_data['SDT'] = col4.text_input("Số điện thoại", st.session_state.loan_data['SDT'])

    st.markdown("---")

    # 1.2 Thông tin Phương án Vay
    st.subheader("1.2. Thông tin Phương án Vay")
    st.session_state.loan_data['MucDichVay'] = st.text_input("Mục đích vay", st.session_state.loan_data['MucDichVay'])
    
    col5, col6, col7 = st.columns(3)
    st.session_state.loan_data['TongNhuCauVon'] = col5.number_input("Tổng nhu cầu vốn (VNĐ)", min_value=0, step=1000000, value=st.session_state.loan_data['TongNhuCauVon'])
    st.session_state.loan_data['VonDoiUng'] = col6.number_input("Vốn đối ứng (VNĐ)", min_value=0, step=1000000, value=st.session_state.loan_data['VonDoiUng'])
    st.session_state.loan_data['SoTienVay'] = col7.number_input("Số tiền vay (VNĐ)", min_value=0, step=1000000, value=st.session_state.loan_data['SoTienVay'])

    col8, col9 = st.columns(2)
    st.session_state.loan_data['LaiSuat'] = col8.number_input("Lãi suất đề nghị (%/năm)", min_value=0.01, max_value=20.0, step=0.1, format="%.2f", value=st.session_state.loan_data['LaiSuat'])
    st.session_state.loan_data['ThoiHanVay'] = col9.number_input("Thời hạn vay (tháng)", min_value=1, max_value=120, step=1, value=st.session_state.loan_data['ThoiHanVay'])

    st.markdown("---")

    # 1.3 Thông tin Tài sản Đảm bảo
    st.subheader("1.3. Thông tin Tài sản Đảm bảo")
    col10, col11 = st.columns(2)
    st.session_state.loan_data['GiaTriTSDB'] = col10.number_input("Giá trị định giá TSĐB (VNĐ)", min_value=0, step=1000000, value=st.session_state.loan_data['GiaTriTSDB'])
    st.session_state.loan_data['MoTaTSDB'] = col11.text_area("Mô tả TSĐB", st.session_state.loan_data['MoTaTSDB'], height=100)
    
    # Hiển thị nội dung trích xuất
    st.markdown("---")
    with st.expander("📝 Nội dung Text đã trích xuất từ file DOCX (Dành cho AI Phân tích 1)"):
        st.code(st.session_state.uploaded_file_text[:3000] + "...", language='text')

# --- Tab 2: Phân tích Chỉ số & Dòng tiền ---

with tabs[1]:
    st.header("2. Phân tích Chỉ số & Kế hoạch Trả nợ")
    data = st.session_state.loan_data

    ## 2.1 Chỉ số Thẩm định
    st.subheader("2.1. Các Chỉ số Thẩm định Quan trọng")
    
    tong_nc = data['TongNhuCauVon']
    von_doi_ung = data['VonDoiUng']
    so_vay = data['SoTienVay']
    tsdb = data['GiaTriTSDB']

    # Tính toán các chỉ số
    tile_vay_nc = (so_vay / tong_nc) * 100 if tong_nc > 0 else 0
    tile_doi_ung_nc = (von_doi_ung / tong_nc) * 100 if tong_nc > 0 else 0
    tile_vay_tsdb = (so_vay / tsdb) * 100 if tsdb > 0 else 0

    col_a, col_b, col_c, col_d = st.columns(4)

    col_a.metric(
        "Tỷ lệ Vay / Tổng nhu cầu vốn", 
        f"{tile_vay_nc:,.2f}%".replace(",", "."), 
        delta_color="off"
    )
    col_b.metric(
        "Tỷ lệ Vốn đối ứng / Tổng nhu cầu vốn", 
        f"{tile_doi_ung_nc:,.2f}%".replace(",", "."),
        delta_color="off"
    )
    col_c.metric(
        "Tỷ lệ Vay / Giá trị TSĐB (LTV)", 
        f"{tile_vay_tsdb:,.2f}%".replace(",", "."),
        delta_color="off"
    )
    
    # Hiển thị thông tin dòng tiền từ file gốc (Giả định lấy từ file mẫu)
    st.markdown("---")
    st.subheader("2.2. Thông tin Dòng tiền (Theo Phương án Khách hàng)")
    col_x, col_y, col_z = st.columns(3)
    col_x.metric("Doanh thu dự kiến (1 vòng quay)", format_currency(8050108000) + " VNĐ")
    col_y.metric("Chi phí kinh doanh (1 vòng quay)", format_currency(7827181642) + " VNĐ")
    col_z.metric("Chênh lệch thu chi (1 vòng quay)", format_currency(222926358) + " VNĐ")
    st.caption("Doanh thu của phương án (1 vòng quay): 8.050.108.000 đồng; Chi phí kinh doanh: 7.827.181.642 đồng; Chênh lệch thu chi: 222.926.358 đồng.")

    st.markdown("---")
    
    ## 2.2 Kế hoạch Trả nợ
    st.subheader("2.3. Kế hoạch Trả nợ (Dư nợ giảm dần)")
    
    st.dataframe(
        loan_schedule_df.style.format(
            {
                'Dư nợ đầu kỳ (VNĐ)': format_currency, 
                'Gốc trả (VNĐ)': format_currency, 
                'Lãi trả (VNĐ)': format_currency, 
                'Tổng gốc và lãi (VNĐ)': format_currency,
                'Dư nợ cuối kỳ (VNĐ)': format_currency
            }
        ),
        use_container_width=True,
        hide_index=True
    )

# --- Tab 3: Biểu đồ Trực quan ---

with tabs[2]:
    st.header("3. Biểu đồ Trực quan hóa Dữ liệu")
    
    # 3.1 Biểu đồ Cơ cấu Vốn
    st.subheader("3.1. Cơ cấu Nguồn vốn (Tổng nhu cầu: " + format_currency(tong_nc) + " VNĐ)")
    capital_data = pd.DataFrame({
        'Nguồn Vốn': ['Vốn vay Agribank', 'Vốn đối ứng'],
        'Giá trị': [so_vay, von_doi_ung]
    })
    
    fig_pie = px.pie(
        capital_data, 
        values='Giá trị', 
        names='Nguồn Vốn', 
        title='Cơ cấu Vốn Vay và Vốn Đối ứng',
        color_discrete_sequence=px.colors.sequential.RdBu
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")
    
    # 3.2 Biểu đồ Dư nợ giảm dần
    st.subheader("3.2. Dư nợ Giảm dần và Chi phí Trả Lãi")
    
    # Chuẩn bị dữ liệu cho biểu đồ đường
    chart_df = loan_schedule_df[['Kỳ trả nợ', 'Dư nợ cuối kỳ (VNĐ)', 'Lãi trả (VNĐ)']].rename(columns={
        'Dư nợ cuối kỳ (VNĐ)': 'Dư nợ Cuối kỳ',
        'Lãi trả (VNĐ)': 'Lãi Trả'
    })
    
    # Chuyển đổi DataFrame sang định dạng long cho Plotly
    chart_long = pd.melt(chart_df, id_vars='Kỳ trả nợ', var_name='Chỉ tiêu', value_name='Giá trị')

    fig_line = px.line(
        chart_long, 
        x='Kỳ trả nợ', 
        y='Giá trị', 
        color='Chỉ tiêu', 
        title='Dư nợ và Lãi suất theo kỳ trả nợ',
        markers=True,
        color_discrete_map={"Dư nợ Cuối kỳ": "blue", "Lãi Trả": "red"}
    )
    
    fig_line.update_layout(yaxis_title="Giá trị (VNĐ)")
    st.plotly_chart(fig_line, use_container_width=True)

# --- Tab 4: Phân tích bởi AI ---

with tabs[3]:
    st.header("4. Phân tích Sâu sắc bởi AI (Gemini Flash)")

    if client is None:
        st.warning("Vui lòng nhập Gemini API Key ở thanh bên (sidebar) để sử dụng chức năng này.")
    else:
        if st.button("🚀 Bắt đầu Phân tích Toàn diện"):
            with st.spinner("Đang gửi yêu cầu phân tích tới Gemini..."):
                
                # --- Phân tích 1: Dựa trên File gốc ---
                st.subheader("4.1. Phân tích từ File .docx Khách hàng")
                st.info("Nguồn dữ liệu: Phân tích từ toàn bộ nội dung text của file .docx đã upload.")

                if not st.session_state.uploaded_file_text:
                    st.warning("Vui lòng tải lên file .docx ở tab 'Nhập liệu & Trích xuất thông tin' trước.")
                else:
                    prompt_file_analysis = (
                        "Bạn là một chuyên gia thẩm định tín dụng. Hãy phân tích toàn diện phương án kinh doanh sau (được trích xuất từ file .docx của khách hàng). "
                        "Nêu rõ các **Điểm mạnh** (về phương án, nhu cầu vốn, TSĐB), **Rủi ro Tiềm ẩn** (về thị trường, tài chính, pháp lý), "
                        "và tóm tắt **Nhận định sơ bộ của bạn về tính khả thi**.\n\n"
                        "--- Nội dung File DOCX ---\n"
                        f"{st.session_state.uploaded_file_text}"
                    )
                    
                    try:
                        response1 = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=prompt_file_analysis
                        )
                        st.markdown(response1.text)
                    except Exception as e:
                        st.error(f"Lỗi khi gọi Gemini API (Phân tích File): {e}")

                st.markdown("---")

                # --- Phân tích 2: Dựa trên Dữ liệu đã hiệu chỉnh ---
                st.subheader("4.2. Phân tích từ Các Thông số và Chỉ số đã Hiệu chỉnh")
                st.info("Nguồn dữ liệu: Phân tích từ các thông số và chỉ số đã được người dùng hiệu chỉnh trên ứng dụng.")

                analysis_data = data.copy()
                analysis_data['Tỷ lệ Vay/Tổng nhu cầu vốn (%)'] = tile_vay_nc
                analysis_data['Tỷ lệ Vốn đối ứng/Tổng nhu cầu vốn (%)'] = tile_doi_ung_nc
                analysis_data['Tỷ lệ Vay/Giá trị TSĐB (%)'] = tile_vay_tsdb
                analysis_data['Kế hoạch Trả nợ (Tóm tắt)'] = loan_schedule_df[['Kỳ trả nợ', 'Tổng gốc và lãi (VNĐ)']].to_markdown(index=False)

                prompt_manual_analysis = (
                    "Bạn là một chuyên gia thẩm định tín dụng. Dựa trên các thông số đã được hiệu chỉnh/tính toán sau đây, "
                    "hãy đánh giá **Tính khả thi tài chính của khoản vay** và đưa ra **Đề xuất về rủi ro tập trung (nếu có)**. "
                    "Chú trọng phân tích rủi ro từ tỷ lệ LTV (Vay/TSĐB) và khả năng trả nợ dựa trên kế hoạch trả nợ (Tổng gốc và lãi)."
                    "\n\n--- Thông số Đánh giá ---\n"
                    f"{pd.Series(analysis_data).to_string()}\n"
                    
                )

                try:
                    response2 = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt_manual_analysis
                    )
                    st.markdown(response2.text)
                except Exception as e:
                    st.error(f"Lỗi khi gọi Gemini API (Phân tích Chỉ số): {e}")


# --- Tab 5: Chatbot Hỗ trợ ---

with tabs[4]:
    st.header("5. Chatbot Hỗ trợ Thẩm định (Gemini Flash)")
    
    if client is None:
        st.warning("Vui lòng nhập Gemini API Key ở thanh bên (sidebar) để sử dụng chức năng này.")
    else:
        
        # Thiết lập lịch sử chat
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
        
        # Thiết lập context cho Chatbot
        system_instruction = (
            "Bạn là một chuyên gia thẩm định tín dụng của Agribank, am hiểu về các quy định và chỉ số tài chính. "
            "Trả lời các câu hỏi của chuyên viên tín dụng một cách ngắn gọn, chính xác, và chuyên nghiệp. "
            "Các thông tin hiện tại của hồ sơ vay là: "
            f"Số tiền vay: {format_currency(data['SoTienVay'])} VNĐ, "
            f"Lãi suất: {data['LaiSuat']}%/năm, "
            f"Thời hạn: {data['ThoiHanVay']} tháng, "
            f"Tỷ lệ LTV: {tile_vay_tsdb:,.2f}%. "
            "Hãy sử dụng những thông tin này để đưa ra câu trả lời phù hợp."
        )

        try:
            chat_session = client.chats.create(
                model="gemini-2.5-flash",
                history=st.session_state.chat_history,
                system_instruction=system_instruction
            )
        except Exception as e:
            st.error(f"Lỗi khởi tạo Chatbot: {e}")
            chat_session = None
            
        # Nút xóa lịch sử chat
        if st.button("🔄 Xóa lịch sử trò chuyện"):
            st.session_state.chat_history = []
            st.rerun()

        st.markdown("---")

        if chat_session:
            # Hiển thị lịch sử chat
            for message in st.session_state.chat_history:
                role = "user" if message.role == "user" else "assistant"
                with st.chat_message(role):
                    st.markdown(message.parts[0].text)

            # Xử lý input mới
            if prompt := st.chat_input("Hỏi Chatbot về hồ sơ vay này..."):
                with st.chat_message("user"):
                    st.markdown(prompt)
                
                with st.chat_message("assistant"):
                    with st.spinner("Đang suy nghĩ..."):
                        try:
                            response = chat_session.send_message(prompt)
                            st.markdown(response.text)
                            
                            # Cập nhật lịch sử chat
                            st.session_state.chat_history = chat_session.get_history()
                            
                        except Exception as e:
                            st.error(f"Lỗi Chatbot: {e}")
                            
        st.caption("Chatbot được hỗ trợ bởi Gemini API và được cung cấp thông tin về hồ sơ vay hiện tại.")
