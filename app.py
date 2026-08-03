import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import random
import string

st.set_page_config(page_title="Generator Dashboard", page_icon="📄", layout="wide")

st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Tool:", ["Contract Generator", "Medical Report Generator", "Insurance Generator", "Salary Generator", "EPF Generator"])

if page == "Contract Generator":
    st.title("📄 Worker Contract Generator")
    st.write("Upload your Excel file to generate employment contracts.")

    DEFAULT_TEMPLATE = """EMPLOYMENT CONTRACT

This agreement is made and entered on 1 January 2025 between ARIANA GLOBAL Malaysia,
(herein called the company) through our lawful attorney present in Bangladesh First Tours and
Travels. Recruiting License No. 981, addressed at Plot # 95, Road # 07, Sector #04, Uttara,
Dhaka-1230, Bangladesh.

    Name: {name}
    Nationality: {nationality}
    Passport No: {passport}
    Place of issue: {place_of_issue}
    Date of issue: {date_of_issue}

In his capacity as the Second Party hereby agreed the following terms and conditions of
Employment Contract:

1- The SECOND PARTY agreed to work with the first party as: General Workers with the
basic salary of RM1,700.00 (Ringgit Malaysia One Thousand Seven Hundred) per month.

2- Period of Employment: 2 (Two years) (With renewable option)

3- Place of employment: Malaysia

4- Air Ticket: For joining the company for the first time (DAC-KUL) will be paid by the
worker and returning after completion of contract will be paid by the employer

5- Visa charge and Levi is borne by Company itself and will not be deducted in workers' salary

6- Working Hours: 8 hours per day, 6 days per week (48 hours per week)

7- Over time: As per Labour Law of Malaysia

8- Probation Period: 90 days

9- Work permit: Work permit from the immigration will be provided by the company free of cost.

10- Accommodation: Free Bachelor accommodation should be provided by the company.

11- Water, Electricity & Gas: Should be provided by the company

12- Medical and Work Insurance: Provided by the company.

13- Local Transportation: Provided by the company during working hour

14- Uniform, and Safety Materials etc: Provided by the Company

15- Annual paid Leave: As per Labour law of Malaysia

16- In case of death of the employee during the contract period, the First Party shall agree
to repatriate the remains of the deceased at the expense of the company. Both in the
case of death and injury, compensation shall be paid according to the Labor Law of the host country.

17- Other Terms & Conditions: As per Malaysian Labor Law

Authorized Signature
Name: Melissa
Designation: Manager
Signature: Melissa
"""

    with st.expander("View / Edit Contract Template"):
        template_text = st.text_area("Template Text", value=DEFAULT_TEMPLATE, height=400)

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="contract_upload")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).replace('*', '').strip() for col in df.columns]
            
            st.success("File uploaded successfully! Here is a preview of your data:")
            st.dataframe(df.head())
            
            required_columns = ['First Name', 'Last Name', 'Nationality', 'Passport No', 'Passport Issue Date']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            else:
                if st.button("Generate PDF Contracts"):
                    with st.spinner("Generating contracts..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for index, row in df.iterrows():
                                pdf = FPDF()
                                pdf.add_page()
                                pdf.set_font("Arial", size=11)
                                
                                full_name = f"{row['First Name']} {row['Last Name']}"
                                contract_content = template_text.format(
                                    name=full_name,
                                    nationality=row['Nationality'],
                                    passport=row['Passport No'],
                                    place_of_issue=row['Nationality'],
                                    date_of_issue=row['Passport Issue Date']
                                )
                                
                                lines = contract_content.split('\n', 1)
                                if lines:
                                    title = lines[0].strip()
                                    body = lines[1] if len(lines) > 1 else ""
                                    pdf.set_font("Arial", 'B', 12)
                                    pdf.cell(0, 10, title, ln=True, align='C')
                                    pdf.set_font("Arial", size=11)
                                    pdf.multi_cell(0, 6, body)
                                else:
                                    pdf.multi_cell(0, 6, contract_content)
                                
                                pdf_bytes = pdf.output(dest='S').encode('latin1')
                                filename = f"Employment_Contract_{full_name}.pdf"
                                zf.writestr(filename, pdf_bytes)
                        
                        st.success("All contracts generated successfully!")
                        st.download_button("⬇️ Download All Contracts (ZIP)", data=zip_buffer.getvalue(), file_name="contracts.zip", mime="application/zip")
        except Exception as e:
            st.error(f"An error occurred: {e}")

elif page == "Medical Report Generator":
    st.title("🏥 Medical Report Generator")
    st.write("Upload your Excel file to generate FOMEMA Medical Reports.")
    
    st.info("Note: The Medical Report layout is strictly formatted according to FOMEMA standards.")

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="medical_upload")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).replace('*', '').strip() for col in df.columns]
            
            st.success("File uploaded successfully! Here is a preview of your data:")
            st.dataframe(df.head())
            
            required_columns = ['First Name', 'Last Name', 'Nationality', 'Passport No', 'Date of Birth', 'Gender', 'Sector']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            else:
                if st.button("Generate Medical Reports"):
                    with st.spinner("Generating medical reports..."):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for index, row in df.iterrows():
                                pdf = FPDF()
                                pdf.add_page()
                                pdf.set_auto_page_break(auto=False)
                                
                                # 1. Header FOMEMA
                                pdf.set_font("Arial", 'BI', 24)
                                pdf.cell(0, 15, "FOMEMA", ln=True, align='L')
                                pdf.ln(5)
                                
                                # 2. Boxed Title
                                pdf.set_font("Arial", 'B', 12)
                                pdf.cell(0, 10, "FOMEMA REPORT", border=1, ln=True, align='C')
                                pdf.ln(5)
                                
                                # 3. DISCLAIMER
                                pdf.set_font("Arial", 'B', 10)
                                pdf.cell(0, 6, "DISCLAIMER", ln=True, align='C')
                                
                                pdf.set_font("Arial", 'B', 7)
                                disclaimer_text = (
                                    "This FOMEMA report (\"this Report\") has been generated automatically by computer software and may contain inaccuracies or errors. It is not intended to replace the expertise and judgment of healthcare professionals. This Report is provided to you on an \"as is\" and \"as available\" basis at the sole discretion of FOMEMA Sdn. Bhd. (\"the Company\") and the Company has not verified this Report for accuracy and does not warrant the accuracy of, or make any other warranties or representations regarding this Report. The information contained in this record is provided for informational and reference purposes only and should not be used for diagnosing or treating any medical condition. Any reliance on the information contained herein is solely at your own risk. You are solely responsible for any interpretation, use and/or reliance made based on this Report and FOMEMA accepts no liability for any loss or damage arising from the interpretation, use, or reliance on the report. This includes any direct, indirect, incidental, consequential, or punitive damages that may result from the use of this record or its contents. It is recommended that you consult a qualified healthcare professional for any medical advice, diagnosis or treatment. Kindly refer to the examining doctor if further clarification is required.\n\n"
                                    "The receipt of this Report shall be deemed as the Recipient's true acknowledgement and acceptance of this disclaimer."
                                )
                                pdf.multi_cell(0, 3.5, disclaimer_text)
                                pdf.ln(8)
                                
                                # 4. PART I
                                pdf.set_font("Arial", 'BU', 10)
                                pdf.cell(0, 6, "PART I. FOREIGN WORKER INFORMATION", ln=True, align='C')
                                pdf.ln(5)
                                
                                # Columns setup
                                pdf.set_font("Arial", 'B', 9)
                                col1_x = 10
                                col1_val_x = 50
                                col2_x = 100
                                col2_val_x = 145
                                
                                full_name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                                worker_code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))
                                
                                # Row 1
                                y = pdf.get_y()
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Worker Name")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, f": {full_name}")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Worker Code")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, f": {worker_code}")
                                
                                # Row 2
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Country of Origin")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, f": {row.get('Nationality', '')}")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Date of Birth")
                                dob = str(row.get('Date of Birth', '')).split()[0]
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, f": {dob}")
                                
                                # Row 3
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Passport Number")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, f": {row.get('Passport No', '')}")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Gender")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, f": {row.get('Gender', '')}")
                                
                                # Row 4
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Job Type")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, f": {row.get('Sector', '')}")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Employer Code")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, ": E6ED012445")
                                
                                # Date Calculations
                                cert_date_str = "16/06/2026"
                                exam_date_str = "15/06/2026"
                                    
                                # Row 5
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Doctor Code")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": D4ES000306")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Physical Examination Date")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, f": {exam_date_str}")
                                
                                # Row 6
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Transaction ID")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": 20240729860413")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Certification Date")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, f": {cert_date_str}")
                                
                                # Row 7
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Passport Expiry Date")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": -")
                                
                                pdf.set_y(y + 10)
                                
                                # 5. PART II
                                pdf.set_font("Arial", 'BU', 10)
                                pdf.cell(0, 6, "PART II. MEDICAL HISTORY", ln=True, align='C')
                                pdf.ln(3)
                                
                                pdf.set_font("Arial", 'B', 8)
                                pdf.cell(100, 6, "1.       CATEGORY 1 DISEASES")
                                pdf.cell(30, 6, "YES", align='C')
                                pdf.cell(30, 6, "NO", align='C')
                                pdf.cell(30, 6, "DATE (DD/MM/YYYY)", align='C')
                                pdf.ln(8)
                                
                                diseases = [
                                    ("1.1", "TUBERCULOSIS"),
                                    ("1.2", "VIRAL HEPATITIS B"),
                                    ("1.3", "VIRAL HEPATITIS C"),
                                    ("1.4", "SYPHILIS"),
                                    ("1.5", "HIV"),
                                    ("1.6", "MALARIA"),
                                    ("1.7", "FILARIASIS")
                                ]
                                
                                pdf.set_font("Arial", '', 8)
                                for num, disease in diseases:
                                    y = pdf.get_y()
                                    pdf.set_xy(10, y)
                                    pdf.cell(10, 6, num)
                                    pdf.set_xy(20, y)
                                    pdf.cell(80, 6, disease)
                                    
                                    # YES Checkbox (empty)
                                    pdf.rect(123, y + 1.5, 3, 3)
                                    
                                    # NO Checkbox (checked with an X)
                                    pdf.rect(153, y + 1.5, 3, 3)
                                    pdf.set_xy(153, y)
                                    pdf.set_font("Arial", 'B', 7)
                                    pdf.cell(3, 6, "X", align='C')
                                    pdf.set_font("Arial", '', 8)
                                    
                                    pdf.ln(5)
                                
                                pdf.ln(2)
                                pdf.set_font("Arial", 'B', 8)
                                pdf.cell(0, 5, "Note:", ln=True)
                                pdf.set_font("Arial", 'B', 7)
                                pdf.cell(0, 4, "- 1. Foreign worker with a medical history of the Category 1 Diseases is deemed to be unsuitable for employment in Malaysia.", ln=True)
                                pdf.cell(0, 4, "- 2. However, foreign worker who gives a medical history of Hepatitis B, Hepatitis C, Syphilis, HIV, Malaria or Filariasis but does not", ln=True)
                                pdf.cell(0, 4, "show any clinical evidence of the above and the blood test results are negative, the foreign worker is deemed to be suitable for employment in Malaysia.", ln=True)
                                
                                # 6. PART V
                                pdf.ln(5)
                                pdf.set_font("Arial", 'BU', 10)
                                pdf.cell(0, 6, "PART V: CERTIFICATION BY DOCTOR (cont'd)", ln=True, align='C')
                                pdf.ln(3)
                                
                                y = pdf.get_y()
                                pdf.set_font("Arial", 'B', 7)
                                pdf.set_xy(140, y)
                                pdf.cell(20, 6, "UNSUITABLE", align='C')
                                pdf.set_xy(170, y)
                                pdf.cell(20, 6, "SUITABLE", align='C')
                                pdf.ln(8)
                                
                                y = pdf.get_y()
                                pdf.set_font("Arial", 'B', 8)
                                pdf.set_xy(10, y)
                                pdf.cell(10, 4, "23.")
                                pdf.set_xy(20, y)
                                pdf.multi_cell(110, 4, "AFTER REVIEWING THE MEDICAL EXAMINATION REPORT,\nI HEREBY CERTIFY THIS FOREIGN WORKER TO BE\nMEDICALLY FOR EMPLOYMENT IN MALAYSIA")
                                
                                # UNSUITABLE Checkbox (empty)
                                pdf.rect(148.5, y + 4, 3, 3)
                                
                                # SUITABLE Checkbox (checked with a tick)
                                pdf.rect(178.5, y + 4, 3, 3)
                                pdf.set_xy(178.5, y + 2.5)
                                pdf.set_font("Arial", 'B', 7)
                                pdf.cell(3, 6, "X", align='C')
                                
                                pdf.set_y(y + 14)
                                y = pdf.get_y()
                                pdf.set_font("Arial", 'B', 8)
                                pdf.set_xy(10, y)
                                pdf.cell(10, 4, "24.")
                                pdf.set_xy(20, y)
                                pdf.multi_cell(110, 4, "Comments (refer to Part V - Item 16)\n-")
                                
                                # footer (Page: 1 of 1)
                                pdf.set_y(282)
                                pdf.set_font("Arial", 'I', 10)
                                pdf.cell(0, 10, f"Page: {pdf.page_no()}", align='C')
                                
                                pdf_bytes = pdf.output(dest='S').encode('latin1')
                                filename = f"Medical_Report_{full_name}.pdf"
                                zf.writestr(filename, pdf_bytes)
                        
                        st.success("All medical reports generated successfully!")
                        st.download_button("⬇️ Download All Medical Reports (ZIP)", data=zip_buffer.getvalue(), file_name="medical_reports.zip", mime="application/zip")
        except Exception as e:
            st.error(f"An error occurred: {e}")

elif page == "Insurance Generator":
    st.title("🛡️ Insurance Slip Generator (SPIKPA)")
    st.write("Upload your Excel file to generate SPIKPA insurance slips.")

    # UI Inputs
    nama_majikan = st.text_input("Nama Majikan")
    alamat_majikan = st.text_area("Alamat")
    tarikh_mula_input = st.date_input("Tarikh Perlindungan (Mula)")
    
    coverage_period = st.radio("Tempoh Perlindungan", ["1 Tahun", "6 Bulan"])
    mula_dt = pd.to_datetime(tarikh_mula_input)
    if coverage_period == "1 Tahun":
        tarikh_tamat_dt = mula_dt + pd.DateOffset(years=1)
    else:
        tarikh_tamat_dt = mula_dt + pd.DateOffset(months=6)
        
    st.text_input("Tarikh Perlindungan (Tamat)", value=tarikh_tamat_dt.strftime("%Y-%m-%d"), disabled=True)
    
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="insurance_upload")

    if uploaded_file is not None and nama_majikan and alamat_majikan:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).replace('*', '').strip() for col in df.columns]
            
            st.success("File uploaded successfully! Preview:")
            st.dataframe(df.head())
            
            required_columns = ['First Name', 'Last Name', 'Nationality', 'Passport No']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            else:
                if st.button("Generate Insurance Slips"):
                    with st.spinner("Generating insurance slips..."):
                        zip_buffer = io.BytesIO()
                        import datetime
                        import os
                        
                        # Process dates
                        e_pass_date = pd.to_datetime(tarikh_mula_input)
                        
                        # Chunking by 10
                        chunk_size = 10
                        chunks = [df[i:i + chunk_size] for i in range(0, df.shape[0], chunk_size)]
                        
                        with zipfile.ZipFile(zip_buffer, "w") as zf:
                            for chunk_idx, chunk in enumerate(chunks):
                                pdf = FPDF(orientation='P', unit='mm', format='A4')
                                pdf.add_page()
                                pdf.set_auto_page_break(auto=False)
                                
                                # Random Reference No
                                ref_no = f"KKM{random.randint(1000000, 9999999)}"
                                
                                # First row info for top dates
                                date_dicetak = e_pass_date - pd.DateOffset(months=1)
                                dicetak_str = date_dicetak.strftime("%d-%m-%Y")
                                
                                # 1. Header Right
                                pdf.set_font("Arial", '', 9)
                                pdf.cell(0, 5, "KKM/SPIKPA/SP/1/2011", ln=True, align='R')
                                pdf.ln(2)
                                
                                # 2. Box Title
                                pdf.set_font("Arial", 'B', 10)
                                pdf.cell(0, 8, "SLIP PENGESAHAN SKIM PERLINDUNGAN INSURANS KESIHATAN PEKERJA ASING(SPIKPA)", border=1, ln=True, align='C')
                                pdf.ln(3)
                                
                                # 3. Logo and Company Info
                                y_logo = pdf.get_y()
                                
                                pdf.set_xy(10, y_logo + 8)
                                pdf.set_font("Arial", 'B', 14)
                                pdf.cell(45, 5, "PROTECTHEALTH", align='L')
                                    
                                pdf.set_xy(60, y_logo)
                                pdf.set_font("Arial", 'B', 9)
                                pdf.cell(0, 5, "ProtectHealth Corporation Sdn Bhd (1212734-T)", ln=True)
                                pdf.set_font("Arial", '', 8)
                                pdf.set_x(60)
                                pdf.cell(0, 4, "F01 & F02, Tingkat 1, Blok 2300, Century Square, Jalan Usahawan, 63000 Cyberjaya, Selangor, Malaysia.", ln=True)
                                pdf.set_x(60)
                                pdf.cell(0, 4, "No. Telefon : (603) 8687 4848 / (603) 8687 4888              Portal Web : www.spikpa.com.my", ln=True)
                                pdf.set_font("Arial", 'I', 8)
                                pdf.set_x(60)
                                pdf.cell(0, 4, "[ Dilantik oleh Kementerian Kesihatan Malaysia sebagai Pembekal Perkhidmatan Elektronik (ESP) untuk", ln=True)
                                pdf.set_x(60)
                                pdf.cell(0, 4, "Skim Perlindungan Insurans Kesihatan Pekerja Asing (SPIKPA) ]", ln=True)
                                
                                pdf.ln(10)
                                
                                # 4. Info Section
                                pdf.set_font("Arial", '', 10)
                                
                                # Row 1: No Ruj & Tarikh Dicetak
                                y_info = pdf.get_y()
                                pdf.set_xy(10, y_info)
                                pdf.cell(35, 6, "No. Ruj. Slip SPIKPA")
                                pdf.cell(5, 6, ":")
                                pdf.cell(45, 6, ref_no)
                                
                                pdf.set_xy(105, y_info)
                                pdf.cell(45, 6, "Tarikh Slip SPIKPA Dicetak")
                                pdf.cell(5, 6, ":")
                                pdf.cell(45, 6, dicetak_str)
                                
                                # Row 2: Nama Majikan
                                pdf.set_xy(10, y_info + 8)
                                pdf.cell(35, 6, "Nama Majikan")
                                pdf.cell(5, 6, ":")
                                pdf.cell(140, 6, nama_majikan)
                                
                                # Row 3: Alamat
                                pdf.set_xy(10, y_info + 16)
                                pdf.cell(35, 6, "Alamat")
                                pdf.cell(5, 6, ":")
                                
                                # Multi-cell for address
                                pdf.set_xy(50, y_info + 16)
                                pdf.multi_cell(140, 5, alamat_majikan)
                                
                                # Adjust Y after multi_cell
                                pdf.set_y(pdf.get_y() + 5)
                                
                                # 5. SKHPPA Table
                                pdf.set_font("Arial", 'B', 9)
                                pdf.cell(0, 8, "MAKLUMAT INSURANS SKIM KEMASUKAN HOSPITAL DAN PEMBEDAHAN PEKERJA ASING (SKHPPA)", border=1, ln=True, align='C')
                                
                                pdf.set_font("Arial", 'B', 8)
                                pdf.cell(30, 8, "Nama Syarikat", border='L', align='C')
                                pdf.cell(0, 8, "CHUBB INSURANCE MALAYSIA BERHAD (formerly known as ACE JERNEH INSURANCE BHD)", border='R', ln=True, align='L')
                                pdf.cell(30, 4, "Ins.", border='LB', align='C')
                                pdf.cell(0, 4, "", border='RB', ln=True)
                                
                                # Subheader
                                pdf.cell(30, 8, "Kod Agen Ins.", border=1, align='C')
                                pdf.cell(40, 8, "No. Polisi Ins.", border=1, align='C')
                                pdf.cell(60, 8, "No.Pindaan / Endorsement", border=1, align='C')
                                pdf.cell(0, 8, "Tarikh Perlindungan Polisi (Tamat)", border=1, ln=True, align='C')
                                
                                # Table values
                                tamat_date = tarikh_tamat_dt
                                malay_months = {
                                    'JANUARY': 'JANUARI', 'FEBRUARY': 'FEBRUARI', 'MARCH': 'MAC', 'APRIL': 'APRIL',
                                    'MAY': 'MEI', 'JUNE': 'JUN', 'JULY': 'JULAI', 'AUGUST': 'OGOS',
                                    'SEPTEMBER': 'SEPTEMBER', 'OCTOBER': 'OKTOBER', 'NOVEMBER': 'NOVEMBER', 'DECEMBER': 'DISEMBER'
                                }
                                month_eng = tamat_date.strftime("%B").upper()
                                month_malay = malay_months.get(month_eng, month_eng)
                                tamat_str_malay = f"{tamat_date.strftime('%d')} {month_malay} {tamat_date.strftime('%Y')}"
                                
                                pdf.set_font("Arial", '', 8)
                                pdf.cell(30, 10, "N0189AHQ", border=1, align='C')
                                pdf.cell(40, 10, "HQ-S1052525-SPA", border=1, align='C')
                                pdf.cell(60, 10, "", border=1, align='C')
                                pdf.cell(0, 10, tamat_str_malay, border=1, ln=True, align='C')
                                
                                pdf.ln(5)
                                
                                # 6. Worker Table
                                pdf.set_font("Arial", 'B', 9)
                                pdf.cell(0, 5, "SENARAI PEKERJA ASING", ln=True)
                                pdf.cell(0, 5, "YANG DIINSURANSKAN", ln=True)
                                
                                pdf.cell(15, 8, "Bil.", border=1, align='C')
                                pdf.cell(60, 8, "Nama", border=1, align='C')
                                pdf.cell(40, 8, "Warganegara", border=1, align='C')
                                pdf.cell(35, 8, "No. Pasport", border=1, align='C')
                                pdf.cell(0, 8, "Tarikh Perlindungan(Mula)", border=1, ln=True, align='C')
                                
                                pdf.set_font("Arial", '', 8)
                                for i, (_, row) in enumerate(chunk.iterrows(), start=1):
                                    full_name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                                    warganegara = str(row.get('Nationality', '')).upper()
                                    pasport = str(row.get('Passport No', ''))
                                    
                                    m_month = e_pass_date.strftime("%B").upper()
                                    m_month_my = malay_months.get(m_month, m_month)
                                    mula_str = f"{e_pass_date.strftime('%d')} {m_month_my} {e_pass_date.strftime('%Y')}"
                                    
                                    pdf.cell(15, 8, str(i), border=1, align='C')
                                    pdf.cell(60, 8, full_name[:35], border=1, align='L')
                                    pdf.cell(40, 8, warganegara, border=1, align='L')
                                    pdf.cell(35, 8, pasport, border=1, align='L')
                                    pdf.cell(0, 8, mula_str, border=1, ln=True, align='C')
                                    
                                pdf.ln(10)
                                
                                # 7. Footer Boxes
                                y_footer = pdf.get_y()
                                
                                # NOTIS PENTING Box
                                pdf.set_xy(10, y_footer)
                                pdf.set_font("Arial", 'B', 9)
                                pdf.cell(90, 6, "NOTIS PENTING", border=1, align='L')
                                pdf.set_xy(10, y_footer + 6)
                                pdf.set_font("Arial", '', 8)
                                pdf.rect(10, y_footer + 6, 90, 20)
                                pdf.set_xy(12, y_footer + 8)
                                pdf.multi_cell(86, 4, "Slip ini wajib disertakan bersama dengan polisi insurans SKHPPA asal (original) dan dokumen-dokumen keperluan lain bagi permohonan atau pembaharuan PLKS.")
                                
                                # AKUAN PENERIMAAN Box
                                pdf.set_xy(110, y_footer)
                                pdf.set_font("Arial", 'B', 9)
                                pdf.cell(90, 6, "AKUAN PENERIMAAN", border=1, align='C')
                                pdf.set_xy(110, y_footer + 6)
                                pdf.rect(110, y_footer + 6, 90, 20)
                                pdf.set_xy(110, y_footer + 12)
                                pdf.cell(90, 6, "OLEH JABATAN IMIGRESEN MALAYSIA(JIM)", align='C')
                                
                                pdf_bytes = pdf.output(dest='S').encode('latin1')
                                
                                chunk_start = chunk_idx * chunk_size + 1
                                chunk_end = chunk_start + len(chunk) - 1
                                filename = f"Insurance_Slip_{chunk_start}_to_{chunk_end}.pdf"
                                zf.writestr(filename, pdf_bytes)
                        
                        st.success("All insurance slips generated successfully!")
                        st.download_button("⬇️ Download All Insurance Slips (ZIP)", data=zip_buffer.getvalue(), file_name="insurance_slips.zip", mime="application/zip")
                        
        except Exception as e:
            st.error(f"An error occurred: {e}")

elif page == "Salary Generator":
    st.title("💰 Salary Generator")
    st.write("Upload your Excel file to generate the monthly salary sheet.")

    # UI Inputs
    payment_amount = st.number_input("Payment Amount", min_value=0.00, value=1700.00, format="%.2f")
    
    col1, col2 = st.columns(2)
    with col1:
        months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
        month_sel = st.selectbox("Month", months, index=6) # Default JULY
    with col2:
        years = [str(y) for y in range(2025, 2031)]
        year_sel = st.selectbox("Year", years, index=1) # Default 2026
        
    payment_desc = f"{month_sel} {year_sel}"
    
    crediting_date = st.date_input("Crediting Date")
    crediting_date_str = crediting_date.strftime("%d/%m/%Y")
    
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="salary_upload")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).replace('*', '').strip() for col in df.columns]
            
            st.success("File uploaded successfully! Preview:")
            st.dataframe(df.head())
            
            required_columns = ['First Name', 'Last Name', 'Passport No']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            else:
                if st.button("Generate Salary Excel"):
                    with st.spinner("Generating Excel..."):
                        excel_buffer = io.BytesIO()
                        
                        # Prepare data
                        table_data = []
                        for _, row in df.iterrows():
                            full_name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                            account_no = ''.join(random.choices(string.digits, k=12))
                            
                            table_data.append({
                                "Beneficiary Name": full_name,
                                "Beneficiary Bank": "MAYBANK",
                                "Beneficiary Account No": account_no,
                                "ID Type": "PASSPORT",
                                "Passport Number": str(row.get('Passport No', '')),
                                "Payment Amount": payment_amount,
                                "Payment Reference": "SALARY",
                                "Payment Description": payment_desc
                            })
                            
                        df_table = pd.DataFrame(table_data)
                        
                        # Header information
                        header_data = [
                            ["Employer Info :", ""],
                            ["Crediting Date (eg. dd/MM/yyyy)", crediting_date_str],
                            ["Payment Reference", "SALARY"],
                            ["Payment Description", payment_desc],
                            ["Bulk Payment Type", "Salary"]
                        ]
                        df_header = pd.DataFrame(header_data)
                        
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            # Write header info without column names
                            df_header.to_excel(writer, index=False, header=False, startrow=0, sheet_name="Salary")
                            
                            # Write table data starting from row 6 (0-indexed -> 6 means 7th row)
                            df_table.to_excel(writer, index=False, startrow=6, sheet_name="Salary")
                            
                            # Auto-adjust column widths
                            worksheet = writer.sheets["Salary"]
                            for col in worksheet.columns:
                                max_length = 0
                                column = col[0].column_letter # Get the column name
                                for cell in col:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                adjusted_width = (max_length + 2)
                                worksheet.column_dimensions[column].width = adjusted_width
                        
                        st.success("Salary Excel generated successfully!")
                        st.download_button(
                            label="⬇️ Download Salary Excel",
                            data=excel_buffer.getvalue(),
                            file_name=f"Salary_{month_sel}_{year_sel}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        except Exception as e:
            st.error(f"An error occurred: {e}")

elif page == "EPF Generator":
    st.title("🏦 EPF Generator")
    st.write("Upload your Excel file to generate the EPF contribution statement (Penyata Caruman).")

    # UI Inputs
    nama_majikan = st.text_input("Nama Majikan")
    alamat_majikan = st.text_area("Alamat Majikan")
    tarikh = st.date_input("Tarikh")
    tarikh_str = tarikh.strftime("%d/%m/%Y")
    salary = st.number_input("Salary", min_value=0.00, value=1700.00, format="%.2f")
    
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"], key="epf_upload")

    if uploaded_file is not None and nama_majikan and alamat_majikan:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = [str(col).replace('*', '').strip() for col in df.columns]
            
            st.success("File uploaded successfully! Preview:")
            st.dataframe(df.head())
            
            required_columns = ['First Name', 'Last Name', 'Passport No']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            else:
                if st.button("Generate EPF PDF"):
                    with st.spinner("Generating EPF statement..."):
                        pdf = FPDF(orientation='P', unit='mm', format='A4')
                        pdf.add_page()
                        pdf.set_auto_page_break(auto=True, margin=15)
                        
                        # Header
                        pdf.set_font("Arial", 'B', 14)
                        pdf.cell(0, 6, "KUMPULAN WANG SIMPANAN PEKERJA", ln=True, align='C')
                        pdf.cell(0, 6, "PENYATA CARUMAN", ln=True, align='C')
                        pdf.set_font("Arial", '', 10)
                        pdf.cell(0, 6, "(Untuk Simpanan Majikan)", ln=True, align='C')
                        month_year = tarikh.strftime("%m/%Y")
                        pdf.cell(0, 6, f"Caruman bulan {month_year} yang telah diproses", ln=True, align='C')
                        pdf.ln(10)
                        
                        # Box for Majikan Info
                        y_before = pdf.get_y()
                        pdf.rect(10, y_before, 190, 45) # Draw a big box
                        
                        pdf.set_xy(12, y_before + 5)
                        pdf.set_font("Arial", '', 10)
                        
                        pdf.cell(40, 5, "No. Rujukan Majikan", align='L')
                        pdf.cell(5, 5, ":")
                        pdf.cell(50, 5, "023456779")
                        
                        pdf.set_xy(140, y_before + 5)
                        pdf.cell(20, 5, "Tarikh :")
                        pdf.cell(30, 5, tarikh_str)
                        
                        pdf.set_xy(12, y_before + 12)
                        pdf.cell(40, 5, "Nama Majikan")
                        pdf.cell(5, 5, ":")
                        pdf.cell(50, 5, nama_majikan)
                        
                        pdf.set_xy(12, y_before + 19)
                        pdf.cell(40, 5, "Alamat Najikan")
                        pdf.cell(5, 5, ":")
                        
                        pdf.set_xy(57, y_before + 19)
                        pdf.multi_cell(80, 5, alamat_majikan)
                        
                        # Add Bil Pekerja right aligned
                        pdf.set_xy(150, y_before + 38)
                        pdf.cell(40, 5, f"Bil. Pekerja : {len(df)}", align='R')
                        
                        pdf.set_y(y_before + 50)
                        
                        # Table Header
                        pdf.set_font("Arial", 'B', 8)
                        
                        # Heights and Widths
                        col_w = [10, 25, 30, 75, 25, 25] # Total = 190
                        h = 10
                        
                        y_header = pdf.get_y()
                        pdf.set_xy(10, y_header)
                        pdf.cell(col_w[0], h, "BIL", border=1, align='C')
                        
                        x = 10 + col_w[0]
                        pdf.set_xy(x, y_header)
                        pdf.cell(col_w[1], h/2, "NO. AHLI", border='LTR', align='C')
                        pdf.set_xy(x, y_header + h/2)
                        pdf.cell(col_w[1], h/2, "KWSP", border='LBR', align='C')
                        
                        x += col_w[1]
                        pdf.set_xy(x, y_header)
                        pdf.cell(col_w[2], h/2, "NO. PENGENALAN", border='LTR', align='C')
                        pdf.set_xy(x, y_header + h/2)
                        pdf.cell(col_w[2], h/2, "DIRI", border='LBR', align='C')
                        
                        x += col_w[2]
                        pdf.set_xy(x, y_header)
                        pdf.cell(col_w[3], h/2, "NAMA PEKERJA 1 AHLI", border='LTR', align='C')
                        pdf.set_font("Arial", '', 7)
                        pdf.set_xy(x, y_header + h/2)
                        pdf.cell(col_w[3], h/2, "(Seperti terdapat dl dalam Dokumen)", border='LBR', align='C')
                        pdf.set_font("Arial", 'B', 8)
                        
                        x += col_w[3]
                        pdf.set_xy(x, y_header)
                        pdf.cell(col_w[4] + col_w[5], h/2, "CARUMAN (RM)", border=1, align='C')
                        pdf.set_xy(x, y_header + h/2)
                        pdf.cell(col_w[4], h/2, "MAJIKAN", border=1, align='C')
                        pdf.set_xy(x + col_w[4], y_header + h/2)
                        pdf.cell(col_w[5], h/2, "PEKERJA", border=1, align='C')
                        
                        pdf.set_y(y_header + h)
                        
                        # Table Body
                        pdf.set_font("Arial", '', 9)
                        
                        caruman_amount = float(salary) * 0.02
                        caruman_str = f"{caruman_amount:.2f}"
                        
                        total_majikan = 0.0
                        total_pekerja = 0.0
                        
                        for i, row in df.iterrows():
                            full_name = f"{row.get('First Name', '')} {row.get('Last Name', '')}".strip()
                            passport_no = str(row.get('Passport No', ''))
                            kwsp_no = ''.join(random.choices(string.digits, k=8))
                            
                            total_majikan += caruman_amount
                            total_pekerja += caruman_amount
                            
                            pdf.cell(col_w[0], 6, str(i+1), border=1, align='C')
                            pdf.cell(col_w[1], 6, kwsp_no, border=1, align='C')
                            pdf.cell(col_w[2], 6, passport_no, border=1, align='C')
                            pdf.cell(col_w[3], 6, full_name[:45], border=1, align='L')
                            pdf.cell(col_w[4], 6, caruman_str, border=1, align='R')
                            pdf.cell(col_w[5], 6, caruman_str, border=1, align='R')
                            pdf.ln()
                            
                        # Totals
                        pdf.set_font("Arial", 'B', 9)
                        pdf.cell(col_w[0]+col_w[1]+col_w[2]+col_w[3], 6, "JUMLAH (RM)", border=1, align='R')
                        pdf.cell(col_w[4], 6, f"{total_majikan:.2f}", border=1, align='R')
                        pdf.cell(col_w[5], 6, f"{total_pekerja:.2f}", border=1, align='R')
                        pdf.ln()
                        
                        pdf.cell(col_w[0]+col_w[1]+col_w[2]+col_w[3]+col_w[4], 6, "JUMLAH BESAR (RM)", border=1, align='R')
                        pdf.cell(col_w[5], 6, f"{total_majikan + total_pekerja:.2f}", border=1, align='R')
                        pdf.ln(10)
                        
                        # Footer Notes
                        pdf.set_font("Arial", 'B', 8)
                        pdf.cell(0, 5, "CATATAN :", ln=True)
                        pdf.set_font("Arial", '', 8)
                        pdf.multi_cell(0, 4, "* Caruman tidak dapat dimasukkan ke akaun ahli kerana maklumat tidak lengkap. Sila rujuk Jadual Caruman Tanpa Maklumat Lengkap (KWSP 1314).")
                        pdf.multi_cell(0, 4, "** Caruman tidak dapat dimasukkan ke akaun ahli kerana akaun ahli telah ditutup.")
                        
                        pdf.ln(2)
                        pdf.set_font("Arial", 'B', 8)
                        pdf.cell(0, 5, "PERHATIAN :", ln=True)
                        pdf.set_font("Arial", '', 8)
                        
                        notes = [
                            "1. Majikan hendaklah menyemak dan memastikan maklumat yang dicetak adalah betul. Jika terdapat kesilapan pada penyata ini, majikan hendaklah memberitahu KWSP secara bertulis dalam tempoh satu bulan dari tarikh penyata ini. Jika tiada bantahan diterima, caruman yang telah dimasukkan ke akaun ahli dianggap betul dan majikan adalah bertanggungjawab terhadap sebarang kesilapan.",
                            "2. Ini adalah Penyata Caruman KWSP yang telah diproses. Penyata Caruman ini tidak boleh digunakan untuk pembayaran caruman KWSP. Bagi tujuan penghantaran Borang Caruman dan/atau bayaran, sila gunakan kemudahan e-Caruman (i-Akaun Majikan) atau lain-lain kemudahan secara elektronik.",
                            "3. Caruman bulanan yang kena dibayar oleh majikan dan pekerja hendaklah tidak kurang daripada caruman minimum berdasarkan Perintah Caji Minimum yang berkuatkuasa selaras dengan Akta Majlis Perundingan Gaji Negara 2011."
                        ]
                        
                        for note in notes:
                            pdf.multi_cell(0, 4, note)
                            pdf.ln(1)
                            
                        pdf.ln(2)
                        pdf.set_font("Arial", 'I', 7)
                        pdf.cell(0, 4, "Surat ini adalah cetakan komputer dan tidak memerlukan tandatangan.", border='T')
                        
                        # Output
                        pdf_bytes = pdf.output(dest='S').encode('latin1')
                        
                        st.success("EPF PDF generated successfully!")
                        st.download_button(
                            label="⬇️ Download EPF PDF",
                            data=pdf_bytes,
                            file_name="EPF_Statement.pdf",
                            mime="application/pdf"
                        )
        except Exception as e:
            st.error(f"An error occurred: {e}")
