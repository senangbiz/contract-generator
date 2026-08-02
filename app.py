import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import random
import string

st.set_page_config(page_title="Generator Dashboard", page_icon="📄", layout="wide")

st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Tool:", ["Contract Generator", "Medical Report Generator"])

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
                                pdf.set_auto_page_break(auto=True, margin=15)
                                
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
                                
                                # Row 5
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Doctor Code")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": D4ES000306")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Physical Examination Date")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, ": 15/06/2026")
                                
                                # Row 6
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Transaction ID")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": 20240729860413")
                                
                                pdf.set_xy(col2_x, y)
                                pdf.cell(45, 6, "Certification Date")
                                pdf.set_xy(col2_val_x, y)
                                pdf.cell(45, 6, ": 16/06/2026")
                                
                                # Row 7
                                y += 8
                                pdf.set_xy(col1_x, y)
                                pdf.cell(40, 6, "Passport Expiry Date")
                                pdf.set_xy(col1_val_x, y)
                                pdf.cell(45, 6, ": -")
                                
                                pdf.set_y(y + 15)
                                
                                # 5. PART II
                                pdf.set_font("Arial", 'BU', 10)
                                pdf.cell(0, 6, "PART II. MEDICAL HISTORY", ln=True, align='C')
                                pdf.ln(5)
                                
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
                                
                                pdf.ln(5)
                                pdf.set_font("Arial", 'B', 8)
                                pdf.cell(0, 5, "Note:", ln=True)
                                pdf.set_font("Arial", 'B', 7)
                                pdf.cell(0, 4, "• 1. Foreign worker with a medical history of the Category 1 Diseases is deemed to be unsuitable for employment in Malaysia.", ln=True)
                                pdf.cell(0, 4, "• 2. However, foreign worker who gives a medical history of Hepatitis B, Hepatitis C, Syphilis, HIV, Malaria or Filariasis but does not", ln=True)
                                pdf.cell(0, 4, "show any clinical evidence of the above and the blood test results are negative, the foreign worker is deemed to be suitable for employment in Malaysia.", ln=True)
                                
                                # 6. PART V
                                pdf.ln(10)
                                pdf.set_font("Arial", 'BU', 10)
                                pdf.cell(0, 6, "PART V: CERTIFICATION BY DOCTOR (cont'd)", ln=True, align='C')
                                pdf.ln(5)
                                
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
                                pdf.rect(148, y + 4, 3, 3)
                                
                                # SUITABLE Checkbox (checked with a tick)
                                pdf.rect(178, y + 4, 3, 3)
                                pdf.set_xy(178, y + 2.5)
                                pdf.set_font("Arial", 'B', 7)
                                pdf.cell(3, 6, "v", align='C')
                                
                                pdf.set_y(y + 16)
                                y = pdf.get_y()
                                pdf.set_font("Arial", 'B', 8)
                                pdf.set_xy(10, y)
                                pdf.cell(10, 4, "24.")
                                pdf.set_xy(20, y)
                                pdf.multi_cell(110, 4, "Comments (refer to Part V - Item 16)\n-")
                                
                                # footer (Page: 1 of 7)
                                pdf.set_y(270)
                                pdf.set_font("Arial", 'I', 10)
                                pdf.cell(0, 10, "Page: 1 of 7", align='C')
                                
                                pdf_bytes = pdf.output(dest='S').encode('latin1')
                                filename = f"Medical_Report_{full_name}.pdf"
                                zf.writestr(filename, pdf_bytes)
                        
                        st.success("All medical reports generated successfully!")
                        st.download_button("⬇️ Download All Medical Reports (ZIP)", data=zip_buffer.getvalue(), file_name="medical_reports.zip", mime="application/zip")
        except Exception as e:
            st.error(f"An error occurred: {e}")
