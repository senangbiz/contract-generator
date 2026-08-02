import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile

st.set_page_config(page_title="Contract Generator", page_icon="📄")

st.title("📄 Worker Contract Generator")
st.write("Upload your Excel file to generate employment contracts.")

# Default template from the script
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

# Text area so the user can see and optionally modify the template
with st.expander("View / Edit Contract Template"):
    template_text = st.text_area("Template Text", value=DEFAULT_TEMPLATE, height=400)

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Load the excel file
    try:
        df = pd.read_excel(uploaded_file)
        
        # Clean up column names by removing asterisks and extra spaces
        df.columns = [str(col).replace('*', '').strip() for col in df.columns]
        
        st.success("File uploaded successfully! Here is a preview of your data:")
        st.dataframe(df.head())
        
        # Check for required columns
        required_columns = ['First Name', 'Last Name', 'Nationality', 'Passport No', 'Passport Issue Date']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
        else:
            if st.button("Generate PDF Contracts"):
                with st.spinner("Generating contracts..."):
                    # Create an in-memory zip file to store all the generated PDFs
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for index, row in df.iterrows():
                            # Create PDF
                            pdf = FPDF()
                            pdf.add_page()
                            pdf.set_font("Arial", size=11)
                            
                            full_name = f"{row['First Name']} {row['Last Name']}"
                            # Format template
                            contract_content = template_text.format(
                                name=full_name,
                                nationality=row['Nationality'],
                                passport=row['Passport No'],
                                place_of_issue=row['Nationality'],
                                date_of_issue=row['Passport Issue Date']
                            )
                            
                            # Write formatted text to PDF
                            pdf.multi_cell(0, 6, contract_content)
                            
                            # Get the PDF content as a string (latin1 encoding is standard for fpdf 1.7)
                            pdf_bytes = pdf.output(dest='S').encode('latin1')
                            
                            filename = f"Employment_Contract_{full_name}.pdf"
                            # Add to zip file
                            zf.writestr(filename, pdf_bytes)
                    
                    st.success("All contracts generated successfully!")
                    
                    # Provide a download button for the zip file
                    st.download_button(
                        label="⬇️ Download All Contracts (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="contracts.zip",
                        mime="application/zip"
                    )
    except Exception as e:
        st.error(f"An error occurred: {e}")
