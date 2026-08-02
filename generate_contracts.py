import pandas as pd
from fpdf import FPDF

# 1. Load the Excel file
df = pd.read_excel("add 15 workers.xlsx")

# 2. Define the contract template with placeholders
template_text = """EMPLOYMENT CONTRACT
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

# 3. Loop through each row in the Excel file and generate a PDF
for index, row in df.iterrows():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    
    # Format the template with the data from the current Excel row
    contract_content = template_text.format(
        name=row['Name'],
        nationality=row['Nationality'],
        passport=row['Passport No'],
        place_of_issue=row['Place of issue'],
        date_of_issue=row['Date of issue']
    )
    
    # Write the formatted text to the PDF
    pdf.multi_cell(0, 6, contract_content)
    
    # Save the new PDF document
    filename = f"Employment_Contract_{row['Name']}.pdf"
    pdf.output(filename)
    print(f"Successfully generated: {filename}")
