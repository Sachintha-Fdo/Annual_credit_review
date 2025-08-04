
import sys
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO
import numpy as np
from email_config import EmailSender
import logging

# Accept parameters from command line or use defaults
excel_file = sys.argv[1] if len(sys.argv) > 1 else 'data/Evaluation_Report.xlsx'
output_pdf = sys.argv[2] if len(sys.argv) > 2 else 'Auto_finance_annual_review_report.pdf'

# Load the Excel file
df = pd.read_excel(excel_file)
print('Excel columns:', list(df.columns))

# Convert numeric columns 
df['VALUATION'] = pd.to_numeric(df['VALUATION'], errors='coerce')
df['ROUNDED_VALUE'] = pd.to_numeric(df['ROUNDED_VALUE'], errors='coerce')
df['NO_REN_IN_ARREARS'] = pd.to_numeric(df['NO_REN_IN_ARREARS'], errors='coerce')

# Filter for blank REPORT_REVIEW_DATE and non-blank PRE_APPROVED_DATE
df_blank_review = df[df['REPORT_REVIEW_DATE'].isna() | (df['REPORT_REVIEW_DATE'] == '')] if 'REPORT_REVIEW_DATE' in df.columns else df
filtered_df = df_blank_review[~(df_blank_review['PRE_APPROVED_DATE'].isna() | (df_blank_review['PRE_APPROVED_DATE'] == ''))] if 'PRE_APPROVED_DATE' in df.columns else df_blank_review

# Define target products for Auto Finance
target_products = [
    'VEHICLE LOAN-REGISTERED', 'TRACTOR LEASE', 'PLEDGE LOAN',
    'OTHER LEASE', 'Murabaha', 'MINI TRUCK LEASE', 'IJARAH LEASE',
    'HIRE PURCHASE-UN-REGISTERED', 'HIRE PURCHASE-REGISTERED', 'VEHICLE LOAN-UN-REGISTERED'
]

grouped = filtered_df.groupby('PRODUCT')
product_dfs = {product: group.copy() for product, group in grouped}
dfs_to_merge = [product_dfs[key] for key in target_products if key in product_dfs]

# Check if dfs_to_merge is empty (no target products found)
if not dfs_to_merge:
    print("No target products found in the filtered data.")
    print("Sending email notification...")
    
    try:
        # Initialize email sender
        email_sender = EmailSender()
        
        # Define filter criteria for the email
        filter_criteria = """- Blank REPORT_REVIEW_DATE
- Non-blank PRE_APPROVED_DATE
- Target products: VEHICLE LOAN-REGISTERED, TRACTOR LEASE, PLEDGE LOAN, OTHER LEASE, Murabaha, MINI TRUCK LEASE, IJARAH LEASE, HIRE PURCHASE-UN-REGISTERED, HIRE PURCHASE-REGISTERED, VEHICLE LOAN-UN-REGISTERED"""
        
        # Send no-data notification using the proper method
        success = email_sender.send_no_data_notification("Auto Finance", filter_criteria)
        
        if success:
            print("[SUCCESS] Email notification sent successfully")
        else:
            print("[ERROR] Failed to send email notification")
            
    except Exception as e:
        print(f"Error sending email notification: {e}")
    
    # Exit the script without generating PDF
    sys.exit(0)

AF_df = pd.concat(dfs_to_merge, ignore_index=True)

# Check if AF_df is empty
if AF_df.empty:
    print("No records found for Auto Finance annual credit review with the filtered data.")
    print("Sending email notification...")
    
    try:
        # Initialize email sender
        email_sender = EmailSender()
        
        # Define filter criteria for the email
        filter_criteria = """- Blank REPORT_REVIEW_DATE
- Non-blank PRE_APPROVED_DATE
- Target products: VEHICLE LOAN-REGISTERED, TRACTOR LEASE, PLEDGE LOAN, OTHER LEASE, Murabaha, MINI TRUCK LEASE, IJARAH LEASE, HIRE PURCHASE-UN-REGISTERED, HIRE PURCHASE-REGISTERED, VEHICLE LOAN-UN-REGISTERED"""
        
        # Send no-data notification using the proper method
        success = email_sender.send_no_data_notification("Auto Finance", filter_criteria)
        
        if success:
            print("[SUCCESS] Email notification sent successfully")
        else:
            print("[ERROR] Failed to send email notification")
            
    except Exception as e:
        print(f"Error sending email notification: {e}")
    
    # Exit the script without generating PDF
    sys.exit(0)

print(f"Auto Finance DataFrame shape: {AF_df.shape}")
print("Proceeding with report generation...")

# Calculate DEVIATION if not present
if 'DEVIATION' not in AF_df.columns:
    AF_df['DEVIATION'] = ((AF_df['VALUATION'] - AF_df['ASSET_VALUATION']) / AF_df['VALUATION']).replace([float('inf'), -float('inf')], pd.NA)
    AF_df['DEVIATION'] = (AF_df['DEVIATION'] * 100).round(2)
    AF_df['DEVIATION'] = AF_df['DEVIATION'].map(lambda x: f"{x:.2f}%" if pd.notna(x) else '')

# === DEVIATION BINS ===


AF_df['_DEVIATION_FLOAT'] = abs(AF_df['DEVIATION'].str.replace('%', '').astype(float) / 100)
AF_df['_DEVIATION_FLOAT'] = pd.to_numeric(AF_df['_DEVIATION_FLOAT'], errors='coerce')


# Deviation Ranges
bins = [0.5, 0.6, 0.7, 0.8, 0.9, 1.000001, 100]
labels = ['50% – 60%', '60% – 70%', '70% – 80%', '80 – 90%', '90% – 100%', 'Above 100%']
AF_df['DEVIATION_RANGE'] = pd.cut(AF_df['_DEVIATION_FLOAT'], bins=bins, labels=labels, right=False, include_lowest=True)

# Deviation Summary
deviation_summary = AF_df['DEVIATION_RANGE'].value_counts(sort=False, dropna=True).reset_index()
deviation_summary.columns = ['DEVIATION_RANGE', 'COUNT']

# Summary by PRE_APPROVED_USER
pre_approved_summary = AF_df['PRE_APPROVED_USER'].value_counts(dropna=False).reset_index()
pre_approved_summary.columns = ['PRE APPROVED USER', 'Count']

# Select only the required columns
required_columns = [
    'FACILITY_NUMBER', 'PRODUCT', 'FACILITY_AMT', 'ASSET_DESCRIPTION',
    'MAKE_DESCRIPTION', 'MODEL_DESCRIPTION', 'YOM',
    'ASSET_VALUATION', 'VALUATION', 'PRE_APPROVED_USER', 'DEVIATION', 'ROUNDED_VALUE', 
    'PRE_APPROVED_AMT', 'REVIEW_RATING', 'NO_REN_IN_ARREARS', '_DEVIATION_FLOAT'
]

AF_df = AF_df[required_columns]  ##  + ['_DEVIATION_FLOAT']

# Extreme ROUNDED_VALUE
max_val = AF_df['ROUNDED_VALUE'].max()
min_val_filtered = AF_df[AF_df['ROUNDED_VALUE'] > 1000]['ROUNDED_VALUE']
min_val = min_val_filtered.min() if not min_val_filtered.empty else None

highest_df = AF_df[AF_df['ROUNDED_VALUE'] == max_val].copy()
lowest_df = AF_df[AF_df['ROUNDED_VALUE'] == min_val].copy()
highest_df['High/Low'] = 'H'
lowest_df['High/Low'] = 'L'
extreme_values_df = pd.concat([highest_df, lowest_df], ignore_index=True)

# Extreme VALUATION
max_valuation = AF_df['VALUATION'].max()
min_val_filtered_v = AF_df[AF_df['VALUATION'] > 10000]['VALUATION']
min_valuation = min_val_filtered_v.min() if not min_val_filtered_v.empty else None

highest_valuation_df = AF_df[AF_df['VALUATION'] == max_valuation].copy()
lowest_valuation_df = AF_df[AF_df['VALUATION'] == min_valuation].copy()
highest_valuation_df['High/Low'] = 'H'
lowest_valuation_df['High/Low'] = 'L'
valuation_extremes_df = pd.concat([highest_valuation_df, lowest_valuation_df], ignore_index=True)

# High Deviation Facilities (DEVIATION_FLOAT > 0.5)

high_deviation_df = AF_df[(AF_df['_DEVIATION_FLOAT'] > 0.5) & (AF_df['VALUATION'] > 10000)].copy()

if not high_deviation_df.empty:
    high_deviation_df = high_deviation_df.sort_values('_DEVIATION_FLOAT', ascending=False)
    # Remove _DEVIATION_FLOAT column for display
    high_deviation_df = high_deviation_df.drop('_DEVIATION_FLOAT', axis=1)
else:
    # Create empty DataFrame with same columns if no high deviation facilities
    high_deviation_df = AF_df.drop('_DEVIATION_FLOAT', axis=1).head(0)

# ----------------------------- Annual Credit Review Rating Summary -----------------------------
# Define conditions for Rating
conditions = [
    AF_df['PRE_APPROVED_AMT'] < 0,
    (AF_df['REVIEW_RATING'] == 'Green') & (AF_df['NO_REN_IN_ARREARS'] <= 1),
    (AF_df['REVIEW_RATING'] == 'Green') & (AF_df['NO_REN_IN_ARREARS'] <= 2),
    (AF_df['REVIEW_RATING'] == 'Green') & (AF_df['NO_REN_IN_ARREARS'] > 2),
    (AF_df['REVIEW_RATING'] == 'Yellow') & (AF_df['NO_REN_IN_ARREARS'] <= 1),
    (AF_df['REVIEW_RATING'] == 'Yellow') & (AF_df['NO_REN_IN_ARREARS'] <= 2),
    (AF_df['REVIEW_RATING'] == 'Yellow') & (AF_df['NO_REN_IN_ARREARS'] > 2),
    (AF_df['REVIEW_RATING'] == 'Orange') & (AF_df['NO_REN_IN_ARREARS'] <= 1),
    (AF_df['REVIEW_RATING'] == 'Orange') & (AF_df['NO_REN_IN_ARREARS'] > 1),
    (AF_df['REVIEW_RATING'] == 'Red') & (AF_df['NO_REN_IN_ARREARS'] <= 1),
    (AF_df['REVIEW_RATING'] == 'Red') & (AF_df['NO_REN_IN_ARREARS'] > 1),
]
# Define corresponding values
choices = ['Poor Negative Exposure', 'Lead Gen - Excellent', 'Affinity Gen - Satisfactory', 'Poor Repayment to CDB',
            'Affinity Gen - Satisfactory', 'Affinity Gen - Satisfactory',
           'Poor Repayment to CDB', 'Poor Repayment to other FIs', 'Poor Overall Repayment',
           'Poor Repayment to other FIs', 'Poor Overall Repayment']

# Apply the conditions
AF_df['Rating'] = np.select(conditions, choices, default='')

# Group by 'Rating' and aggregate the count and sum of 'FACILITY_AMT'
Rating = AF_df.groupby('Rating').agg(
    Count=('Rating', 'size'),
    Total_FACILITY_AMT=('FACILITY_AMT', 'sum')
).reset_index()

# Calculate total count and total sum of FACILITY_AMT for percentage calculation
total_count = Rating['Count'].sum()
total_facility_amt = Rating['Total_FACILITY_AMT'].sum()

# Calculate percentage based on Count
Rating['Count_Percentage %'] = ((Rating['Count'] / total_count) * 100).round(2).astype(str) + '%'

# Calculate percentage based on Total_FACILITY_AMT
Rating['FACILITY_AMT_Percentage %'] = ((Rating['Total_FACILITY_AMT'] / total_facility_amt) * 100).round(2).astype(str) + '%'

# ----------------------------- Annual Credit Review Grade Summary -----------------------------
# Define conditions for Grade
Rate_conditions = [
    (AF_df['Rating'] == 'Affinity Gen - Satisfactory'),
    (AF_df['Rating'] == 'Lead Gen - Excellent'),
    (AF_df['Rating'] == 'Poor Negative Exposure'),
    (AF_df['Rating'] == 'Poor Overall Repayment'),
    (AF_df['Rating'] == 'Poor Repayment to CDB'),
    (AF_df['Rating'] == 'Poor Repayment to other FIs'),
]

# Define corresponding values
Rate_choices = ['Satisfactory', 'Excellent', 'Average', 'Poor',
            'Average', 'Average']

# Apply the conditions
AF_df['Grade'] = np.select(Rate_conditions, Rate_choices, default='')

# Group by 'Grade' and aggregate the count and sum of 'FACILITY_AMT'
Grade_result = AF_df.groupby('Grade').agg(
    Count=('Grade', 'size'),
    Total_FACILITY_AMT=('FACILITY_AMT', 'sum')
).reset_index()

# Calculate percentage based on Count
Grade_result['Count_Percentage %'] = ((Grade_result['Count'] / total_count) * 100).round(2).astype(str) + '%'

# Calculate percentage based on Total_FACILITY_AMT
Grade_result['FACILITY_AMT_Percentage %'] = ((Grade_result['Total_FACILITY_AMT'] / total_facility_amt) * 100).round(2).astype(str) + '%'

# ----------------------------- PDF Report Generation -----------------------------
styles = getSampleStyleSheet()
WIDER_SIZE = (1300, 595.27)
doc = SimpleDocTemplate(output_pdf, pagesize=WIDER_SIZE, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
elements = []

# Paragraph style for cell content
cell_style = ParagraphStyle(name='CellStyle', fontSize=7, leading=7, wordWrap='CJK')

def add_table(title, df, col_widths=None):
    elements.append(Paragraph(f"<b>{title}</b>", styles['Heading3']))
    elements.append(Spacer(2, 10))

    safe_df = df.astype(str).replace('nan', '')
    data = [[Paragraph(str(cell), cell_style) for cell in row] for row in [safe_df.columns.tolist()] + safe_df.values.tolist()]

    if col_widths is None:
        col_widths = [68] * len(safe_df.columns)

    table = Table(data, repeatRows=1, hAlign='LEFT', colWidths=col_widths)
    
    # Define table style
    table_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.2, colors.grey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]
    
    # Add bold formatting for totals row if it exists
    if len(safe_df) > 0 and safe_df.iloc[-1, 0] == 'TOTAL':
        table_style.append(('FONTNAME', (0, len(safe_df)), (-1, len(safe_df)), 'Helvetica-Bold'))
        table_style.append(('FONTSIZE', (0, len(safe_df)), (-1, len(safe_df)), 9))
    
    table.setStyle(TableStyle(table_style))
    elements.append(table)
    elements.append(Spacer(1, 20))

# Helper function to format columns with commas

def format_with_commas(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) and x == x and x != '' else x)
    return df

# Helper function to add totals row
def add_totals_row(df):
    totals_row = {}
    for col in df.columns:
        if col in ['Count', 'Total_FACILITY_AMT']:
            totals_row[col] = df[col].sum()
        elif col in ['Count_Percentage %', 'FACILITY_AMT_Percentage %']:
            totals_row[col] = '100.00%'
        else:
            totals_row[col] = 'TOTAL'
    
    totals_df = pd.DataFrame([totals_row])
    return pd.concat([df, totals_df], ignore_index=True)

# Add the tables
# Add totals to the specified tables
pre_approved_summary = add_totals_row(pre_approved_summary)
Rating = add_totals_row(Rating)
Grade_result = add_totals_row(Grade_result)

# Format columns before adding tables
format_with_commas(extreme_values_df, ['ASSET_VALUATION', 'VALUATION', 'ROUNDED_VALUE', 'FACILITY_AMT', 'Total_FACILITY_AMT'])
format_with_commas(valuation_extremes_df, ['ASSET_VALUATION', 'VALUATION', 'ROUNDED_VALUE', 'FACILITY_AMT', 'Total_FACILITY_AMT'])
format_with_commas(Rating, ['FACILITY_AMT', 'Total_FACILITY_AMT'])
format_with_commas(Grade_result, ['FACILITY_AMT', 'Total_FACILITY_AMT'])
format_with_commas(high_deviation_df, ['ASSET_VALUATION', 'VALUATION', 'ROUNDED_VALUE', 'FACILITY_AMT'])

table1_widths = [80, 85, 65, 90, 85, 90, 40, 90, 90, 100, 60, 80, 55]
add_table("1. Highest & Lowest Pre-Approved Amount", extreme_values_df, col_widths=table1_widths)

table2_widths = [80, 85, 65, 90, 85, 90, 40, 90, 90, 100, 60, 80, 55]
add_table("2. Highest & Lowest VALUATION Records", valuation_extremes_df, col_widths=table2_widths)

table3_widths = [120, 80]
add_table("3. Deviation Summary", deviation_summary, col_widths=table3_widths)

table4_widths = [200, 80]
add_table("4. PRE_APPROVED_USER Summary", pre_approved_summary, col_widths=table4_widths)

# Add Rating Summary table
Rating_table_widths = [120, 80, 100, 100, 100]
add_table("5. Annual Credit Review Rating Summary", Rating, col_widths=Rating_table_widths)

# Add Grade Summary table
Grade_table_widths = [120, 80, 100, 100, 100]
add_table("6. Annual Credit Review Grade Summary", Grade_result, col_widths=Grade_table_widths)

# Add High Deviation Facilities table
if not high_deviation_df.empty:
    high_deviation_table_widths = [80, 85, 65, 90, 85, 90, 40, 90, 90, 100, 60, 80, 55]
    add_table("7. High Deviation Facilities (DEVIATION_FLOAT > 50%)", high_deviation_df, col_widths=high_deviation_table_widths)

# Build PDF
doc.build(elements)
print(f"PDF report saved as: {output_pdf}")
