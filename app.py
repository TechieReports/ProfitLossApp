import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

def process_spend_and_revenue(spend_file, revenue_file):
    spends_df = pd.read_excel(spend_file)
    revenue_raw_df = pd.read_csv(revenue_file)

    # Aggregate raw revenue data
    revenue_agg = revenue_raw_df.groupby(['campid', 'date'], as_index=False).agg({
        'clicks': 'sum',
        'estimated_earnings': 'sum'
    })
    revenue_agg.rename(columns={
        'campid': 'Camp ID',
        'date': 'Day',
        'clicks': 'Total Clicks',
        'estimated_earnings': 'Revenue'
    }, inplace=True)
    revenue_agg['Day'] = pd.to_datetime(revenue_agg['Day'])

    # Extract campaign IDs from spend file
    spends_df['Camp ID'] = spends_df['Ad set name'].str.extract(r'\((\d+)\)').astype(int)
    spends_df['Day'] = pd.to_datetime(spends_df['Day'])

    # Merge spend and revenue data
    merged_data = spends_df.merge(revenue_agg, how='left', on=['Camp ID', 'Day'])
    merged_data['RPC'] = merged_data['Revenue'] / merged_data['Total Clicks']
    merged_data['RPC'] = merged_data['RPC'].replace([float('inf'), -float('inf')], 0).fillna(0)
    merged_data['Profit/Loss'] = merged_data['Revenue'] - merged_data['Amount spent (USD)']

    # Create Excel file in memory
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Profit Loss Analysis"

    # Add headers with formatting
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True)

    for col, header in enumerate(merged_data.columns, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font

    # Add data rows
    for row in dataframe_to_rows(merged_data, index=False, header=False):
        ws.append(row)

    # Conditional formatting for Profit/Loss
    profit_loss_col_index = merged_data.columns.get_loc('Profit/Loss') + 1
    for row in ws.iter_rows(min_row=2, min_col=profit_loss_col_index, max_col=profit_loss_col_index):
        for cell in row:
            if cell.value is not None:
                if cell.value > 0:
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                elif cell.value < 0:
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.title("Profit/Loss Analysis Tool")
st.write("Upload Spend and Raw Revenue Files to Generate a Profit/Loss Report.")

spend_file = st.file_uploader("Upload Spend File (Excel)", type=["xlsx"])
revenue_file = st.file_uploader("Upload Raw Revenue File (CSV)", type=["csv"])

if spend_file and revenue_file:
    st.write("Processing...")
    output_file = process_spend_and_revenue(spend_file, revenue_file)
    st.success("Processing complete!")
    st.download_button(
        label="Download Profit/Loss Report",
        data=output_file,
        file_name="Profit_Loss_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
