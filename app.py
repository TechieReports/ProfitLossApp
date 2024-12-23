import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

def process_spend_and_revenue(spend_file, revenue_file, filters):
    spends_df = pd.read_excel(spend_file)
    revenue_raw_df = pd.read_csv(revenue_file)

    # Aggregate raw revenue data
    revenue_agg = revenue_raw_df.groupby(['campid', 'date'], as_index=False).agg({
        'clicks': 'sum',
        'estimated_earnings': 'sum'
    })
    revenue_agg.rename(columns={
        'campid': 'Camp ID',
        'date': 'Date',
        'clicks': 'Total Clicks',
        'estimated_earnings': 'Revenue'
    }, inplace=True)
    revenue_agg['Date'] = pd.to_datetime(revenue_agg['Date'])

    # Extract campaign IDs from spend file
    spends_df['Camp ID'] = spends_df['Ad set name'].str.extract(r'\((\d+)\)').astype(int)
    spends_df['Date'] = pd.to_datetime(spends_df['Day'])

    # Handle "Cost per result" and create "CPR"
    if 'Cost per result' in spends_df.columns:
        spends_df.rename(columns={'Cost per result': 'CPR'}, inplace=True)
    else:
        spends_df['CPR'] = 0  # Add CPR with default value if missing

    # Merge spend and revenue data
    merged_data = spends_df.merge(revenue_agg, how='left', on=['Camp ID', 'Date'])
    merged_data['RPC'] = merged_data['Revenue'] / merged_data['Total Clicks']
    merged_data['RPC'] = merged_data['RPC'].replace([float('inf'), -float('inf')], 0).fillna(0)
    merged_data['Profit/Loss'] = merged_data['Revenue'] - merged_data['Amount spent (USD)']

    # Apply filters
    if filters['start_date']:
        merged_data = merged_data[merged_data['Date'] >= filters['start_date']]
    if filters['end_date']:
        merged_data = merged_data[merged_data['Date'] <= filters['end_date']]
    if filters['campaign_ids']:
        merged_data = merged_data[merged_data['Camp ID'].isin(filters['campaign_ids'])]
    if filters['profit_loss_min'] is not None:
        merged_data = merged_data[merged_data['Profit/Loss'] >= filters['profit_loss_min']]
    if filters['profit_loss_max'] is not None:
        merged_data = merged_data[merged_data['Profit/Loss'] <= filters['profit_loss_max']]

    # Sort by Date (oldest to newest) and Camp ID (smallest to largest)
    merged_data.sort_values(by=['Date', 'Camp ID'], inplace=True)

    # Reorder columns
    column_order = ['Camp ID', 'Ad set name', 'Date', 'Amount spent (USD)', 'Revenue', 'CPR', 'RPC', 'Profit/Loss']
    merged_data = merged_data[column_order]

    # Create Excel file in memory
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Profit Loss Analysis"

    # Add headers with formatting
    headers = ['Camp ID', 'Campaign Name', 'Date', 'Spend', 'Revenue', 'CPR', 'RPC', 'Profit/Loss']
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True)

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font

    # Freeze header row
    ws.freeze_panes = ws['A2']

    # Add data rows and format the date column
    for row_idx, row in enumerate(dataframe_to_rows(merged_data, index=False, header=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            # Format Date column to "MMM DD"
            if col_idx == 3:  # Date is the 3rd column
                cell.number_format = "MMM DD"

    # Adjust column widths
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value is not None)
        adjusted_width = max_length + 2
        ws.column_dimensions[col[0].column_letter].width = adjusted_width

    # Conditional formatting for Profit/Loss
    profit_loss_col_index = headers.index('Profit/Loss') + 1
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
    st.write("Filters:")
    start_date = st.date_input("Start Date", value=None)
    end_date = st.date_input("End Date", value=None)
    campaign_ids = st.text_input("Filter by Campaign IDs (comma-separated):")
    profit_loss_min = st.number_input("Minimum Profit/Loss", value=None)
    profit_loss_max = st.number_input("Maximum Profit/Loss", value=None)

    # Parse Campaign IDs
    campaign_ids = [int(c.strip()) for c in campaign_ids.split(",") if c.strip().isdigit()]

    filters = {
        'start_date': pd.to_datetime(start_date) if start_date else None,
        'end_date': pd.to_datetime(end_date) if end_date else None,
        'campaign_ids': campaign_ids,
        'profit_loss_min': profit_loss_min,
        'profit_loss_max': profit_loss_max
    }

    if st.button("Process"):
        output_file = process_spend_and_revenue(spend_file, revenue_file, filters)
        st.success("Processing complete!")
        st.download_button(
            label="Download Profit/Loss Report",
            data=output_file,
            file_name="Profit_Loss_Analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
