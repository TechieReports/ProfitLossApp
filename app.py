import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import datetime

def process_spend_and_revenue(spend_files, revenue_files, filters):
    spends_df = pd.concat([pd.read_excel(file) for file in spend_files])
    revenue_raw_df = pd.concat([pd.read_csv(file) for file in revenue_files])

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

    # Sort by Date (oldest to newest) and Camp ID (smallest to largest)
    merged_data.sort_values(by=['Date', 'Camp ID'], inplace=True)

    # Reorder columns
    column_order = ['Camp ID', 'Ad set name', 'Date', 'Amount spent (USD)', 'Revenue', 'CPR', 'RPC', 'Profit/Loss']
    merged_data = merged_data[column_order]

    # Aggregate data for top and bottom-performing campaigns
    aggregated_data = merged_data.groupby('Camp ID', as_index=False).agg({
        'Amount spent (USD)': 'sum',
        'Revenue': 'sum',
        'Profit/Loss': 'sum'
    })
    top_5_campaigns = aggregated_data.sort_values(by='Profit/Loss', ascending=False).head(5)
    bottom_5_campaigns = aggregated_data.sort_values(by='Profit/Loss', ascending=True).head(5)

    return merged_data, top_5_campaigns, bottom_5_campaigns

# Streamlit UI
st.title("Profit/Loss Analysis Tool with Enhanced Filters")
st.write("Upload Spend and Raw Revenue Files to Generate a Profit/Loss Report.")

spend_files = st.file_uploader("Upload Spend Files (Excel)", type=["xlsx"], accept_multiple_files=True)
revenue_files = st.file_uploader("Upload Revenue Files (CSV)", type=["csv"], accept_multiple_files=True)

if spend_files and revenue_files:
    st.write("Filters:")

    # Load the data to populate filter options
    spends_df = pd.concat([pd.read_excel(file) for file in spend_files])
    spends_df['Camp ID'] = spends_df['Ad set name'].str.extract(r'\((\d+)\)').astype(int)
    campaigns = spends_df['Camp ID'].unique()

    # Date range filter with default values
    default_start_date = datetime.date.today() - datetime.timedelta(days=30)
    default_end_date = datetime.date.today()

    start_date, end_date = st.date_input(
        "Select Date Range",
        value=(default_start_date, default_end_date)
    )

    # Campaign filter
    selected_campaigns = st.multiselect("Select Campaigns", options=campaigns)

    filters = {
        'start_date': pd.to_datetime(start_date) if start_date else None,
        'end_date': pd.to_datetime(end_date) if end_date else None,
        'campaign_ids': selected_campaigns
    }

    if st.button("Process"):
        filtered_data, top_5_campaigns, bottom_5_campaigns = process_spend_and_revenue(spend_files, revenue_files, filters)

        # Display filtered data
        st.subheader("Filtered Data")
        st.dataframe(filtered_data)

        # Display top 5 campaigns
        st.subheader("Top 5 Performing Campaigns")
        st.dataframe(top_5_campaigns)

        # Display bottom 5 campaigns
        st.subheader("Bottom 5 Performing Campaigns")
        st.dataframe(bottom_5_campaigns)

        # Download filtered data
        output = BytesIO()
        filtered_data.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button(
            label="Download Filtered Data",
            data=output,
            file_name="Filtered_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
s
