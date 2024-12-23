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
        'date': 'Date',
        'clicks': 'Total Clicks',
        'estimated_earnings': 'Revenue'
    }, inplace=True)
    revenue_agg['Date'] = pd.to_datetime(revenue_agg['Date'])

    # Extract campaign IDs from spend file
    spends_df['Camp ID'] = spends_df['Ad set name'].str.extract(r'\((\d+)\)').astype(int)
    spends_df['Date'] = pd.to_datetime(spends_df['Day'])

    # Merge spend and revenue data
    merged_data = spends_df.merge(revenue_agg, how='left', on=['Camp ID', 'Date'])
    merged_data['RPC'] = merged_data['Revenue'] / merged_data['Total Clicks']
    merged_data['RPC'] = merged_data['RPC'].replace([float('inf'), -float('inf')], 0).fillna(0)
    merged_data['Profit/Loss'] = merged_data['Revenue'] - merged_data['Spend']

    # Rename columns and reorder them
    merged_data.rename(columns={
        'Ad set name': 'Camp Name',
        'Amount spent (USD)': 'Spend',
        'Cost per purchase': 'CPR'
    }, inplace=True)

    column_order = ['Camp ID', 'Camp Name', 'Date', 'Spend', 'Revenue', 'CPR', 'RPC', 'Profit/Loss']
    merged_data = merged_data[column_order]

    # Create Excel file in memory
    # (Keep the rest of your file creation logic here)
