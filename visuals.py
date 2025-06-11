import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import os
from datetime import datetime

def get_analysis_period(all_data: dict):
    """
    Extract the analysis period (start and end dates) from the data
    Returns: tuple (start_date, end_date) as datetime objects
    """
    all_dates = []
    
    # Look for only the reported_at column
    for sheet_name, df in all_data.items():
        # Check if reported_at column exists
        if 'reported_at' in df.columns:
            try:
                # Convert to datetime if not already
                if not pd.api.types.is_datetime64_any_dtype(df['reported_at']):
                    dates = pd.to_datetime(df['reported_at'], errors='coerce')
                else:
                    dates = df['reported_at']
                
                # Remove NaT values and add to all_dates
                valid_dates = dates.dropna()
                if len(valid_dates) > 0:
                    all_dates.extend(valid_dates.tolist())
                    print(f"📅 Found {len(valid_dates)} dates in {sheet_name}.reported_at")
                    
            except Exception as e:
                print(f"⚠️  Could not parse dates from {sheet_name}.reported_at: {e}")
                continue
    
    if not all_dates:
        print("⚠️  No valid dates found in reported_at columns, using current month")
        # Fallback to current month if no dates found
        current_date = datetime.now()
        start_date = current_date.replace(day=1)
        end_date = current_date
        return start_date, end_date
    
    # Convert to pandas datetime if not already
    all_dates = pd.to_datetime(all_dates)
    
    # Get min and max dates
    start_date = all_dates.min()
    end_date = all_dates.max()
    
    print(f"📅 Analysis period: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")
    
    return start_date, end_date

def create_trend_chart(df, output_path: str):
    """
    Create a trend chart showing attack patterns over weekly periods.
    
    Args:
        df: DataFrame containing 'reported_at', 'count', and 'indicator_ip' columns
        output_path: Path where the chart image will be saved
    """
    # Make a copy to avoid modifying the original dataframe
    df = df.copy()
    
    # Convert to datetime format
    df['reported_at'] = pd.to_datetime(df['reported_at'])

    # Convert to numeric values only
    df['count'] = pd.to_numeric(df['count'], errors='coerce')

    # Filter data 
    start_april = pd.to_datetime('2025-04-01')
    end_april = pd.to_datetime('2025-04-30')
    df_april = df[(df['reported_at'] >= start_april) & (df['reported_at'] <= end_april)]

    # If no data in April 2025, return without creating chart
    if df_april.empty:
        return

    # Define weekly date ranges
    week_ranges = {
        'Week 1': (pd.to_datetime('2025-04-01'), pd.to_datetime('2025-04-06')),
        'Week 2': (pd.to_datetime('2025-04-07'), pd.to_datetime('2025-04-13')),
        'Week 3': (pd.to_datetime('2025-04-14'), pd.to_datetime('2025-04-20')),
        'Week 4': (pd.to_datetime('2025-04-21'), pd.to_datetime('2025-04-27')),
        'Week 5': (pd.to_datetime('2025-04-28'), pd.to_datetime('2025-04-30'))
    }

    # Lists to store data for plotting
    weeks = []
    unique_ips = []
    total_hits = []

    for week_label, (start_date, end_date) in week_ranges.items():
        # Filter data for the current week
        df_current_week = df_april[(df_april['reported_at'] >= start_date) & (df_april['reported_at'] <= end_date)]

        # Calculate unique IPs
        weeks.append(week_label)
        unique_ips.append(df_current_week['indicator_ip'].nunique())

        # Calculate total hit counts (sum of 'count' column)
        total_hits.append(df_current_week['count'].sum())

    # Create a DataFrame for plotting
    plot_df = pd.DataFrame({
        'Week': weeks,
        'Unique IP Count': unique_ips,
        'Total Hit Count': total_hits
    })

    # Create the figure and a set of subplots
    fig, ax1 = plt.subplots(figsize=(12, 7))

    # Plot Unique IP Count as a bar chart on the left y-axis
    sns.barplot(x='Week', y='Unique IP Count', data=plot_df, ax=ax1, color='darkblue', label='Unique IP Count')
    ax1.set_xlabel('Weeks')
    ax1.set_ylabel('Unique IP Count', color='darkblue')
    ax1.tick_params(axis='y', labelcolor='darkblue')
    ax1.set_title('Attack Trends')

    # Create a second y-axis for Total Hit Count (line chart)
    ax2 = ax1.twinx()
    sns.lineplot(x='Week', y='Total Hit Count', data=plot_df, ax=ax2, color='red', marker='o', label='Total Hit Count')
    ax2.set_ylabel('Total Hit Count', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Combine legends from both axes
    lines, labels = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines + lines2, labels + labels2, loc='upper right')

    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    # Save the chart instead of showing it
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()


def create_network_traffic_chart(df, output_path: str):
    """
    Create a pie chart showing network traffic distribution by attacked ports.
    
    Args:
        df: DataFrame containing 'Attacked_Port' column
        output_path: Path where the chart image will be saved
    """
    # Make a copy to avoid modifying the original dataframe
    df = df.copy()
    
    # Check if 'Attacked_Port' column exists
    if 'Attacked_Port' not in df.columns:
        return
    
    # Replace 'NA' strings with actual NaN values for proper handling
    df['Attacked_Port'] = df['Attacked_Port'].replace('NA', pd.NA)

    # Drop rows where 'Attacked_Port' is NaN before splitting
    # This ensures we only process valid port strings
    df_cleaned_ports = df.dropna(subset=['Attacked_Port'])
    
    # Check if we have any data left
    if df_cleaned_ports.empty:
        return

    # Split the 'Attacked_Port' strings by '&&' and explode the list to new rows
    # Then strip any whitespace from the resulting port numbers
    all_individual_ports = df_cleaned_ports['Attacked_Port'].str.split('&&').explode().str.strip()

    # Convert ports to numeric, coercing errors to NaN. This will also handle empty strings if any
    # And then drop any NaN values that result from conversion errors (e.g., if a split results in an empty string)
    all_individual_ports = pd.to_numeric(all_individual_ports, errors='coerce').dropna()

    # Check if we have any valid ports
    if all_individual_ports.empty:
        return

    # Convert the ports to integer type for consistency
    all_individual_ports = all_individual_ports.astype(int)

    # Count the occurrences of each unique port
    port_counts = all_individual_ports.value_counts()

    # Define the number of top ports to display
    num_top_ports = 10

    # Get the top N major ports
    major_ports = port_counts.head(num_top_ports)

    # Calculate the sum of the remaining (other) ports
    other_ports_count = port_counts.iloc[num_top_ports:].sum()

    # Create a new Series for the pie chart data
    # Only add 'Other' if there are indeed other ports to sum
    if other_ports_count > 0:
        pie_chart_data = pd.concat([major_ports, pd.Series({'Other': other_ports_count})])
    else:
        pie_chart_data = major_ports

    # Create the pie chart
    plt.figure(figsize=(10, 10))
    plt.pie(pie_chart_data, labels=pie_chart_data.index, autopct='%1.1f%%', startangle=90,
            pctdistance=0.85, labeldistance=1.05)  # autopct to show percentages, startangle for 90 degrees
    plt.title('Network Traffic by Protocol')
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    
    # Save the chart instead of showing it
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()