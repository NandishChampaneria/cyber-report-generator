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

    # Get the date range from the data
    start_date = df['reported_at'].min()
    end_date = df['reported_at'].max()

    # If no valid dates, return without creating chart
    if pd.isna(start_date) or pd.isna(end_date):
        print(f"⚠️ No valid dates found in data for trend chart")
        return

    # Create weekly date ranges
    date_ranges = []
    current_date = start_date
    week_num = 1
    
    while current_date <= end_date:
        week_end = min(current_date + pd.Timedelta(days=6), end_date)
        date_ranges.append((f'Week {week_num}', (current_date, week_end)))
        current_date = week_end + pd.Timedelta(days=1)
        week_num += 1

    # Lists to store data for plotting
    weeks = []
    unique_ips = []
    total_hits = []

    for week_label, (week_start, week_end) in date_ranges:
        # Filter data for the current week
        df_current_week = df[(df['reported_at'] >= week_start) & (df['reported_at'] <= week_end)]

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

    # Remove the rectangle/box outline
    for spine in ax1.spines.values():
        spine.set_visible(False)

    # Add horizontal grid lines
    ax1.yaxis.grid(True, linestyle='--', linewidth=0.7, color='#cccccc', alpha=0.7)
    ax1.set_axisbelow(True)

    # Plot Unique IP Count as a thin, sleek bar chart on the left y-axis
    bar_width = 0.4
    bar_positions = range(len(plot_df['Week']))
    ax1.bar(bar_positions, plot_df['Unique IP Count'], width=bar_width, color='darkblue', label='Unique IP Count', zorder=3)
    ax1.set_xticks(bar_positions)
    ax1.set_xticklabels(plot_df['Week'], rotation=45, ha='right')
    ax1.set_xlabel('Weeks')
    ax1.set_ylabel('Unique IP Count', color='darkblue')
    ax1.tick_params(axis='y', labelcolor='darkblue')
    ax1.set_title('Attack Trends')

    # Create a second y-axis for Total Hit Count (thicker line chart)
    ax2 = ax1.twinx()
    ax2.plot(bar_positions, plot_df['Total Hit Count'], color='red', marker='o', label='Total Hit Count', linewidth=3, zorder=4)
    ax2.set_ylabel('Total Hit Count', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Combine legends from both axes
    lines, labels = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines + lines2, labels + labels2, loc='upper right')

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
        print(f"⚠️ 'Attacked_Port' column not found in data for network traffic chart")
        return
    
    # Common port to protocol mapping
    port_protocol_mapping = {
        21: 'FTP', 22: 'SSH', 23: 'TELNET', 25: 'SMTP', 53: 'DNS', 80: 'HTTP',
        110: 'POP3', 143: 'IMAP', 443: 'HTTPS', 993: 'IMAPS', 995: 'POP3S',
        135: 'RPC', 139: 'NetBIOS', 445: 'SMB', 3389: 'RDP', 1433: 'MSSQL',
        3306: 'MySQL', 5432: 'PostgreSQL', 6379: 'Redis', 27017: 'MongoDB',
        1521: 'Oracle', 161: 'SNMP', 162: 'SNMPTRAP', 69: 'TFTP', 123: 'NTP',
        389: 'LDAP', 636: 'LDAPS', 514: 'Syslog', 515: 'LPR', 587: 'SMTP-TLS',
        465: 'SMTPS', 990: 'FTPS', 992: 'TelnetS', 8080: 'HTTP-Alt',
        8443: 'HTTPS-Alt', 5060: 'SIP', 5061: 'SIPS', 1723: 'PPTP',
        1194: 'OpenVPN', 500: 'IPSec', 4500: 'IPSec-NAT', 1812: 'RADIUS',
        1813: 'RADIUS-Acct', 119: 'NNTP', 563: 'NNTPS', 993: 'IMAPS',
        220: 'IMAP3', 179: 'BGP', 520: 'RIP', 521: 'RIPng', 2049: 'NFS',
        111: 'Portmapper', 2000: 'Cisco-SCCP', 1701: 'L2TP', 1702: 'L2F',
        47: 'GRE', 1863: 'MSNP', 5222: 'XMPP', 5269: 'XMPP-Server',
        6667: 'IRC', 194: 'IRC', 6697: 'IRC-SSL', 5900: 'VNC', 5901: 'VNC',
        5902: 'VNC', 5903: 'VNC', 631: 'IPP', 9100: 'JetDirect',
        10000: 'Webmin', 8000: 'HTTP-Alt2', 8888: 'HTTP-Alt3', 9999: 'Abyss',
        7001: 'Cassandra', 9042: 'Cassandra', 9160: 'Cassandra',
        11211: 'Memcached', 50070: 'Hadoop', 9200: 'Elasticsearch',
        5984: 'CouchDB', 6379: 'Redis', 27017: 'MongoDB', 28017: 'MongoDB-Web'
    }
    
    # Replace 'NA' strings with actual NaN values for proper handling
    df['Attacked_Port'] = df['Attacked_Port'].replace('NA', pd.NA)

    # Drop rows where 'Attacked_Port' is NaN before splitting
    df_cleaned_ports = df.dropna(subset=['Attacked_Port'])
    if df_cleaned_ports.empty:
        print(f"⚠️ No valid port data found for network traffic chart")
        return

    # Split the 'Attacked_Port' strings by '&&' and explode the list to new rows
    all_individual_ports = df_cleaned_ports['Attacked_Port'].str.split('&&').explode().str.strip()
    all_individual_ports = pd.to_numeric(all_individual_ports, errors='coerce').dropna()
    if all_individual_ports.empty:
        print(f"⚠️ No valid numeric ports found for network traffic chart")
        return
    all_individual_ports = all_individual_ports.astype(int)

    port_counts = all_individual_ports.value_counts()
    num_top_ports = 10
    major_ports = port_counts.head(num_top_ports)
    other_ports_count = port_counts.iloc[num_top_ports:].sum()

    # Create labels with protocol names
    protocol_labels = []
    protocol_counts = []
    for port, count in major_ports.items():
        if port in port_protocol_mapping:
            protocol_labels.append(f"{port_protocol_mapping[port]}:{port}")
        else:
            protocol_labels.append(f"Port {port}")
        protocol_counts.append(count)
    if other_ports_count > 0:
        protocol_labels.append("Other Ports")
        protocol_counts.append(other_ports_count)

    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(12, 9), facecolor='white')
    ax.set_facecolor('white')

    # Use a rainbow colormap for the pie chart
    cmap = plt.cm.rainbow
    colors = [cmap(i / len(protocol_labels)) for i in range(len(protocol_labels))]

    wedges, texts, autotexts = plt.pie(
        protocol_counts,
        labels=protocol_labels,
        autopct=lambda pct: f'{pct:.1f}%' if pct > 3 else '',
        startangle=90,
        colors=colors,
        wedgeprops={'linewidth': 2, 'edgecolor': 'white'},
        pctdistance=0.7,
        labeldistance=1.15
    )

    for text in texts:
        text.set_fontsize(13)
        text.set_fontweight('bold')
        text.set_family('Arial')
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(11)
        autotext.set_family('Arial')

    plt.title('Network Traffic by Protocol', 
              fontsize=16, fontweight='bold', pad=25, 
              color='#2C3E50', family='Arial')
    ax.axis('equal')
    plt.tight_layout()
    plt.subplots_adjust(left=0.1, right=0.9, bottom=0.1, top=0.9)
    plt.savefig(output_path, dpi=300, bbox_inches='tight', 
                facecolor='white', edgecolor='none', format='png')
    plt.close()
    print(f"📊 Network traffic chart with protocols saved to {output_path}")

def create_data_distribution_chart(all_data: dict, output_path: str):
    """
    Create a pie chart showing the distribution of data across different categories.
    
    Args:
        all_data: Dictionary containing dataframes from all Excel files
        output_path: Path where the chart image will be saved
    """
    # Initialize counters
    category_counts = {}
    
    # Map file names to display categories
    file_category_mapping = {
        'IP': 'IP Addresses',
        'Domain': 'Subdomains', 
        'Email': 'Email Addresses',
        'Hash': 'Hashes',
    }
    
    # Count rows in each category
    for sheet_name, df in all_data.items():
        # Match the sheet name to our categories
        category_found = False
        for file_key, display_name in file_category_mapping.items():
            if file_key.lower() in sheet_name.lower():
                if display_name not in category_counts:
                    category_counts[display_name] = 0
                category_counts[display_name] += len(df)
                category_found = True
                print(f"📊 Found {len(df)} rows in {sheet_name} -> {display_name}")
                break
        
        if not category_found:
            print(f"⚠️ Unknown category for file: {sheet_name}")
    
    # Check if we have any data
    if not category_counts:
        print("⚠️ No data found for distribution chart")
        return
    
    # Create figure with white background
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(12, 8), facecolor='white')
    ax.set_facecolor('white')
    
    # Sort by count for better visualization
    sorted_categories = dict(sorted(category_counts.items(), key=lambda x: x[1], reverse=True))
    
    # Professional color palette - corporate blues and complementary colors
    colors = ['#48cae4', '#0077b6', '#03045e', '#caf0f8', '#6A994E', '#4A5D23']
    
    # Calculate percentages for labels
    total_records = sum(sorted_categories.values())
    percentages = [count/total_records * 100 for count in sorted_categories.values()]
    
    # Create pie chart with professional styling
    wedges, texts, autotexts = plt.pie(
        sorted_categories.values(),
        labels=None,  # We'll create custom labels
        autopct=lambda pct: f'{pct:.1f}%' if pct > 5 else '',  # Only show percentage if > 5%
        startangle=90,
        colors=colors[:len(sorted_categories)],
        wedgeprops={'linewidth': 2, 'edgecolor': 'white'},
        pctdistance=0.85
    )
    
    # Style the percentage text
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(11)
        autotext.set_family('Arial')
    
    # Add professional title
    plt.title('Attack Indicators', 
              fontsize=16, fontweight='bold', pad=25, 
              color='#2C3E50', family='Arial')
    
    # Create custom legend with better formatting
    legend_elements = []
    for i, (category, count) in enumerate(sorted_categories.items()):
        percentage = (count / total_records) * 100
        legend_label = f'{category}\n{count:,} records ({percentage:.1f}%)'
        legend_elements.append(plt.Rectangle((0,0),1,1, facecolor=colors[i], 
                                           edgecolor='white', linewidth=1))
    
    # Position legend to the right with better styling
    legend = plt.legend(legend_elements, 
                       [f'{cat}\n{count:,} records ({(count/total_records)*100:.1f}%)' 
                        for cat, count in sorted_categories.items()],
                       loc='center left',
                       bbox_to_anchor=(1.05, 0.5),
                       fontsize=10,
                       frameon=True,
                       fancybox=True,
                       shadow=True,
                       framealpha=0.9,
                       facecolor='#F8F9FA',
                       edgecolor='#E9ECEF')
    
    # Style the legend
    legend.get_frame().set_linewidth(1)
    
    
    # Ensure equal aspect ratio for circular pie
    ax.axis('equal')
    
    # Adjust layout to prevent clipping
    plt.tight_layout()
    plt.subplots_adjust(left=0.1, right=0.75, bottom=0.1, top=0.9)
    
    # Save with high quality
    plt.savefig(output_path, dpi=300, bbox_inches='tight', 
                facecolor='white', edgecolor='none', format='png')
    plt.close()
    
    print(f"📊 Professional data distribution chart saved to {output_path}")
    return sorted_categories