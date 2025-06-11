from data_processor import read_excel_data
from report_generator import generate_docx_report, get_organization_info
from visuals import create_trend_chart, create_network_traffic_chart, get_analysis_period
from config import OPENAI_API_KEY, LLM_MODEL
import openai
import re
import os


def build_combined_prompt(all_data: dict):
    combined_rows = []
    for sheet_name, df in all_data.items():
        records = df.to_dict(orient="records")
        limited_records = records[:15]
        combined_rows.append(f"\n--- {sheet_name} ---")
        for i, row in enumerate(limited_records):
            combined_rows.append(f"{i+1}. {row}")
    combined_data_text = "\n".join(combined_rows)
    
    prompt = f"""
You are a cybersecurity analyst working for Sequretek Labs.

You are provided with honeypot logs captured across multiple dimensions of cybersecurity telemetry. Below is the combined snapshot of this data:

{combined_data_text}

Your task is to analyze this data and generate a complete security report. Write substantial analysis content under each section header. Each section should contain 2-3 paragraphs of analysis.

Format your response EXACTLY like this:

1. Attack Indicators
Based on the honeypot data analysis, identify the critical attack indicators.

2. Honeypot Attack Trends
The honeypot deployment has captured significant attack activity during the monitoring period... [Write detailed paragraphs about timing patterns across different weeks highlighting the changes. Mention the description of each week with its name and what changed in the subsequent week and why that happened with respect to unique IP counts and total hit count and other things if you find so. Dont mentions any attributes from the data itself instead give an analysis based on the data]

3. Network Traffic by Protocol
Analysis of network protocols reveals... [Write detailed paragraphs about most targeted protocols]

4. Indicator of Attacks
Just Type - "Indicators are given below".

5. Top IP Addresses
The most active attacking IP addresses demonstrate... [Give ONLY a table for the top 20 IP addresses with respect to severity of attacks(should include highest to lowest severity). In the table also include severity and action of each value]

6. Credential Patterns
Attack credential patterns reveal... [Identify username/password patterns, common combinations, brute force attempts. Create a table for top 6 most comomon usernames that have been attacked along with the protocol service they have been attacked on and then create a table for top 6 passwords that have been attacked along with the protocol service they have been attacked on]

7. Subdomains
Subdomain enumeration and targeting shows... [Give a table for the top 20 subdomains and the number of attacks they have been involved in]

8. Email Addresses
Email-related attack vectors include... [Give a table for the top 20 email addresses and the number of attacks they have been involved in]

9. Hashes
Malware hash analysis indicates... [Give a table for the top 5 hashes and the number of attacks they have been involved in]

IMPORTANT: Write substantial content under each numbered section. Do not just list the headers. Analyze the actual data provided above.
"""
    return prompt


def call_llm(prompt: str):
    if not OPENAI_API_KEY:
        print("❌ OpenAI API key not configured")
        return "[Error: API key not configured]"
    
    # Set the API key
    openai.api_key = OPENAI_API_KEY
    
    try:
        response = openai.chat.completions.create(
            model=LLM_MODEL,
            messages=[
                {"role": "system", "content": "You are a cybersecurity analyst writing structured reports."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=4000
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"❌ OpenAI API request failed: {e}")
        return "[Error: Failed to generate report.]"


def parse_markdown_table(table_text):
    """Parse a markdown table into a list of dictionaries"""
    lines = [line.strip() for line in table_text.split('\n') if line.strip()]
    if len(lines) < 3:  # Need at least header, separator, and one data row
        return None
    
    # Extract headers
    header_line = lines[0]
    if not header_line.startswith('|') or not header_line.endswith('|'):
        return None
    
    headers = [h.strip() for h in header_line.split('|')[1:-1]]
    
    # Skip separator line (line 1)
    data_rows = []
    for line in lines[2:]:
        if line.startswith('|') and line.endswith('|'):
            row_data = [cell.strip() for cell in line.split('|')[1:-1]]
            if len(row_data) == len(headers):
                row_dict = dict(zip(headers, row_data))
                data_rows.append(row_dict)
    
    return data_rows if data_rows else None


def extract_tables_from_text(text):
    """Extract markdown tables from text and return (clean_text, tables)"""
    lines = text.split('\n')
    clean_lines = []
    current_table_lines = []
    tables = []
    in_table = False
    
    for line in lines:
        stripped_line = line.strip()
        
        # Check if this line looks like a table row
        if stripped_line.startswith('|') and stripped_line.endswith('|') and '|' in stripped_line[1:-1]:
            if not in_table:
                in_table = True
                current_table_lines = []
            current_table_lines.append(stripped_line)
        elif stripped_line.startswith('|') and ('---' in stripped_line or '====' in stripped_line):
            # This is a table separator line
            if in_table:
                current_table_lines.append(stripped_line)
        else:
            # Not a table line
            if in_table:
                # End of table, process it
                if len(current_table_lines) >= 3:  # Header + separator + at least one row
                    table_text = '\n'.join(current_table_lines)
                    parsed_table = parse_markdown_table(table_text)
                    if parsed_table:
                        tables.append(parsed_table)
                current_table_lines = []
                in_table = False
            
            # Add non-table line to clean text
            if stripped_line:  # Only add non-empty lines
                clean_lines.append(line)
    
    # Handle case where text ends with a table
    if in_table and len(current_table_lines) >= 3:
        table_text = '\n'.join(current_table_lines)
        parsed_table = parse_markdown_table(table_text)
        if parsed_table:
            tables.append(parsed_table)
    
    clean_text = '\n'.join(clean_lines)
    return clean_text, tables


def parse_report_sections(report_text):
    # Define the section headers as they appear in the report
    section_names = [
        "Attack Indicators",
        "Honeypot Attack Trends",
        "Network Traffic by Protocol", 
        "Indicator of Attacks",
        "Top IP Addresses",
        "Credential Patterns",
        "Subdomains",
        "Email Addresses",
        "Hashes"
    ]
    
    sections = {}
    section_tables = {}
    
    # Try a simpler approach: split by numbered sections
    lines = report_text.split('\n')
    current_section = None
    current_content = []
    
    for line in lines:
        line = line.strip()
        if not line:
            if current_content:
                current_content.append('')
            continue
            
        # Check if this line is a section header (starts with number and matches our sections)
        found_section = None
        for section_name in section_names:
            # Look for patterns like "1. Attack Indicators", "1. **Attack Indicators**", or just the section name
            patterns_to_check = [
                f"1. {section_name}",
                f"2. {section_name}",
                f"3. {section_name}",
                f"4. {section_name}",
                f"5. {section_name}",
                f"6. {section_name}",
                f"7. {section_name}",
                f"8. {section_name}",
                f"9. {section_name}",
                f"1. **{section_name}**",
                f"2. **{section_name}**",
                f"3. **{section_name}**",
                f"4. **{section_name}**",
                f"5. **{section_name}**",
                f"6. **{section_name}**",
                f"7. **{section_name}**",
                f"8. **{section_name}**",
                f"9. **{section_name}**",
                section_name
            ]
            
            if (any(line.startswith(pattern) for pattern in patterns_to_check) or
                line.strip() == section_name or
                line.strip().endswith(section_name)):
                found_section = section_name
                break
        
        if found_section:
            # Save previous section content
            if current_section and current_content:
                raw_text = '\n'.join(current_content).strip()
                clean_text, tables = extract_tables_from_text(raw_text)
                sections[current_section] = clean_text
                section_tables[current_section] = tables if tables else None
            
            # Start new section
            current_section = found_section
            current_content = []
        else:
            # Add content to current section
            if current_section:
                current_content.append(line)
    
    # Don't forget the last section
    if current_section and current_content:
        raw_text = '\n'.join(current_content).strip()
        clean_text, tables = extract_tables_from_text(raw_text)
        sections[current_section] = clean_text
        section_tables[current_section] = tables if tables else None
    
    # Ensure all expected sections are present
    for name in section_names:
        if name not in sections:
            sections[name] = f"No content found for {name} section."
            section_tables[name] = None
    
    return sections, section_tables


def main():
    print("📥 Loading honeypot data from all sheets...")
    all_data = read_excel_data("data/Honeypot_Dummy_Data.xlsx")
    
    # Create output directory if it doesn't exist
    os.makedirs("output", exist_ok=True)
    
    # Extract analysis period from the data
    print("📅 Extracting analysis period from data...")
    analysis_start, analysis_end = get_analysis_period(all_data)
    
    # Generate charts for sheets that have the required columns
    print("📊 Generating charts...")
    images = {}
    network_charts = {}
    
    for sheet_name, df in all_data.items():
        # Check if the dataframe has the required columns for trend chart
        required_cols = ['reported_at', 'count', 'indicator_ip']
        if all(col in df.columns for col in required_cols):
            chart_path = f"output/{sheet_name}_trend_chart.png"
            create_trend_chart(df, chart_path)
            images[sheet_name] = chart_path
        else:
            images[sheet_name] = None
        
        # Check if the dataframe has the required column for network traffic chart
        if 'Attacked_Port' in df.columns:
            network_chart_path = f"output/{sheet_name}_network_traffic_chart.png"
            create_network_traffic_chart(df, network_chart_path)
            network_charts[sheet_name] = network_chart_path
        else:
            network_charts[sheet_name] = None
    
    print("🧠 Generating AI analysis...")
    prompt = build_combined_prompt(all_data)
    full_report = call_llm(prompt)

    print("📝 Processing report sections...")
    content_sections, tables = parse_report_sections(full_report)
    
    # Map images to the specific sections that should have charts
    section_images = {}
    
    # Find the first available trend chart
    trend_chart_path = None
    for sheet_name, chart_path in images.items():
        if chart_path:
            trend_chart_path = chart_path
            break
    
    # Find the first available network traffic chart
    network_chart_path = None
    for sheet_name, chart_path in network_charts.items():
        if chart_path:
            network_chart_path = chart_path
            break
    
    # Assign images to specific sections
    for section_name in content_sections.keys():
        if "Honeypot Attack Trends" in section_name:
            section_images[section_name] = trend_chart_path
        elif "Network Traffic by Protocol" in section_name:
            section_images[section_name] = network_chart_path
        elif "Attack Trends" in section_name:
            section_images[section_name] = trend_chart_path
        else:
            section_images[section_name] = None
    
    # Get organization info from logos folder
    print("🏢 Detecting organization info from logos folder...")
    org_name, org_logo_path = get_organization_info("logos")
    
    print("📄 Writing report to DOCX file...")
    generate_docx_report(content_sections, section_images, tables, "output/generated_report.docx", 
                        analysis_dates=(analysis_start, analysis_end), logo_folder_path="logo/Meta.jpg")
    print("✅ Report saved at: output/generated_report.docx")


if __name__ == "__main__":
    main()