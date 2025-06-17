# Cyber-Security Honeypot Report Generator

Generate professional cybersecurity reports from honeypot data with custom branding and automated analysis.

## Features

- Automated report generation from Excel honeypot data
- Custom cover page with organization logo and name
- Table of Contents with section links
- Visualizations and tables for key metrics
- About page
- Professional header and footer with page numbers

## Project Structure

```
cyber-report-generator/
├── main.py
├── report_generator.py
├── data/
│   └── Honeypot_Dummy_Data.xlsx
├── logo/
│   └── Meta.jpg  # (example, add the organization logo here)
├── templates/
│   ├── sequretek.png
│   ├── base.png
│   └── template.docx
├── output/
├── requirements.txt
├── .gitignore
├── README.md
└── ...
```

## Setup

1. **Clone the repository:**
   ```
   git clone https://github.com/yourusername/cyber-report-generator.git
   cd cyber-report-generator
   ```
2. **Install dependencies:**
   ```
   pip install -r requirements.txt
   ```
3. **Add your OpenRouter API key:**
   - Create a file named `.env` in the project root directory.
   - Get your API key from [OpenRouter](https://openrouter.ai/keys)
   - Add your OpenRouter API key in this format:
     ```
     OPENROUTER_API_KEY=sk-or-v1-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
     ```
   - **Do not share or commit your .env file.**

## Usage

1. Place your honeypot Excel file in the `data/` folder (e.g., `data/Honeypot_Dummy_Data.xlsx`).
2. Place your organization logo in the `logo/` folder (e.g., `logo/Meta.jpg`).
3. Run the script:
   ```
   python main.py
   ```
4. The report will be generated in the `output/` folder as `generated_report.docx`.

**Important:**
- To update page numbers and the Table of Contents, open the generated `.docx` in Microsoft Word and press `Ctrl+A` then `F9`.
- The system uses the logo filename to determine the organization name automatically.
- Template images (Sequretek logo and base image) are in the `templates/` folder and used automatically.

## Sample Data

- Example honeypot data: `data/Honeypot_Dummy_Data.xlsx`
- Example logo: `logo/Meta.jpg`
- Template images: `templates/sequretek.png`, `templates/base.png`

## Known Limitations

- Page numbers and TOC in the DOCX are field codes and must be updated in Word (see above).
- Only Microsoft Word (or compatible editors) can update these fields automatically.

## License

MIT

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change. 