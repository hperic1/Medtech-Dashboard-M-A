# MedTech M&A & Investment Dashboard

An interactive dashboard for tracking and analyzing MedTech industry M&A activity and venture investments, with integrated JP Morgan report data and automated deal entry capabilities.

## Features

### 📊 Tab 1: Deal Activity Dashboard
- **Split-screen view** with separate tables for M&A Activity and Investment Activity
- **Interactive filtering** by quarter and month
- **Sortable columns** for easy data exploration
- **Top 3 Deals Cards** showing the largest deals by dollar amount
- **Quarterly charts** with:
  - Stacked bar charts for deal volume ($)
  - Line graph overlay for deal count
  - Outside end data labels for clarity

### 📈 Tab 2: JP Morgan Summary
- **YTD Overview Chart** showing quarterly segments for:
  - M&A Activity
  - Venture Investment
  - IPOs
  - Licensing Deals
- **Key Deals Highlights** extracted from JP Morgan reports
- **Automatic updates** when new reports are uploaded

### ➕ Tab 3: Data Entry & Management
Three methods for adding new deals:

1. **Article Scraping**
   - Paste URL or article text
   - Automatic extraction of deal information
   - Duplicate detection
   - Standardized formatting

2. **Manual Entry**
   - Forms for M&A and Investment deals
   - All standard columns available
   - Built-in validation

3. **JP Morgan Report Upload**
   - PDF processing capability
   - Automatic data extraction
   - Direct update to dashboard

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Setup

1. Clone this repository:
```bash
git clone https://github.com/yourusername/medtech-dashboard.git
cd medtech-dashboard
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Add your data file:
   - Place `MedTech_YTD_Standardized.xlsx` in the `data/` directory
   - Or upload it through the Streamlit interface

4. Run the dashboard locally:
```bash
streamlit run app.py
```

## Deploying to Streamlit Cloud

1. Push your code to GitHub:
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/yourusername/medtech-dashboard.git
git push -u origin main
```

2. Go to [share.streamlit.io](https://share.streamlit.io)

3. Sign in with GitHub

4. Click "New app"

5. Select your repository, branch (main), and main file (app.py)

6. Click "Deploy"

Your dashboard will be live at: `https://yourusername-medtech-dashboard-app-xxxxx.streamlit.app`

## Project Structure

```
medtech_dashboard/
├── app.py                          # Main Streamlit application
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .streamlit/
│   └── config.toml                # Streamlit configuration
├── data/
│   └── MedTech_YTD_Standardized.xlsx  # Your data file
└── utils/
    ├── __init__.py
    ├── data_loader.py             # Data loading and saving functions
    ├── charts.py                  # Chart generation functions
    └── scraper.py                 # Web scraping utilities
```

## Data Format

### M&A Activity Sheet
Required columns:
- Company
- Acquirer
- Deal Type (Merger / Acquisition)
- Technology/Description
- Deal Value
- Quarter (Q1, Q2, Q3, Q4)
- Month

### Investment Activity Sheet
Required columns:
- Company
- Funding type (VC / PE)
- Technology/Description
- Amount Raised
- Lead Investors
- Quarter (Q1, Q2, Q3, Q4)
- Month

## Usage Tips

### Filtering Data
- Use the dropdown filters at the top of each table
- Select "All" to view all data
- Combine quarter and month filters for precise views

### Viewing Charts
- Charts automatically update based on your data
- Hover over data points for detailed information
- Use the camera icon to download chart images

### Adding New Deals
1. **Quick Entry**: Use manual form for immediate data entry
2. **Batch Import**: Paste article text for multiple deal extraction
3. **URL Import**: Provide article URLs for automatic scraping

### Data Export
- Your data is automatically saved to the Excel file
- Download the updated file from the sidebar
- All changes are persistent

## Troubleshooting

### Dashboard won't load
- Ensure `MedTech_YTD_Standardized.xlsx` is in the `data/` directory
- Check that all dependencies are installed: `pip install -r requirements.txt`

### Charts not displaying
- Verify your data has valid numeric values in the deal value columns
- Check that Quarter column uses format: Q1, Q2, Q3, Q4

### Article scraping not working
- Some websites block automated scraping
- Try pasting the article text directly instead
- Check your internet connection

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License.

## Contact

For questions or support, please open an issue on GitHub.

## Acknowledgments

- Data powered by JP Morgan MedTech Industry Reports
- Built with Streamlit
- Charts generated using Plotly
