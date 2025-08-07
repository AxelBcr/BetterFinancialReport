# Better Visuals For Your EasyBourse Financial Report

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-4.15+-green.svg)
![License](https://img.shields.io/badge/License-GPL%20v3-green.svg)

<h3>Automated portfolio data extraction and visualization system for EasyBourse with Power BI integration</h3>

[**Features**](#key-features) • [**Demo**](#demo) • [**Installation**](#installation) • [**Usage**](#usage) • [**Documentation**](#documentation)

</div>

---

## Overview

This project provides a complete **automated solution** for extracting portfolio valuation data from EasyBourse platform, consolidating historical data in Excel, and creating advanced visualizations through Power BI. It eliminates manual data collection while maintaining a comprehensive historical database of your investments.

### Why This Project?

- **Save Time**: Automate hours of manual data entry
- **Track History**: Build a complete historical database of your portfolio evolution
- **Better Insights**: Power BI provides superior analytics compared to native EasyBourse interface
- **Real-time Updates**: Automatic hourly refresh when running
- **Data Security**: All data stays local on your machine

## Key Features

- **Automated Web Scraping** - Selenium-based secure connection and data extraction
- **Virtual Keyboard Support** - Handles EasyBourse's security features automatically
- **Historical Data Management** - Intelligent merging of new data with existing records
- **Automatic Backups** - Maintains the last 10 versions for data safety
- **Power BI Integration** - Direct connection for advanced analytics
- **Continuous Updates** - Hourly automatic refresh while running
- **Headless Mode** - Runs silently in the background

<div align="center">

## Demo

### 1. Automated Data Extraction

The script automatically connects to EasyBourse using Selenium WebDriver in headless mode.  
Once authenticated, it directly accesses the real-time CSV export URL and downloads the complete portfolio valuation data.  
The entire process is logged in the console, showing connection status, download progress, and data parsing steps.

<div align="center">
<img src="README_Data/CMD_Script.gif" style="max-width: 100%; width="600"/>
<br><br>
</div>

---

### 2. Power BI Dashboard Update

After successful data extraction, the script processes the CSV file and updates the Excel database with new information.   
The Excel file serves as the data source for Power BI, maintaining full historical records while adding new valuation dates.   
After opening the PowerBi file (.pbix) you need to make sure that you refresh the database.

<div align="center">
<img src="README_Data/Update_PowerBi.gif" style="max-width: 100%; width="600"/>
<br><br>
</div>

---

### 3. Interactive Dashboard Analysis

The Power BI dashboard transforms raw portfolio data into actionable insights.   
Users can explore their portfolio through multiple perspectives:   
value over time, sector allocation pie charts, individual position, gain/loss and more.     
   
Interactive filters allow deep dives into specific time periods or asset classes.

<div align="center">
<img src="README_Data/Use_PowerBi.gif" style="max-width: 100%; width="600"/>
<br><br>
</div>

</div>

## Installation

### Prerequisites

- Python 3.8 or higher
- Google Chrome (latest version)
- ChromeDriver (compatible with Chrome version)
- Power BI Desktop
- Active EasyBourse account   
 
### Quick Setup   
---
0. **Download the repo as a zip file and extract**   
   Or follow the next steps
---
1. **Clone the repository**
```bash
git clone https://github.com/AxelBcr/BetterFinancialReport.git
cd BetterFinancialReport
```

2. **Install dependencies**
```bash
# Using the provided script
Install_Requirements.bat

# Or manually
pip install -r requirements.txt
```

3. **Configure credentials**
```python
# Edit logins.py
id = "your_easybourse_id"
password = "your_password"
```

4. **Install ChromeDriver**
   - Check your Chrome version: `chrome://version/`
   - Download from [ChromeDriver](https://chromedriver.chromium.org/)
   - Add to system PATH

## Usage

### Standard Operation

Simply double-click `Update_Dashboard.bat` to:
1. Launch the extraction process
2. Update the Excel database
3. Continue running with hourly updates (can be changed in the .bat)

### Manual Execution

```bash
python easybourse_valorisation.py
```

### First Time Setup

On first run, the script will:
- Create the Excel database (`EasyBourse.xlsx`),   
  I recommend not deleting the .xlsx that is already in the repo,   
  just delete the entries in it so it's empty for you.
- Set up the backup folder (`Save/`)
- Extract your complete current portfolio

## Documentation

### Project Structure

```
BetterFinancialReport/
│
├── easybourse_valorisation.py      #Extraction script
├── logins.py                       #Storing Id and Password here
├── requirements.txt                #Required library to install
├── Update_Dashboard.bat            #.bat file to automate the extraction
├── Install_Requirements.bat        #.bat file to easily install requirements
├── EasyBourse.xlsx                 #Excel database  
├── EasyBourse.pbix                 #PowerBi report       
├── Save/                           #Folder where the last 10 Excel database are saved    
└── README_Data/                    #Just storing GIFs for the README
```
---

### Data Structure

| Column | Type | Description |
|--------|------|-------------|
| **Date** | DateTime | Valuation date |
| **Valeur** | String | Security name |
| **Code Isin** | String | International identifier |
| **Quantité** | Float | Number of shares |
| **Cours** | Float | Current price |
| **Prix moyen** | Float | Average purchase price |
| **Valorisation** | Float | Total position value |
| **+/- value** | Float | Unrealized P&L |
| **Performance (%)** | Float | Return percentage |
| **Valeur totale** | Float | Portfolio total value |
| **Solde espèces** | Float | Cash balance |
| **Total positions sous dossier** | Float | Sum of all positions |

---

<div align="center">

**Note:** This project is not affiliated with EasyBourse, La Banque postale, or any financial institution.

</div>
