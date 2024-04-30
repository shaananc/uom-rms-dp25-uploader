# RMS Budget Uploader Tool

This Python tool automates the uploading of budget data from an Excel file into a Research Management System (RMS). It reads data configurations from a YAML file, fetches and sums costs from the Excel spreadsheet, and then uploads the processed data to the RMS. This tool is particularly useful for managing financial data across multiple years and categories in academic and research settings.

It will prompt you to enter the RMS username and password, and then it will upload the data to the RMS.

## Requirements

- Python 3.6+
- A Selenium WebDriver (ChromeDriver or GeckoDriver) installed and added to the PATH or configured in the YAML file
- A completed University of Melbourne DP25 Budget Template in Excel format

## Installation

1. Install the required Python packages:

```bash
pip install -r requirements.txt
```

2. Download and install the appropriate WebDriver for your browser

## Configuration

See the `config.yml` file for the configuration options. You can specify the path to the Excel file, the WebDriver path, the RMS URL, and the data columns to read from the Excel file

## Usage

Run the script with:

```bash
python budget-arc.py
```

When prompted, enter your RMS username and password. The script will then upload the data to the RMS.

## Limitations

- The script currently only supports the University of Melbourne's RMS excel template, and may not work in future years without modification
- The script has only been tested with the following budget categories:
  - Personnel
  - Travel
  - Teaching Relief
  - Other
- You may have to click the "login" button manually if the script fails to do so
- The script may fail due to RMS timeouts or other issues, in which case you will have to restart the script
