# GSTScope

Automated GST filing status checker using Selenium, Flask, and Excel.

## Overview

GSTScope is a Python-based automation tool designed to check the GST filing status of multiple clients from the GST Portal.

The application takes client details through an Excel template, automates the GST Portal workflow using Selenium, and generates a month-wise Excel report containing GSTR-3B and GSTR-1/IFF filing statuses.

CAPTCHA verification is handled manually during the automation process.

## Features

- Upload client details through Excel
- Select the required financial year
- Automated GST Portal login using Selenium
- Manual CAPTCHA entry
- GSTIN-based taxpayer search
- Month-wise GSTR-3B filing status
- Month-wise GSTR-1/IFF filing status
- Automatic logout after processing
- Live processing progress
- Skip individual clients
- Pause and resume processing
- Save partial progress during execution
- Download results while processing
- Generate structured Excel reports
- Downloadable Excel input template

## How It Works

1. Prepare the client Excel file using the provided template.
2. Upload the Excel file to GSTScope.
3. Select the required financial year.
4. Start the automation.
5. Enter the CAPTCHA manually when prompted.
6. GSTScope searches for the taxpayer using the provided GSTIN.
7. The application retrieves the GSTR-3B and GSTR-1/IFF filing details.
8. Filing statuses are processed month-wise.
9. Results are displayed in the application.
10. The final report can be downloaded as an Excel file.

## Input Format

The input Excel file should contain the following columns:

| Column | Description |
|---|---|
| GST Portal User ID | GST Portal login user ID |
| GST Portal Password | GST Portal login password |
| Client Name | Name or identifier of the client |
| GSTIN | GST Identification Number |

A blank input template is provided in the repository as:

`input_template.xlsx`


## Output

GSTScope generates an Excel report containing:

- User ID
- GSTIN
- Client Name
- Month-wise GSTR-3B filing status
- Month-wise GSTR-1/IFF filing status

The report follows the financial year structure from April to March.

## Pause & Resume

GSTScope supports session persistence during processing.

If the automation is stopped or interrupted, the application saves the current progress and can resume processing from the saved point instead of starting again from the beginning.

## Technology Stack

- **Python**
- **Selenium**
- **Flask**
- **Pandas**
- **OpenPyXL**
- **WebDriver Manager**
- **HTML / CSS / JavaScript**

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/Anshika-Dubey/GSTScope.git
cd GSTScope
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the application

```bash
python gst_checker.py
```

The GSTScope web interface will be available through the local browser application.
