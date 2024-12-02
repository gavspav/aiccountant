# AIccountant

AIccountant is a Python-based tool that helps you consolidate and analyze your financial transactions from multiple sources including Gmail, Amazon, PayPal, and bank statements.

## Components

### 1. Gmail Parser (`gmail_parser.py`)
This script connects to your Gmail account and searches for transaction-related emails. It:
- Extracts transaction details like dates, amounts, and vendors
- Supports various email formats from different vendors
- Exports the data to a CSV file for further processing

### 2. Transaction Consolidator (`transaction_consolidator.py`)
This script combines transaction data from multiple sources:
- Processes CSV files from Gmail, Amazon, PayPal, and bank statements
- Identifies potential duplicate transactions across sources
- Exports consolidated data to an Excel file with color-coded duplicates

## Installation

1. Clone the repository:
```bash
git clone [your-repository-url]
cd aiccountant
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install pandas openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

## Gmail Setup

To allow the Gmail Parser to access your emails, you need to:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable the Gmail API:
   - Go to "APIs & Services" > "Library"
   - Search for "Gmail API"
   - Click "Enable"
4. Set up OAuth 2.0 credentials:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Choose "Desktop app" as application type
   - Download the client configuration file
   - Rename it to `credentials.json` and place it in the project directory

## Usage

1. First, parse Gmail transactions:
```bash
python gmail_parser.py
```
- On first run, this will open a browser window for Gmail authentication
- Follow the prompts to allow access
- The script will create `gmail_transactions.csv`

2. Prepare your other transaction files:
- Save your Amazon order history as `amazon_order_history.csv`
- Export your PayPal transactions to `paypal.csv`
- Export your bank statements to `bank.csv`

3. Consolidate all transactions:
```bash
python transaction_consolidator.py
```
- This will create `consolidated_transactions.xlsx`
- Potential duplicate transactions will be color-coded
- Review the Excel file to identify and resolve duplicates

## File Format Requirements

### Amazon Order History CSV
Required columns:
- `date`: Order date
- `total`: Order total
- `items`: Description of items

### PayPal CSV
Required columns:
- `Date`: Transaction date
- `Name`: Transaction description
- `Amount`: Transaction amount
- `Status`: Transaction status
- `Type`: Transaction type

### Bank Statement CSV
Required columns:
- `Date`: Transaction date
- `Description`: Transaction description
- `Amount`: Transaction amount

## Notes
- The Gmail parser searches for transaction-related emails from the last 30 days by default
- Transactions marked as "pending" will be excluded from the final output
- All dates are standardized to UTC before processing
- The consolidator looks for potential duplicates within a 7-day window

## Troubleshooting

1. Gmail Authentication Issues:
   - Ensure `credentials.json` is in the correct location
   - Check that the Gmail API is enabled in your Google Cloud Console
   - Delete `token.json` and re-authenticate if needed

2. Date Parsing Errors:
   - Ensure your CSV files use consistent date formats
   - Check for any "pending" transactions that might cause issues

3. Excel Export Errors:
   - Make sure no other programs have the Excel file open
   - Verify that all required columns are present in your input files

## Contributing
Feel free to submit issues and enhancement requests!
