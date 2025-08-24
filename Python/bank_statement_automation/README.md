# 🏦 Bank Statement Automation Tool

A Python tool that automatically cleans, categorizes, and analyzes bank statement CSV exports, making financial reconciliation and reporting a breeze!

## ✨ Features

- **📊 CSV Import**: Reads messy bank statement CSV files
- **🧹 Data Cleaning**: Standardizes dates, amounts, and descriptions
- **🏷️ Auto-Categorization**: Automatically categorizes transactions using keyword matching
- **📈 Smart Analysis**: Generates category summaries and monthly reports
- **📤 Excel Export**: Exports clean data and summary reports to Excel
- **⚙️ Customizable Rules**: Easy to add new categorization rules

## 🚀 Quick Start

### 1. Setup

```bash
# Clone or download the project
cd bank_statement_automation

# Install dependencies
pip install -r requirements.txt
```

### 2. Prepare Your Data

Place your bank statement CSV file in the `data/` folder. The tool expects columns:
- `Date`: Transaction date
- `Description`: Transaction description
- `Debit`: Debit amount (leave empty for credits)
- `Credit`: Credit amount (leave empty for debits)
- `Balance`: Account balance

### 3. Run the Tool

```bash
python main.py
```

### 4. Check Results

The tool will create:
- `output/cleaned_output.xlsx` - Clean transaction data
- `output/cleaned_output_summary.xlsx` - Summary reports with multiple sheets

## 📁 Project Structure

```
bank_statement_automation/
├── data/
│   └── example_bank_export.csv    # Sample data
├── output/
│   ├── cleaned_output.xlsx        # Clean transaction data
│   └── cleaned_output_summary.xlsx # Summary reports
├── main.py                        # Main automation script
├── rules.py                       # Categorization rules
├── requirements.txt               # Python dependencies
└── README.md                      # This file
```

## 🏷️ Built-in Categories

The tool automatically categorizes transactions into:

- **Food & Dining**: Starbucks, McDonald's, restaurants, delivery
- **Income**: Payroll, salary, deposits, bonuses
- **Housing**: Rent, mortgage, housing expenses
- **Shopping**: Amazon, Walmart, Target, retail stores
- **Transportation**: Uber, gas, parking, public transit
- **Utilities**: Electric, water, internet, phone bills
- **Entertainment**: Netflix, Spotify, movies, gym
- **Healthcare**: Doctor visits, pharmacy, medical expenses
- **Education**: Tuition, books, courses
- **Travel**: Hotels, flights, vacation expenses
- **Insurance**: Auto, home, health insurance
- **Investments**: Stocks, bonds, brokerage accounts

## ⚙️ Customization

### Adding New Categories

Edit `rules.py` to add custom categorization rules:

```python
# Add to CATEGORY_RULES dictionary
"Your Category": [
    "keyword1", "keyword2", "keyword3"
]
```

### Programmatically Adding Rules

```python
from rules import add_custom_rule

# Add new keywords to existing category
add_custom_rule("Food & Dining", ["new_restaurant", "catering"])

# Add completely new category
add_custom_rule("Pet Expenses", ["vet", "pet food", "grooming"])
```

## 📊 Output Files

### Main Transaction File
- Clean, categorized transaction data
- Unified amount column (positive for credits, negative for debits)
- Standardized date format

### Summary Report
- **Transactions Sheet**: All cleaned transaction data
- **Category Summary**: Count, total, and average by category
- **Monthly Summary**: Monthly totals and transaction counts

## 🔧 Advanced Usage

### Custom File Paths

Modify the file paths in `main.py`:

```python
input_file = 'path/to/your/bank_statement.csv'
output_file = 'path/to/your/output.xlsx'
```

### Batch Processing

To process multiple files, you can modify the script:

```python
import glob

# Process all CSV files in data folder
for csv_file in glob.glob('data/*.csv'):
    # Process each file
    process_bank_statement(csv_file)
```

## 🛠️ Troubleshooting

### Common Issues

1. **"Import pandas could not be resolved"**
   - Install dependencies: `pip install -r requirements.txt`

2. **"Input file not found"**
   - Make sure your CSV file is in the `data/` folder
   - Check the filename matches what's expected in `main.py`

3. **Date parsing errors**
   - Ensure your date column is in a recognizable format (YYYY-MM-DD, MM/DD/YYYY, etc.)

4. **Excel export errors**
   - Make sure you have write permissions in the output folder
   - Close any open Excel files before running


---
