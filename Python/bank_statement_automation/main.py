import pandas as pd
import os
from datetime import datetime
from rules import categorize_transaction, get_category_summary

def load_bank_statement(file_path):
    """Load bank statement CSV file"""
    try:
        df = pd.read_csv(file_path)
        print(f"✅ Successfully loaded {len(df)} transactions from {file_path}")
        return df
    except Exception as e:
        print(f"❌ Error loading file: {e}")
        return None

def clean_columns(df):
    """Clean and standardize columns"""
    # Create unified 'Amount' column
    def calculate_amount(row):
        debit = row['Debit'] if not pd.isna(row['Debit']) else 0
        credit = row['Credit'] if not pd.isna(row['Credit']) else 0
        return credit - debit
    
    df['Amount'] = df.apply(calculate_amount, axis=1)
    print("✅ Created unified Amount column")
    
    # Parse dates
    try:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        print("✅ Parsed dates successfully")
    except Exception as e:
        print(f"❌ Date parsing failed: {e}")
    
    return df

def categorize_transactions(df):
    """Categorize transactions based on description keywords"""
    df['Category'] = df['Description'].apply(categorize_transaction)
    print("✅ Categorized transactions")
    
    # Print category summary
    category_counts = df['Category'].value_counts()
    print("\n📊 Category Summary:")
    for category, count in category_counts.items():
        print(f"  {category}: {count} transactions")
    
    return df

def export_clean_data(df, output_path):
    """Export cleaned data to Excel"""
    try:
        # Create output directory if it doesn't exist
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Export to Excel
        df.to_excel(output_path, index=False)
        print(f"✅ Cleaned data exported to {output_path}")
        
        # Also export a summary report
        summary_path = output_path.replace('.xlsx', '_summary.xlsx')
        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            # Main data
            df.to_excel(writer, sheet_name='Transactions', index=False)
            
            # Category summary
            category_summary = df.groupby('Category').agg({
                'Amount': ['count', 'sum', 'mean']
            }).round(2)
            category_summary.columns = ['Transaction Count', 'Total Amount', 'Average Amount']
            category_summary.to_excel(writer, sheet_name='Category Summary')
            
            # Monthly summary
            df['Month'] = df['Date'].dt.to_period('M')
            monthly_summary = df.groupby('Month').agg({
                'Amount': ['sum', 'count']
            }).round(2)
            monthly_summary.columns = ['Total Amount', 'Transaction Count']
            monthly_summary.to_excel(writer, sheet_name='Monthly Summary')
        
        print(f"✅ Summary report exported to {summary_path}")
        
    except Exception as e:
        print(f"❌ Error exporting file: {e}")

def main():
    """Main function to run the bank statement automation"""
    print("🏦 Bank Statement Automation Tool")
    print("=" * 40)
    
    # File paths
    input_file = 'data/example_bank_export.csv'
    output_file = 'output/cleaned_output.xlsx'
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"❌ Input file not found: {input_file}")
        print("Please place your bank statement CSV file in the data/ folder")
        return
    
    # Load data
    df = load_bank_statement(input_file)
    if df is None:
        return
    
    # Clean data
    df = clean_columns(df)
    
    # Categorize transactions
    df = categorize_transactions(df)
    
    # Export cleaned data
    export_clean_data(df, output_file)
    
    print("\n🎉 Bank statement processing complete!")
    print(f"📁 Check the output/ folder for your cleaned files")

if __name__ == "__main__":
    main() 