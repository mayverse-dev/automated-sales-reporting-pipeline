import pandas as pd
import os
import re


class SalesReportGenerator:
    def __init__(self, input_path, output_path):
        self.input_path = input_path
        self.output_path = output_path
        self.df = None  # Placeholder for our data

    def load_data(self):
        """Loads data and prints initial stats."""
        print(f"🔄 Loading file: {self.input_path}...")
        try:
            # --- THE FIX IS HERE ---
            # on_bad_lines='skip': If a row is broken, skip it instead of crashing.
            # engine='python': A more flexible reading engine than the default 'c'.
            self.df = pd.read_csv(
                self.input_path, on_bad_lines='skip', engine='python')

            print(f"   Found {len(self.df)} raw rows.")
        except FileNotFoundError:
            print("❌ Error: File not found. Check your path.")
            exit()
        except pd.errors.ParserError:
            print("❌ Critical Error: The file format is too corrupted for basic parsing.")
            exit()

    def clean_data(self):
        print("Cleaning data...")

        # 1. Remove completely empty rows and duplicates
        self.df.dropna(how='all', inplace=True)
        self.df.drop_duplicates(inplace=True)

        # 2. Standardize Text (Title Case & Trim Spaces)
        text_cols = ['Sales Rep', 'Region', 'Product', 'Status']
        for col in text_cols:
            if col in self.df.columns:
                self.df[col] = self.df[col].str.strip().str.title()
                self.df[col].fillna("Unknown", inplace=True)

        # 3. Fix "Region" Inconsistencies (e.g., 'North' vs 'north')
        # (The title() above fixed capitalization, now we fill missing)
        self.df['Region'].fillna("TBD", inplace=True)

        # 4. Clean Currency (The hardest part)
        # Remove '$', ',', 'USD' and convert to float
        if 'Total Sales' in self.df.columns:
            self.df['Total Sales'] = (
                self.df['Total Sales']
                .astype(str)
                .str.replace(r'[$, USD]', '', regex=True)
            )
            self.df['Total Sales'] = pd.to_numeric(
                self.df['Total Sales'], errors='coerce').fillna(0)

        # 5. Standardize Dates to YYYY-MM-DD
        if 'Date' in self.df.columns:
            self.df['Date'] = pd.to_datetime(
                self.df['Date'], errors='coerce').dt.date

        print(f"   Rows remaining after cleaning: {len(self.df)}")

    def generate_summary(self):
        print("Generating summary pivot tables...")

        # Summary: Total Sales by Region
        summary_df = self.df.groupby(
            'Region')['Total Sales'].sum().reset_index()
        summary_df = summary_df.sort_values(by='Total Sales', ascending=False)
        return summary_df

    def save_formatted_excel(self):
        """Saves data with professional Excel formatting."""
        print(f"💾 Saving report to {self.output_path}...")

        # We use ExcelWriter to create multiple sheets
        with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:

            # --- SHEET 1: CLEAN DATA ---
            self.df.to_excel(writer, sheet_name='Clean Data', index=False)

            # --- SHEET 2: MANAGEMENT SUMMARY ---
            summary_df = self.generate_summary()
            summary_df.to_excel(
                writer, sheet_name='Summary Report', index=False)

            # --- APPLYING FORMATTING (The "Wow" Factor) ---
            workbook = writer.book
            worksheet_data = writer.sheets['Clean Data']
            worksheet_summary = writer.sheets['Summary Report']

            # Define Formats
            header_fmt = workbook.add_format(
                {'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
            currency_fmt = workbook.add_format({'num_format': '$#,##0.00'})

            # Apply Header Format to Clean Data
            for col_num, value in enumerate(self.df.columns.values):
                worksheet_data.write(0, col_num, value, header_fmt)
                # Auto-adjust column width
                worksheet_data.set_column(col_num, col_num, 20)

            # Apply Currency Format to 'Total Sales' column (Column F is index 5)
            # finding the index of 'Total Sales'
            sales_col_idx = self.df.columns.get_loc("Total Sales")
            worksheet_data.set_column(
                sales_col_idx, sales_col_idx, 15, currency_fmt)

            # Apply Header Format to Summary Sheet
            for col_num, value in enumerate(summary_df.columns.values):
                worksheet_summary.write(0, col_num, value, header_fmt)
                worksheet_summary.set_column(col_num, col_num, 20)

            # Apply Currency to Summary Sales column (Index 1)
            worksheet_summary.set_column(1, 1, 20, currency_fmt)

        print("✅ Automation Complete. Check output folder.")


# --- RUNNING THE SYSTEM ---
if __name__ == "__main__":
    # Get current directory
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Define paths
    input_csv = os.path.join(base_dir, 'input_files', 'raw_sales_data.csv')
    output_excel = os.path.join(
        base_dir, 'output_reports', 'Executive_Sales_Report.xlsx')

    # Initialize and Run
    bot = SalesReportGenerator(input_csv, output_excel)
    bot.load_data()
    bot.clean_data()
    bot.save_formatted_excel()
