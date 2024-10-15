import pandas as pd
import re
from datetime import datetime, timedelta
import calendar

def get_last_thursday(year, month):
    """
    Get the last Thursday of a given month and year.
    """
    last_day = calendar.monthrange(year, month)[1]  # Get the last day of the month
    last_date = datetime(year, month, last_day)  # Create a datetime object for the last day
    while last_date.weekday() != 3:  # 3 represents Thursday (Monday=0, ..., Thursday=3, ..., Sunday=6)
        last_date -= timedelta(days=1)
    return last_date

def extract_details(symbol):
    pattern_type1 = re.compile(r"([A-Z]+)(\d{2})(\d{1})(\d{2})(\d+)([CEP]{2})")
    pattern_type2 = re.compile(r"([A-Z]+)(\d{2})([A-Z]{3})(\d+)([CEP]{2})")
    pattern_type3 = re.compile(r"([A-Z]+)(\d{2})([A-Z]{1})(\d{2})(\d+)([CEP]{2})")
    pattern_type4 = re.compile(r"([A-Z]+)(\d{2})([A-Z]{3})(FUT)")
    pattern_type5 = re.compile(r"([A-Z]+)(\d{2})([A-Z]{3})(\d+\.\d+)([CEP]{2})")  # New pattern

    # Matching pattern_type1
    if match := pattern_type1.match(symbol):
        underlying, year, month, day, strike, option_type = match.groups()
        year = '20' + year
        expiry_date = datetime(int(year), int(month), int(day), 15, 30)  # Set expiry time to 15:30:00
        return underlying, expiry_date, strike, option_type

    # Matching pattern_type2
    elif match := pattern_type2.match(symbol):
        underlying, year, month_str, strike, option_type = match.groups()
        year = '20' + year
        try:
            month = datetime.strptime(month_str, '%b').month
            last_day = calendar.monthrange(int(year), month)[1]
            expiry_date = datetime(int(year), month, last_day, 15, 30)  # Set expiry time to 15:30:00
            return underlying, expiry_date, strike, option_type
        except ValueError:
            print(f"Invalid month abbreviation in symbol: {symbol}")
            return None, None, None, None

    # Matching pattern_type3
    elif match := pattern_type3.match(symbol):
        underlying, year, month_char, day, strike, option_type = match.groups()
        year = '20' + year
        month_map = {'O': 10, 'N': 11, 'D': 12}
        month = month_map.get(month_char.upper(), None)
        if month:
            expiry_date = datetime(int(year), int(month), int(day), 15, 30)  # Set expiry time to 15:30:00
            return underlying, expiry_date, strike, option_type
        else:
            print(f"Invalid month character in symbol: {symbol}")
            return None, None, None, None

    # Matching pattern_type4 (FUT contracts, e.g., HDFCBANK21JUNFUT)
    elif match := pattern_type4.match(symbol):
        underlying, year, month_str, option_type = match.groups()
        year = '20' + year
        try:
            month = datetime.strptime(month_str, '%b').month
            # Calculate expiry date: last Thursday of the contract month
            expiry_date = get_last_thursday(int(year), month).replace(hour=15, minute=30)  # Expire at 15:30
            return underlying, expiry_date, None, option_type
        except ValueError:
            print(f"Invalid month abbreviation in symbol: {symbol}")
            return None, None, None, None

    # Matching pattern_type5 (e.g., ITC21JUL207.5CE)
    elif match := pattern_type5.match(symbol):
        underlying, year, month_str, strike, option_type = match.groups()
        year = '20' + year
        try:
            month = datetime.strptime(month_str, '%b').month
            last_day = calendar.monthrange(int(year), month)[1]
            expiry_date = datetime(int(year), month, last_day, 15, 30)  # Set expiry time to 15:30:00
            return underlying, expiry_date, strike, option_type
        except ValueError:
            print(f"Invalid month abbreviation in symbol: {symbol}")
            return None, None, None, None

    # Unmatched symbol
    print(f"Unmatched symbol format: {symbol}")
    return None, None, None, None

# Read the Excel file into a DataFrame
file_path = "C:\\Users\\prana\\OneDrive\\Desktop\\Documents\\python\\data analysis project - trading journal\\tradebook_input_og.xlsx"
tradebook_df = pd.read_excel(file_path)

# Convert 'Order Execution Time' to the desired format (YYYY-MM-DD HH:MM:SS)
tradebook_df['Formatted Order Execution Time'] = pd.to_datetime(tradebook_df['Order Execution Time']).dt.strftime('%Y-%m-%d %H:%M:%S')

# Drop unwanted columns
tradebook_df = tradebook_df.drop(['Segment', 'ISIN', 'Auction', 'Trade ID', 'Order Execution Time'], axis=1)

# Apply extract_details function to the 'Symbol' column
def process_symbol(symbol):
    underlying, expiry_date, strike, option_type = extract_details(symbol)
    return pd.Series([underlying, expiry_date, strike, option_type])

# Create new columns for 'Underlying', 'Expiry Date', 'Strike', and 'Option Type'
tradebook_df[['Underlying', 'Expiry Date', 'Strike', 'Option Type']] = tradebook_df['Symbol'].apply(process_symbol)

# Add the 'Product' column based on 'Option Type'
tradebook_df['Product'] = tradebook_df['Option Type'].apply(lambda x: 'Option' if x in ['PE', 'CE'] else 'Futures' if x == 'FUT' else None)

# Save the DataFrame to an Excel file
output_file_path = "C:\\Users\\prana\\OneDrive\\Desktop\\Documents\\python\\data analysis project - trading journal\\tradebook.xlsx"
tradebook_df.to_excel(output_file_path, index=False)

print(f"Data saved to {output_file_path}")

if __name__ == "__main__":
    print("Tradebook processing complete.")
