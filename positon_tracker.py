import subprocess
import pandas as pd

# Step 1: Run the 'tradebook_builder.py' script to generate tradebook_df
subprocess.run(['python', 'tradebook_builder.py'])

# Step 2: Import the tradebook_df DataFrame from 'tradebook_builder.py'
from tradebook_builder import tradebook_df



# Function to build the active position tracker with cost and average price tracking
def build_active_position_tracker(tradebook_df):
    # Sort the DataFrame chronologically by 'Formatted Order Execution Time'
    tradebook_df = tradebook_df.sort_values(by='Formatted Order Execution Time')

    # Dictionary to keep track of active trades
    active_trades = {}

    # List to store snapshots for each order execution time
    trade_snapshots = []

    # Loop through each row in the tradebook_df
    for _, row in tradebook_df.iterrows():
        symbol = row['Symbol']
        trade_type = row['Trade Type'].lower()  # 'buy' or 'sell'
        quantity = row['Quantity']
        price = row['Price']  # Get the price for the trade
        time = pd.to_datetime(row['Formatted Order Execution Time'])  # Ensure it's a datetime object
        expiry_date = pd.to_datetime(row['Expiry Date'])  # Ensure expiry is a datetime object

        # Calculate cost for the trade
        cost = price * quantity

        # Check and remove expired symbols at the start of each iteration
        expired_symbols = [sym for sym, data in active_trades.items() if time > data['expiry_date']]
        for expired_symbol in expired_symbols:
            del active_trades[expired_symbol]  # Remove expired positions

        # If the symbol has expired, skip processing
        if time > expiry_date:
            continue

        # If the symbol is not in active trades, initialize it with 0 quantity, cost, and avg price
        if symbol not in active_trades:
            active_trades[symbol] = {'quantity': 0, 'cost': 0, 'avg_price': 0, 'expiry_date': expiry_date}

        # Update the quantity, cost, and avg price based on trade type
        if trade_type == 'buy':
            active_trades[symbol]['quantity'] += quantity  # Increase quantity for 'buy'
            active_trades[symbol]['cost'] += cost  # Add to the cost
        elif trade_type == 'sell':
            active_trades[symbol]['quantity'] -= quantity  # Decrease quantity for 'sell'
            active_trades[symbol]['cost'] -= cost  # Deduct the cost (assuming sell offsets the cost)

        # Remove the symbol if the quantity becomes zero
        if active_trades[symbol]['quantity'] == 0:
            del active_trades[symbol]
        else:
            # Calculate the average price as Cost / Quantity if Quantity is not zero
            active_trades[symbol]['avg_price'] = active_trades[symbol]['cost'] / active_trades[symbol]['quantity']

        # Take a snapshot of the current state of active trades (non-zero quantities only)
        symbols_at_time = []
        quantities_at_time = []
        costs_at_time = []
        avg_prices_at_time = []

        for sym, data in active_trades.items():
            if data['quantity'] != 0:  # Only keep symbols with non-zero quantities
                symbols_at_time.append(sym)
                quantities_at_time.append(data['quantity'])
                costs_at_time.append(data['cost'])
                avg_prices_at_time.append(data['avg_price'])

        # Store the snapshot for the current order execution time
        trade_snapshots.append({
            'Formatted Order Execution Time': time.strftime('%Y-%m-%d %H:%M:%S'),
            'Symbols': ', '.join(symbols_at_time),  # Comma-separated list of symbols
            'Quantities': ', '.join(map(str, quantities_at_time)),  # Corresponding quantities
            'Costs': ', '.join(map(str, costs_at_time)),  # Corresponding cumulative costs
            'Avg Prices': ', '.join(map(lambda x: f"{x:.2f}", avg_prices_at_time))  # Corresponding average prices
        })

    # Convert the snapshots into a DataFrame
    trade_snapshots_df = pd.DataFrame(trade_snapshots)

    # Drop duplicates and keep only the last entry for each time
    trade_snapshots_df = trade_snapshots_df.drop_duplicates(subset=['Formatted Order Execution Time'], keep='last')

    return trade_snapshots_df

# Step 3: Use the imported tradebook_df to build the active position tracker
trade_snapshots_df = build_active_position_tracker(tradebook_df)

# Step 4: Export active trades snapshot DataFrame to Excel
output_file_path = "C:\\Users\\prana\\OneDrive\\Desktop\\Documents\\python\\data analysis project - trading journal\\position_tracker_snapshot_with_avg_price.xlsx"
trade_snapshots_df.to_excel(output_file_path, index=False)

# Final message
print("Position tracker snapshot with avg price is ready and saved as position_tracker_snapshot_with_avg_price.xlsx")