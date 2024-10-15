import subprocess
import pandas as pd

# Step 1: Run the 'tradebook_builder.py' script to generate tradebook_df
subprocess.run(['python', 'tradebook_builder.py'])

# Step 2: Import the tradebook_df DataFrame from 'tradebook_builder.py'
from tradebook_builder import tradebook_df

# Step 3: Initialize dictionaries for active trades and realized PnL
active_trades = {}  # Tracks active positions for each symbol
realized_pnl_data = []  # Tracks realized PnL as rows with detailed info

# Step 4: Iterate through each row in tradebook_df
for _, row in tradebook_df.iterrows():
    symbol = row['Symbol']
    trade_type = row['Trade Type'].lower()  # 'buy' or 'sell'
    quantity = row['Quantity']
    price = row['Price']
    expiry_date = pd.to_datetime(row['Expiry Date'])  # Ensure expiry date is in datetime format
    execution_time = pd.to_datetime(row['Formatted Order Execution Time'])  # Ensure execution time is in datetime format
    underlying = row['Underlying']
    strike_price = row['Strike']
    option_type = row['Option Type']
    product = 'Option' if option_type in ['PE', 'CE'] else 'Futures'  # Determining if it's an Option or Futures
    trade_cost = price * quantity  # Calculate the cost of the trade

    # Step 4.1: Check for expired positions
    expired_symbols = []
    for active_symbol, data in active_trades.items():
        if execution_time > data['expiry_date']:
            # Expired position found; we close it out at zero price
            if data['quantity'] < 0:  # Short position
                pnl = data['avg_price'] * abs(data['quantity'])  # Profit from short (since it's bought back at zero)
            else:  # Long position
                pnl = -data['avg_price'] * abs(data['quantity'])  # Loss from long (since it's sold at zero)
            
            # Add the PnL entry with detailed information
            realized_pnl_data.append({
                'Formatted Order Execution Time': execution_time.strftime('%Y-%m-%d %H:%M:%S'),
                'Symbol': active_symbol,
                'Underlying': data['underlying'],
                'Expiry Date': data['expiry_date'].strftime('%Y-%m-%d'),
                'Strike Price': data['strike'],
                'Option Type': data['option_type'],
                'Product': data['product'],
                'Realized PnL': pnl
            })
            expired_symbols.append(active_symbol)  # Mark the symbol for removal

    # Remove expired symbols from active_trades
    for expired_symbol in expired_symbols:
        del active_trades[expired_symbol]

    # Step 4.2: Process the current row for buy/sell logic
    # Check if the symbol exists in active_trades, if not, initialize it
    if symbol not in active_trades:
        active_trades[symbol] = {'quantity': 0, 'cost': 0, 'avg_price': 0, 'expiry_date': expiry_date,
                                 'underlying': underlying, 'strike': strike_price, 'option_type': option_type,
                                 'product': product}

    # ** Buy Logic **
    if trade_type == 'buy':
        # Case 1: Covering a short position (quantity < 0)
        if active_trades[symbol]['quantity'] < 0:
            # PnL for covering the short position
            cover_quantity = min(abs(active_trades[symbol]['quantity']), quantity)  # Only cover as much as we're short
            pnl = cover_quantity * (active_trades[symbol]['avg_price'] - price)  # Profit from covering the short

            # Add the PnL entry with detailed information
            realized_pnl_data.append({
                'Formatted Order Execution Time': execution_time.strftime('%Y-%m-%d %H:%M:%S'),
                'Symbol': symbol,
                'Underlying': underlying,
                'Expiry Date': expiry_date.strftime('%Y-%m-%d'),
                'Strike Price': strike_price,
                'Option Type': option_type,
                'Product': product,
                'Realized PnL': pnl
            })

            # Update the active position
            active_trades[symbol]['quantity'] += cover_quantity  # Cover the short position
            active_trades[symbol]['cost'] -= cover_quantity * active_trades[symbol]['avg_price']  # Reduce the cost of short

            # Any remaining quantity will increase the long position
            remaining_quantity = quantity - cover_quantity
            if remaining_quantity > 0:
                active_trades[symbol]['quantity'] += remaining_quantity
                active_trades[symbol]['cost'] += remaining_quantity * price

        # Case 2: Increasing an existing long position (quantity > 0 or 0)
        else:
            # No PnL for increasing a long position, just update the cost and quantity
            active_trades[symbol]['quantity'] += quantity
            active_trades[symbol]['cost'] += trade_cost

        # Recalculate the average price after the buy (for both covering short and increasing long)
        if active_trades[symbol]['quantity'] != 0:  # Prevent division by zero
            active_trades[symbol]['avg_price'] = abs(active_trades[symbol]['cost'] / active_trades[symbol]['quantity'])

    # ** Sell Logic **
    elif trade_type == 'sell':
        # Case 1: Reducing a long position (quantity > 0)
        if active_trades[symbol]['quantity'] > 0:
            # Sell a portion or the full long position
            sell_quantity = min(active_trades[symbol]['quantity'], quantity)  # Only sell as much as we hold
            pnl = sell_quantity * (price - active_trades[symbol]['avg_price'])  # Profit from selling the long position

            # Add the PnL entry with detailed information
            realized_pnl_data.append({
                'Formatted Order Execution Time': execution_time.strftime('%Y-%m-%d %H:%M:%S'),
                'Symbol': symbol,
                'Underlying': underlying,
                'Expiry Date': expiry_date.strftime('%Y-%m-%d'),
                'Strike Price': strike_price,
                'Option Type': option_type,
                'Product': product,
                'Realized PnL': pnl
            })

            # Update the active position
            active_trades[symbol]['quantity'] -= sell_quantity
            active_trades[symbol]['cost'] -= sell_quantity * active_trades[symbol]['avg_price']

            # Any remaining quantity will increase the short position
            remaining_quantity = quantity - sell_quantity
            if remaining_quantity > 0:
                active_trades[symbol]['quantity'] -= remaining_quantity
                active_trades[symbol]['cost'] += remaining_quantity * price

        # Case 2: Increasing a short position (quantity < 0 or 0)
        else:
            # No PnL for increasing a short position, just update the cost and quantity
            active_trades[symbol]['quantity'] -= quantity
            active_trades[symbol]['cost'] += trade_cost  # Add to the cost for the short position

        # Recalculate the average price after the sell (for both reducing long and increasing short)
        if active_trades[symbol]['quantity'] != 0:  # Prevent division by zero
            active_trades[symbol]['avg_price'] = abs(active_trades[symbol]['cost'] / active_trades[symbol]['quantity'])

    # If the quantity for the symbol becomes zero, assume the trade is closed
    if active_trades[symbol]['quantity'] == 0:
        del active_trades[symbol]  # Remove the symbol from active trades if fully sold/covered

# Step 5: Convert realized PnL data to a DataFrame
realized_pnl_df = pd.DataFrame(realized_pnl_data)

# Step 6: Add a 'Cumulative PnL' column
realized_pnl_df['Cumulative PnL'] = realized_pnl_df['Realized PnL'].cumsum()

# Step 7: Export to Excel
output_file_path = "C:\\Users\\prana\\OneDrive\\Desktop\\Documents\\python\\data analysis project - trading journal\\active_trade_and_pnl.xlsx"
realized_pnl_df.to_excel(output_file_path, index=False)

print("Data exported to 'active_trade_and_pnl.xlsx'")
