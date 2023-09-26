import pandas as pd
import numpy as np
import random
import traceback
import logging
import csv

# Configure logging
logging.basicConfig(filename='log.csv', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Weighted random selection without replacement
def weighted_random_selection(obj, weights, n):
    # normalize weights
    weights = weights / weights.sum()
    # randomly choose n items
    return np.random.choice(obj, n, p=weights, replace=False)

def open_file(file_path):
    try:
        if file_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
            file_type = 'excel'
        elif file_path.endswith('.csv'):
            file_type = 'csv'
            df = pd.read_csv(file_path)
        else:
            raise ValueError("Unsupported file format")

        df = df.sort_values(by=df.columns[1], ascending=False)
        logging.info(f"Opened file: {file_path}")
        return df, file_type
    except Exception as e:
        logging.error(f"Failed to open file: {str(e)}")
        raise ValueError(f"Failed to open file: {str(e)}")

def generate_tickets(df, num_tickets):
    try:
        dflen = len(df)

        if num_tickets <= 0:
            raise ValueError("Number of tickets must be positive")

        if num_tickets > dflen:
            raise ValueError(f"Number of tickets must be less than {dflen}")

        num_given = num_tickets != dflen

        if not num_given:
            normalizedChances = df["Chances"] / df["Chances"].sum()
            num_tickets = dflen
            # winners is a column of question marks string
            winners = pd.Series(["?" for i in range(dflen)])
            # assign seat numbers to all rows. higher the chances, higher the chance of getting a lower seat number
            seat = np.random.choice(range(1, len(df) + 1), size=len(df), replace=False)

            df['Seat'] = seat
            df['Winner'] = winners
            df = df.sort_values(by=['Seat'])
            logging.info(f"Generated Seats no num given: {dflen}")
        else:
            winners = weighted_random_selection(df.index, df['Chances'] * 100, num_tickets)
            df['Winner'] = False

            df.loc[winners, 'Winner'] = True
            # Sort df by winner and chances
            df = df.sort_values(by=['Winner', 'Chances'], ascending=False)
            # assign seat numbers to all rows
            df['Seat'] = range(1, len(df) + 1)

            logging.info(f"Generated winners, num given: {num_tickets}")

        # Display num of winners
        num_winners = df['Winner'].value_counts()
        logging.info(f"Number of winners: {num_winners}")

        return df

    except Exception as e:
        traceback.print_exc()
        logging.error(f"Failed to generate tickets: {str(e)}")
        raise ValueError(f"Failed to generate tickets: {str(e)}")

def save_to_file(df, file_name, file_type):
    try:
        if file_type == 'excel':
            with pd.ExcelWriter(file_name) as writer:
                df.to_excel(writer, index=False)
            logging.info(f"Saved to Excel file: {file_name}")
        elif file_type == 'csv':
            df.to_csv(file_name, index=False)
            logging.info(f"Saved to CSV file: {file_name}")
        else:
            raise ValueError("Unsupported file format")
    except Exception as e:
        logging.error(f"Failed to save file: {str(e)}")
        raise ValueError(f"Failed to save file: {str(e)}")

if __name__ == "__main__":
    file_path = input("Enter the path of the Excel or CSV file: ")
    df, file_type = open_file(file_path)
    
    num_tickets_input = input("Enter the number of tickets to generate (or press Enter to use all): ").strip()
    if num_tickets_input == "":
        num_tickets = len(df)
    else:
        num_tickets = int(num_tickets_input)
    
    df = generate_tickets(df, num_tickets)
    
    output_file = input("Enter the output file name: ")
    save_to_file(df, output_file, file_type)
