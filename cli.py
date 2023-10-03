import pandas as pd
import numpy as np
import logging
import traceback

# Configure logging
logging.basicConfig(filename='log.csv', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Weighted random selection without replacement
def weighted_random_selection(obj, weights, n):
    weights = weights / weights.sum()
    return np.random.choice(obj, n, p=weights, replace=False)

def open_file(file_path):
    try:
        if file_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
            file_type = 'excel'
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            file_type = 'csv'
        else:
            raise ValueError("Unsupported file format")
        
        df = df.sort_values(by=df.columns[1], ascending=False)
        logging.info(f"Opened file: {file_path}")
        return df, file_type
    except Exception as e:
        logging.error(f"Failed to open file: {str(e)}")
        raise ValueError(f"Failed to open file: {str(e)}")

def generate_tickets(df, num_tickets, num_entered):
    try:
        if num_tickets <= 0:
            raise ValueError("Number of tickets must be positive")

        if num_tickets > len(df):
            raise ValueError(f"Number of tickets must be less than or equal to the number of records")
        
        logging.info(f"Generating winners, num given: {num_tickets}")
        
        weights = df['Chances'] / df['Chances'].sum()
        if num_entered:
            # Randomly select winners' indices based on weights
            print("num_tickets: ", num_tickets)
            winners_indices = np.random.choice(df.index, num_tickets, p=weights, replace=False)
            # Mark winners in the 'Winner' column
            df['Winner'] = False
            df.loc[winners_indices, 'Winner'] = True
            df = df.sort_values(by=['Winner', 'Chances'], ascending=False)
            df['Seat'] = range(1, len(df) + 1)
        else:
            # Randomly select CustomerNumbers as winners without actual indices
            winners_indices = np.random.choice(df['CustomerNumber'], num_tickets, p=weights, replace=False)
            # Mark winners as question marks in the 'Winner' column
            df['Winner'] = "?"
            df = df.sort_values(by=['Winner', 'Chances'], ascending=False)
            for i in range(len(df)):
                df.loc[df['CustomerNumber'] == winners_indices[i], 'Seat'] = i + 1
    
        num_winners = df['Winner'].value_counts()
        logging.info(f"Number of winners: {num_winners.to_string(header=False)}")

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
        elif file_type == 'csv':
            df.to_csv(file_name, index=False)
        else:
            raise ValueError("Unsupported file format")
        
        logging.info(f"Saved to {file_type.upper()} file: {file_name}")
    except Exception as e:
        logging.error(f"Failed to save file: {str(e)}")
        raise ValueError(f"Failed to save file: {str(e)}")

if __name__ == "__main__":
    try:
        file_path = input("Enter the path of the Excel or CSV file: ")
        df, file_type = open_file(file_path)

        num_tickets_input = input("Enter the number of tickets to generate (or press Enter to use all): ").strip()
        num_tickets = len(df) 
        num_given = False
        if not num_tickets_input == "":
            num_tickets = int(num_tickets_input)
            num_given = True
        
        df = generate_tickets(df, num_tickets, num_given)
        
        # output_file = input("Enter the output file name: ")
        save_to_file(df, file_path, file_type)
        print("Done!")
    except Exception as e:
        print(f"Error: {str(e)}")
