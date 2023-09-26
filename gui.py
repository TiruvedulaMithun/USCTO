import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pandastable import Table, TableModel
import random
import traceback
import numpy as np
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

class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Viewer")
        self.file_open = False
        self.df = None
        self.file_name = None
        self.file_type = None
        # Initialize text box number field. left-aligned. a button right-aligned in the same line
        self.textbox = tk.Entry(root, width=50)
        
        self.textbox.grid(row=0, column=0, padx=5, pady=5)
        self.textbox.focus_set()
        # add text placeholder in textbox
        self.textbox.insert(0, 'Enter the number of tickets (or leave empty to use all)')

        # Initialize button to generate tickets in the same line as textbox
        self.button = tk.Button(root, text="Pick Random", command=self.generate_tickets, justify=tk.RIGHT)
        # Add button to grid
        self.button.grid(row=0, column=1, padx=5, pady=5)

        # Add a separator line
        self.line = tk.Frame(root, height=1, width=400, bg="grey80", relief='groove')
        self.line.grid(row=1, columnspan=2, sticky="ew")
        
        # Create a frame to hold the table in the next line
        self.frame = tk.Frame(root)
        self.frame.grid(row=2, columnspan=2, sticky="nsew")
        
        # Create a menu bar
        menubar = tk.Menu(root)
        root.config(menu=menubar)
        
        # Create a "File" menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open File", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=root.quit)

        # Initialize table
        self.table = Table(self.frame, showtoolbar=True, showstatusbar=True)
        self.table.show()
        
    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")])
        # Grey out root and show loading
        self.root.config(cursor="wait")
        self.root.update()
        if file_path:
            try:
                if file_path.endswith(('.xls', '.xlsx')):
                    df = pd.read_excel(file_path)
                    self.file_type = 'excel'
                elif file_path.endswith('.csv'):
                    self.file_type = 'csv'
                    df = pd.read_csv(file_path)
                else:
                    raise ValueError("Unsupported file format")
                self.file_name = file_path
                df = df.sort_values(by=df.columns[1], ascending=False)
                self.file_open = True
                # Update the table with the data from the file
                self.df = df
                logging.info(f"Opened file: {file_path}")
            except Exception as e:
                tk.messagebox.showerror("Error", f"Failed to open file: {str(e)}")
                logging.error(f"Failed to open file: {str(e)}")
        else:
            tk.messagebox.showerror("Error", f"Failed to open file")
            logging.error(f"Failed to open file")
        # Un-grey out root
        self.root.config(cursor="")
        self.root.update()

    def generate_tickets(self):
        if(not self.file_open):
            tk.messagebox.showerror("Error", f"Please open a file first")
            return
        try:
            num_given = False
            num_tickets = self.textbox.get().strip()
            dflen = len(self.df)
            if(num_tickets == "Enter number of tickets here" or num_tickets == ""):
                num_tickets = dflen
                # winners is a column of question marks string
                winners = pd.Series(["?" for i in range(dflen)])
                # assign seat numbers to all rows. higher the chances, higher the chance of getting a lower seat number
                seat = np.random.choice(range(1, len(self.df) + 1), size=len(self.df), replace=False)
                
                self.df['Seat'] = seat
                self.df['Winner'] = winners
                self.df = self.df.sort_values(by=['Seat'])
                self.table.updateModel(TableModel(self.df))
                logging.info(f"Generated Seats no num given: {dflen}")
                
            else:
                num_given = True
                num_tickets = int(self.textbox.get())
                if(num_tickets < 0):
                    raise ValueError("Number of tickets must be positive")
                if(num_tickets > dflen):
                    raise ValueError(f"Number of tickets must be less than {dflen}")
                winners = weighted_random_selection(self.df.index, self.df['Chances'] * 100, num_tickets)
                self.df['Winner'] = False
            
                self.df.loc[winners, 'Winner'] = True
                # Sort df by winner and chances
                self.df = self.df.sort_values(by=['Winner', 'Chances'], ascending=False)
                # assign seat numbers to all rows
                self.df['Seat'] = range(1, len(self.df) + 1)
                self.table.updateModel(TableModel(self.df.loc[winners]))
                
                logging.info(f"Generated winners, num given: {num_tickets}")
            
            self.table.redraw()
           
            # Display num of winners
            num_winners = self.df['Winner'].value_counts()
            logging.info(f"Number of winners: {num_winners}")
            
            if(self.file_type == 'excel'):
                with pd.ExcelWriter(self.file_name) as writer:
                    self.df.to_excel(writer, index=False)
                    logging.info(f"Saved to Excel file: {self.file_name}")
            elif(self.file_type == 'csv'):
                self.df.to_csv(self.file_name, index=False)
                logging.info(f"Saved to CSV file: {self.file_name}")
            else:
                raise ValueError("Unsupported file format")

        except Exception as e:
            traceback.print_exc()
            tk.messagebox.showerror("Error", f"Failed to generate tickets: {str(e)}")
            logging.error(f"Failed to generate tickets: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
