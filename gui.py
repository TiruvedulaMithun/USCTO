import tkinter as tk
from tkinter import filedialog, messagebox  # Import necessary modules for GUI
import pandas as pd  # Import pandas library for data handling
from pandastable import Table, TableModel  # Import PandasTable for displaying data in Tkinter
import numpy as np  # Import NumPy for numerical operations
import logging  # Import logging for tracking events and errors

# Configure logging to save logs in 'log.csv' file with timestamp, log level, and messages
logging.basicConfig(filename='log.csv', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Create a class for the Excel Viewer Application
class ExcelViewerApp:
    def __init__(self, root):
        # Initialize the main application window and set initial values
        self.root = root
        self.root.title("USCTO Seat Picker")  # Set window title
        self.file_open = False  # Flag to check if a file is opened
        self.df = None  # Variable to store DataFrame (Excel/CSV data)
        self.file_name = None  # Variable to store the name of the opened file
        self.file_type = None  # Variable to store the type of the opened file (Excel or CSV)

        # Initialize GUI widgets (textboxes, buttons, frames, menus)
        self.setup_widgets()

    def setup_widgets(self):
        # Create a text entry widget for entering the number of tickets
        self.textbox = tk.Entry(self.root, width=50)
        self.textbox.grid(row=0, column=0, padx=5, pady=5)
        self.textbox.insert(0, 'Enter the number of tickets (or leave empty to use all)')

        # Create a button to generate random tickets
        self.button = tk.Button(self.root, text="Pick Random", command=self.generate_tickets)
        self.button.grid(row=0, column=1, padx=5, pady=5)

        # Create a separator line
        self.line = tk.Frame(self.root, height=1, width=400, bg="grey80", relief='groove')
        self.line.grid(row=1, columnspan=2, sticky="ew")

        # Create a frame for displaying the table
        self.frame = tk.Frame(self.root)
        self.frame.grid(row=2, columnspan=2, sticky="nsew")

        # Create a menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Create a File menu with options (Open File, Exit)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open File", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        # Initialize PandasTable to display data
        self.table = Table(self.frame, showtoolbar=True, showstatusbar=True)
        self.table.show()

    def open_file(self):
        try:
            # Open a file dialog to select Excel or CSV file
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")])
            if file_path:
                # Load data from the selected file and display it in the table
                self.load_data(file_path)
        except Exception as e:
            # Handle and display error if loading data fails
            self.show_error(f"Failed to open file: {str(e)}")
            logging.error(f"Failed to open file: {str(e)}")

    def load_data(self, file_path):
        try:
            # Determine file type and read data accordingly (Excel or CSV)
            if file_path.endswith(('.xls', '.xlsx')):
                self.file_type = 'excel'
                self.df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                self.file_type = 'csv'
                self.df = pd.read_csv(file_path)
            else:
                raise ValueError("Unsupported file format")
            
            # Sort data by the second column (index 1) in descending order
            self.df = self.df.sort_values(by=self.df.columns[1], ascending=False)
            # Update the PandasTable with the loaded data
            self.table.updateModel(TableModel(self.df))
            # Set the file_open flag to indicate that a file is successfully opened
            self.file_open = True
            self.file_name = file_path  # Store the opened file name
            logging.info(f"Opened file: {file_path}")
        except Exception as e:
            # Handle and display error if loading data fails
            self.show_error(f"Failed to open file: {str(e)}")
            logging.error(f"Failed to open file: {str(e)}")

    def generate_tickets(self):
        try:
            if not self.file_open:
                # Display an error if no file is opened when generating tickets
                self.show_error("Please open a file first")
                return
            
            num_entered, num_tickets = self.get_num_tickets()
            winners = self.get_random_winners(num_tickets, num_entered)
            self.assign_seats(winners, num_entered)
            self.save_file()  # Save the generated data to a file

            # Log information about the generated winners and number of tickets
            logging.info(f"Generated winners, num given: {num_tickets}")
            num_winners = self.df['Winner'].value_counts()
            logging.info(f"Number of winners: {num_winners.to_string(header=False)}")
        except Exception as e:
            # Handle and display error if generating tickets fails
            self.show_error(f"Failed to generate tickets: {str(e)}")
            logging.error(f"Failed to generate tickets: {str(e)}")

    def get_num_tickets(self):
        num_tickets = self.textbox.get().strip()  # Get user input from the textbox
        if num_tickets == "Enter the number of tickets (or leave empty to use all)" or num_tickets == "":
            return False, len(self.df)  # Return False (flag) and total number of records if no input is given
        else:
            if(int(num_tickets) < 0):
                raise ValueError("Number of tickets must be positive")  # Raise error for negative input
            if(int(num_tickets) > len(self.df)):
                raise ValueError(f"Number of tickets must be less than {len(self.df)}")  # Raise error if input exceeds total records
            return True, int(num_tickets)  # Return True (flag) and entered number of tickets

    def get_random_winners(self, num_tickets, num_entered):
        # Calculate weights based on 'Chances' column to pick random winners
        weights = self.df['Chances'] / self.df['Chances'].sum()
        if num_entered:
            # Randomly select winners' indices based on weights
            winners_indices = np.random.choice(self.df.index, num_tickets, p=weights, replace=False)
            # Mark winners in the 'Winner' column
            self.df['Winner'] = False
            self.df.loc[winners_indices, 'Winner'] = True
        else:
            # Randomly select CustomerNumbers as winners without actual indices
            winners_indices = np.random.choice(self.df['CustomerNumber'], num_tickets, p=weights, replace=False)
            # Mark winners as question marks in the 'Winner' column
            self.df['Winner'] = "?"
        return winners_indices

    def assign_seats(self, winners_indices, num_entered):
        # Sort DataFrame by 'Winner' and 'Chances' columns in descending order
        self.df = self.df.sort_values(by=['Winner', 'Chances'], ascending=False)
        if num_entered:
            # Assign seat numbers based on the sorted DataFrame
            self.df['Seat'] = range(1, len(self.df) + 1)
            actual_indices = winners_indices  # Use winners' indices as actual indices
        else:
            # Assign seat numbers based on winners' indices and update the PandasTable
            for i in range(len(self.df)):
                self.df.loc[self.df['CustomerNumber'] == winners_indices[i], 'Seat'] = i + 1
            actual_indices = self.df['CustomerNumber'].isin(winners_indices)
        self.table.updateModel(TableModel(self.df.loc[actual_indices]))  # Update PandasTable with actual indices
        self.table.redraw()  # Redraw the table to reflect the changes

    def save_file(self):
        # Save the DataFrame to a file based on the file type (Excel or CSV)
        if self.file_type == 'excel':
            with pd.ExcelWriter(self.file_name) as writer:
                self.df.to_excel(writer, index=False)  # Save DataFrame to Excel file without index
                logging.info(f"Saved to Excel file: {self.file_name}")
        elif self.file_type == 'csv':
            self.df.to_csv(self.file_name, index=False)  # Save DataFrame to CSV file without index
            logging.info(f"Saved to CSV file: {self.file_name}")
        else:
            raise ValueError("Unsupported file format")  # Raise error for unsupported file type

    def show_error(self, message):
        # Display an error message dialog with the given message
        messagebox.showerror("Error", message)

# Entry point of the program
if __name__ == "__main__":
    root = tk.Tk()  # Create the main Tkinter window
    app = ExcelViewerApp(root)  # Initialize the Excel Viewer Application
    root.mainloop()  # Start the Tkinter main loop to run the application
