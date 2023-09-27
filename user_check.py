# Import required modules
import subprocess
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path


class UserCheckLogic():
    def __init__(self, excel_file):
        """
        Initialization of UserCheckLogic class

        Args:
            excel_file (str): Path to Excel file used for user status check.
        """
        # Check if the Excel file exists before opening it
        if not Path(excel_file).is_file():
            raise FileNotFoundError(f"Datoteka '{excel_file}' ne postoji.")
        self.excel_file = excel_file

    def current_time(self):
        """
        Returns current date to be used as part of sheet name.

        Returns:
            datetime.date: Current date.
        """
        current_date = datetime.now().date()
        return current_date

    def check_user_status(self):
        """
        Perform the check of user status based on Excel file and updates it with results.
        """
        # Open the Excel file
        self.workbook = load_workbook(self.excel_file)
        self.sheet = self.workbook.active
        self.sheet.title = f"Provjera_{self.current_time()}"

        # Find the indexes of column "Username" and "Account Status"
        username_col_pos = 2  # Column B
        account_status_col_pos = 5  # Column E

        # Set the initial row to second row so we would ignore column titles
        start_row = 2

        # Iterating through rows via "Username" column
        for row in self.sheet.iter_rows(min_row=start_row, min_col=username_col_pos,
                                        values_only=True):
            username_column = row[0]  # Iterate through "Username" column
            if username_column is not None: # Check if username values are not emtpy
                # Split username to remove domain if present
                username_parts = username_column.split("/")
                username = username_parts[-1]  # Take the last part as username
                
                # Commit net user command
                command = f"net user {username}"  # Add " /domain" or applicable at the end
                result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, text=True)

                # Analyse if the user is active or not or even if the user exist or not       
                if "Account active               No" in result.stdout:
                    account_status = "Disabled"
                elif "Account active               Yes" in result.stdout:
                    account_status = "Active"
                else:
                    account_status = "User not found"

                # Update the Excel file with gathered information
                self.sheet.cell(row=start_row, column=account_status_col_pos,
                                value=account_status)
            else:
                self.sheet.cell(start_row, column=account_status_col_pos,
                                value="No username provided")
            start_row += 1  # Increase the row number for next entry

    def save_to_file(self):
        """Save to Excel file"""
        self.workbook.save(self.excel_file)