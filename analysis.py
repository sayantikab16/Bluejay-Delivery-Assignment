import pandas as pd
from datetime import datetime, timedelta

# Function to load the Excel file and perform analysis
def analyze_excel_file(file_path, consecutive_days_threshold=7):
    try:
        # Read the Excel sheet into a DataFrame
        df = pd.read_excel(file_path)

        # Trim column names to remove leading/trailing spaces
        df.columns = df.columns.str.strip()

        # Initialize sets to keep track of printed employees
        consecutive_printed = set()
        short_break_printed = set()
        long_shift_printed = set()

        
        print("part 1: ")

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in consecutive_printed:
                continue

            # Check for consecutive days worked
            if index > 0 and employee_name == df.at[index - 1, 'Employee Name']:
                consecutive_days = 1
                for i in range(index - 1, -1, -1):
                    if df.at[i, 'Employee Name'] == employee_name:
                        consecutive_days += 1
                    else:
                        break
                if consecutive_days >= consecutive_days_threshold:
                    print(f"Employee: {employee_name}, Position: {position_id}")
                    consecutive_printed.add(employee_name)

   
        print("part 2: ")

        employee_breaks = {}  # Dictionary to track breaks for each employee

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in short_break_printed:
                continue

            if employee_name in employee_breaks:
                last_time_out = employee_breaks[employee_name]
                time_in = row['Time']

                if isinstance(time_in, str) and isinstance(last_time_out, str):
                    time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                    last_time_out = datetime.strptime(last_time_out, '%m/%d/%Y %I:%M %p')

                    time_diff = (time_in - last_time_out).total_seconds() / 3600
                    if 1 < time_diff < 10:
                        print(f"Employee: {employee_name}, Position: {position_id}")
                        short_break_printed.add(employee_name)
                else:
                    time_in = None

            employee_breaks[employee_name] = row['Time Out']


        print("part 3: ")

        for index, row in df.iterrows():
            employee_name = row['Employee Name']
            position_id = row['Position ID']

            if employee_name in long_shift_printed:
                continue

            # Check for shifts longer than 14 hours
            duration_str = row['Timecard Hours (as Time)']
            if pd.notna(duration_str):
                try:
                    hours, minutes = map(int, duration_str.split(':'))
                    duration = timedelta(hours=hours, minutes=minutes)
                except ValueError:
                    # Handle invalid duration format 
                    duration = None
            else:
                # Handle missing values
                duration = None

            if duration is not None and duration.total_seconds() / 3600 > 14:
                print(f"Employee: {employee_name}, Position: {position_id}")
                long_shift_printed.add(employee_name)

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":

    file_path = 'Assignment_Timecard.xlsx'

    # Call the analyze_excel_file function with the specified consecutive_days_threshold
    analyze_excel_file(file_path, consecutive_days_threshold=7)
