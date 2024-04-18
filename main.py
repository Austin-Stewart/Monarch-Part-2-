import tkinter as tk
from tkinter import filedialog
import pandas as pd
import openpyxl

# Function to process the text data and save to Excel
def process_and_save(source_file, destination_file):
    # Read text data from the source file
    with open(source_file, 'r') as file:
        text_data = file.read()

    # Split text into lines
    lines = text_data.strip().split('\n')

    # Initialize lists to store data
    regions = []
    region_names = []
    county_nums = []
    county_names = []
    first_names = []
    last_names = []
    current_worker_ids = []
    error_codes = []
    error_types = []
    payment_numbers = []
    line_items = []
    # Client ids
    client_ids = []
    facilities = []
    service_codes = []
    begin_dates = []
    end_dates = []
    error_dates = []
    entry_amounts = []

    # Process each line of text
    current_region = None
    current_region_name = None
    current_first_name = None
    current_last_name = None
    current_worker_id = None
    current_county_num = None
    current_county_name = None
    for line in lines:
        if "REGION LINE" in line:
            current_region = line[30:32]
            current_region_name = line[39:52]
        elif "COUNTY LINE" in line:
            current_county_num = line[30:32]
            current_county_name = line[39:66]
        elif "Assigned Worker LINE" in line:
            current_first_name = line[69:75]
            current_last_name = line[48:55]
            current_worker_id = line[39:45]  # Assign current worker ID here
        elif "Entry Line 3" in line:
            regions.append(current_region)
            region_names.append(current_region_name)
            county_nums.append(current_county_num)
            county_names.append(current_county_name)
            first_names.append(current_first_name)
            last_names.append(current_last_name)
            current_worker_ids.append(current_worker_id)  # Append to list here
            error_codes.append(line[14:18])
            error_types.append(line[20:25])
            payment_numbers.append(line[27:36])
            line_items.append(line[38:41])
            # Client ID
            client_ids.append(line[44:52])
            facilities.append(line[54:64])
            service_codes.append(line[85:90])
            begin_dates.append(line[110:121])
            end_dates.append(line[122:133])
            error_dates.append(line[134:145])
        elif "Entry Line 6" in line:
            entry_amounts.append(line[31:39])

    # Print the length of each array for debugging
    print("Length of arrays:")
    print("Regions:", len(regions))
    print("Region Names:", len(region_names))
    print("County Nums:", len(county_nums))
    print("County Names:", len(county_names))
    print("First Names:", len(first_names))
    print("Last Names:", len(last_names))
    print("IDs:", len(client_ids))
    print("Error Codes:", len(error_codes))
    print("Error Types:", len(error_types))
    print("Entry Numbers:", len(payment_numbers))
    print("Line Items:", len(line_items))
    print("Caps IDs:", len(client_ids))
    print("Facilities:", len(facilities))
    print("Service Codes:", len(service_codes))
    print("Begin Dates:", len(begin_dates))
    print("End Dates:", len(end_dates))
    print("Error Dates:", len(error_dates))
    print("Entry Amounts:", len(entry_amounts))

    # Create a DataFrame using pandas
    df = pd.DataFrame({
        'ERROR CODE': error_codes,
        'ERROR TYPE': error_types,
        'PAYMENT #': payment_numbers,
        'LINE ITEM': line_items,
        'CLIENT ID': client_ids,
        'FACILITY #': facilities,
        'SERVICE CODE': service_codes,
        'BEGIN DATE': begin_dates,
        'END DATE': end_dates,
        'ERROR DATE': error_dates,
        'WORKER ID': current_worker_ids,  # Use current_worker_ids list
        'WORKER LAST NAME': last_names,
        'WORKER FIRST NAME': first_names,
        'COUNTY #': county_nums,
        'COUNTY NAME': county_names,
        'REGION': regions,
        'SERVICE AMOUNT': entry_amounts,
    })

    # Write DataFrame to Excel
    df.to_excel(destination_file, index=False)
    status_label.config(text="Data processed and saved to Excel.")


# Function to handle source file selection
def select_source_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if file_path:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, file_path)


# Function to handle destination file selection
def select_destination_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        destination_entry.delete(0, tk.END)
        destination_entry.insert(0, file_path)


# Create the main application window
app = tk.Tk()
app.title("Text to Excel Converter")

# Create and place widgets
source_label = tk.Label(app, text="Source Text File:")
source_label.pack()

source_entry = tk.Entry(app, width=50)
source_entry.pack()

source_button = tk.Button(app, text="Select Source File", command=select_source_file)
source_button.pack()

destination_label = tk.Label(app, text="Destination Excel File:")
destination_label.pack()

destination_entry = tk.Entry(app, width=50)
destination_entry.pack()

destination_button = tk.Button(app, text="Select Destination File", command=select_destination_file)
destination_button.pack()

process_button = tk.Button(app, text="Process and Save",
                           command=lambda: process_and_save(source_entry.get(), destination_entry.get()))
process_button.pack()

status_label = tk.Label(app, text="")
status_label.pack()

# Start the main event loop
app.mainloop()