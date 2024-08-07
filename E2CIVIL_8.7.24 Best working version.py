# import necessary libraries
import os
import csv
import datetime
import time
import sys
from colorama import Fore, Back, Style

while True:
    try:
        os.system('cls' if os.name == 'nt' else 'clear')

        # set the directory location
        file_location = 'C:\\Users\\rtabor\\Documents\\POINTDATA'

        # create a dictionary of file names and modified date
        files_dict = {}

        # loop through all files in the directory
        for file in os.listdir(file_location):
            # check if file is a csv
            if file.endswith('.csv'):
                # get the modified date of the file
                mtime = os.path.getmtime(os.path.join(file_location, file))
                # convert modified date to datetime object
                mtime_dt = datetime.datetime.fromtimestamp(mtime)
                # add file to dictionary with key as modified date
                files_dict[mtime_dt] = file

        # sort the list of files by modified date
        sorted_files = sorted(files_dict.items())

        def check_header_format(header_row):
            if any(cell.lower() == 'name' for cell in header_row):
                return "EMLID"
            elif any(cell.replace('.', '').isdigit() for cell in header_row):
                return "CIVIL3D"
            return "UNKNOWN"

        # loop through sorted files and display number, filename, and the format
        for i, file in enumerate(sorted_files):
            # open the file
            with open(os.path.join(file_location, file[1]), 'r') as csvfile:
                # read the csv
                reader = csv.reader(csvfile)
                header_row = next(reader, [])
                result = check_header_format(header_row)
                print(Fore.WHITE + str(i) + ': ' + file[1] + (Fore.GREEN if result == "EMLID" else Fore.BLUE) + ' (' + result + ')')

        # Prompt user to pick a number associated with the file to edit
        user_choice = int(input(Fore.WHITE + 'Pick a number associated with the file to edit: '))

        # Get the filename from the sorted list
        file_name = sorted_files[user_choice][1]

        # Verify the total number of rows in the input file
        total_rows = 0
        with open(os.path.join(file_location, file_name), 'r') as csvfile:
            reader = csv.reader(csvfile)
            total_rows = sum(1 for row in reader)
        print(Fore.WHITE + 'Total rows in input file: ' + str(total_rows))

        # Open the file in read mode
        with open(os.path.join(file_location, file_name), 'r') as csvfile:
            # Read the csv
            reader = csv.reader(csvfile)
            
            # Get the header
            header = next(reader)
            
            # Initialize row_count
            row_count = 1
            row_index = 1  # Initialize row_index to keep track of the rows processed
            
            # Collect rows to write later
            rows_to_write = []
            if header == ['Name', 'Longitude', 'Latitude', 'Ellipsoidal height']:
                for row in reader:
                    if len(row) >= 4:
                        rows_to_write.append([row_count, row[1], row[2], row[3]])
                        row_count += 1
                    else:
                        print(Fore.RED + 'Skipping malformed row: ' + str(row))
                    row_index += 1
                print(Fore.WHITE + file_name + ' has been converted from:' + Fore.GREEN + ' (EMLID)' + Fore.WHITE + ' to' + Fore.BLUE + ' (CIVIL3D)')
            elif 'Longitude' in header and 'Latitude' in header and 'Ellipsoidal height' in header:
                # Handle DAVIES IMPORT V2 format and directly convert to CIVIL3D
                indices = [header.index('Name'), header.index('Longitude'), header.index('Latitude'), header.index('Ellipsoidal height')]
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                for row in reader:
                    if len(row) > max(indices):
                        try:
                            rows_to_write.append([row_count] + [row[i] for i in indices[1:]])
                            row_count += 1
                        except IndexError as e:
                            print(Fore.RED + 'Error in row: ' + str(row) + '. Error: ' + str(e))
                    else:
                        print(Fore.RED + 'Skipping malformed row: ' + str(row))
                    row_index += 1
                print(Fore.WHITE + file_name + ' has been cleaned and converted to CIVIL3D format')
            elif header == ['Name', 'Easting', 'Northing', 'Elevation', 'Description', 'Longitude', 'Latitude', 'Ellipsoidal height', 'Easting RMS', 'Northing RMS', 'Elevation RMS', 'Lateral RMS', 'Antenna height', 'Antenna height units', 'Solution status', 'Averaging start', 'Averaging end', 'Samples', 'PDOP', 'Base easting', 'Base northing', 'Base elevation', 'Base longitude', 'Base latitude', 'Base ellipsoidal height', 'Baseline', 'CS name']:
                # Write the new header for EMLID format
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                for row in reader:
                    if len(row) >= 8:
                        rows_to_write.append([row_count, row[5], row[6], row[7]])
                        row_count += 1
                    else:
                        print(Fore.RED + 'Skipping malformed row: ' + str(row))
                    row_index += 1
                print(Fore.WHITE + file_name + ' has been converted from:' + Fore.BLUE + ' (CIVIL3D)' + Fore.WHITE + ' to' + Fore.GREEN + ' (EMLID)')
            else:
                # Insert a new row on top of all the data, and paste in ['Name', 'Longitude', 'Latitude', 'Ellipsoidal height']
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                rows_to_write.append(header)
                for line in reader:
                    rows_to_write.append(line)
                    row_index += 1
                print(Fore.WHITE + file_name + ' has been converted from: Unknown to' + Fore.GREEN + ' (EMLID)')

        # Write back to the same file
        with open(os.path.join(file_location, file_name), 'w', newline='') as writefile:
            writer = csv.writer(writefile)
            writer.writerows(rows_to_write)

        print(Fore.WHITE + 'Processed ' + str(row_index - 1) + ' rows out of ' + str(total_rows) + ' total rows')
        print("Press enter to continue...")
        while not sys.stdin.read(1):
            time.sleep(1)
    except Exception as e:
        print(Fore.RED + 'An error occurred: ' + str(e))
        print("Press enter to continue...")
        while not sys.stdin.read(1):
            time.sleep(1)