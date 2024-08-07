import os
import csv
import datetime
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, filename='app.log', filemode='w',
                    format='%(name)s - %(levelname)s - %(message)s')

# Set the default directory location
DEFAULT_FILE_LOCATION = 'C:\\Users\\rtabor\\Documents\\POINTDATA'
ROWS_PER_PAGE = 100  # Number of rows to show per page

def check_header_format(header_row):
    if any(cell.lower() == 'name' for cell in header_row):
        return "EMLID"
    elif any(cell.replace('.', '').isdigit() for cell in header_row):
        return "CIVIL3D"
    return "UNKNOWN"

def process_files(sort_by='date', descending=True):
    try:
        logging.info('Processing files...')
        # Clear the file list
        for widget in file_list_frame.winfo_children():
            widget.destroy()
        
        # Set the directory location
        file_location = file_location_entry.get()
        if not file_location:
            return

        # Create a dictionary of file names and modified date
        files_dict = {}

        # Loop through all files in the directory
        for file in os.listdir(file_location):
            # Check if file is a csv
            if file.endswith('.csv'):
                # Get the modified date of the file
                mtime = os.path.getmtime(os.path.join(file_location, file))
                # Convert modified date to datetime object
                mtime_dt = datetime.datetime.fromtimestamp(mtime)
                # Add file to dictionary with key as modified date
                files_dict[file] = (file, mtime_dt)

        # Sort the list of files
        sorted_files = sorted(files_dict.values(), key=lambda x: x[0 if sort_by == 'name' else 1], reverse=descending)

        # Display files in the listbox with select buttons
        for i, (file_name, _) in enumerate(sorted_files):
            with open(os.path.join(file_location, file_name), 'r') as csvfile:
                reader = csv.reader(csvfile)
                header_row = next(reader, [])
                result = check_header_format(header_row)
                color = 'green' if result == "EMLID" else 'blue' if result == "CIVIL3D" else 'white'
                
                file_frame = tk.Frame(file_list_frame, bg='#000000')
                file_label = tk.Label(file_frame, text=f"{i}: {file_name} ({result})", fg=color, bg='#000000', font=("Helvetica", 16, "bold"), anchor="w")
                file_label.pack(side=tk.LEFT, padx=10, pady=5)
                file_label.bind("<Button-1>", lambda e, f=file_name: ask_conversion(f))
                
                file_frame.pack(fill=tk.X)

        # Store the sorted files globally for later use
        global sorted_files_global
        sorted_files_global = sorted_files
        logging.info('Files processed successfully.')

    except Exception as e:
        logging.error(f"An error occurred while processing files: {e}")

def ask_conversion(file_name):
    # Create a new window for the prompt
    prompt_window = tk.Toplevel(root)
    prompt_window.title("Confirm Conversion")
    prompt_window.geometry(f'{root.winfo_width()}x{root.winfo_height()//2}+{root.winfo_x()}+{root.winfo_y() + root.winfo_height()//2}')
    prompt_window.configure(bg='#000000')

    # Add prompt text
    prompt_label = tk.Label(prompt_window, text=f"Do you want to convert {file_name}?", fg='white', bg='#000000')
    prompt_label.pack(pady=10)

    # Preview the CSV file content
    preview_frame = tk.Frame(prompt_window, bg='#000000')
    preview_frame.pack(fill=tk.BOTH, expand=True)

    preview_canvas = tk.Canvas(preview_frame, bg='#000000')
    preview_scrollbar = tk.Scrollbar(preview_frame, orient="vertical", command=preview_canvas.yview)
    preview_content_frame = tk.Frame(preview_canvas, bg='#000000')
    preview_content_frame.bind(
        "<Configure>",
        lambda e: preview_canvas.configure(
            scrollregion=preview_canvas.bbox("all")
        )
    )
    preview_canvas.create_window((0, 0), window=preview_content_frame, anchor="nw")
    preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
    preview_canvas.pack(side="left", fill="both", expand=True)
    preview_scrollbar.pack(side="right", fill="y")

    def load_preview_page(page_number):
        # Clear previous content
        for widget in preview_content_frame.winfo_children():
            widget.destroy()

        file_location = file_location_entry.get()
        with open(os.path.join(file_location, file_name), 'r') as csvfile:
            reader = csv.reader(csvfile)
            for _ in range(page_number * ROWS_PER_PAGE):
                next(reader, None)  # Skip rows

            for i, row in enumerate(reader):
                if i >= ROWS_PER_PAGE:
                    break
                row_color = '#000000' if i % 2 == 0 else '#111111'
                row_frame = tk.Frame(preview_content_frame, bg=row_color)
                for cell in row:
                    cell_label = tk.Label(row_frame, text=cell, fg='white', bg=row_color, anchor="w")
                    cell_label.pack(side=tk.LEFT, padx=5, pady=2)
                row_frame.pack(fill=tk.X)

    def confirm_conversion():
        edit_file(file_name)
        prompt_window.destroy()
    
    def cancel_conversion():
        prompt_window.destroy()

    button_frame = tk.Frame(prompt_window, bg='#000000')
    button_frame.pack(pady=10)

    yes_button = ttk.Button(button_frame, text="Yes", command=confirm_conversion)
    yes_button.pack(side=tk.LEFT, padx=20, pady=10)

    no_button = ttk.Button(button_frame, text="Cancel", command=cancel_conversion)
    no_button.pack(side=tk.RIGHT, padx=20, pady=10)

    def on_preview_mouse_wheel(event):
        preview_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    preview_canvas.bind("<Enter>", lambda e: preview_canvas.bind_all("<MouseWheel>", on_preview_mouse_wheel))
    preview_canvas.bind("<Leave>", lambda e: preview_canvas.unbind_all("<MouseWheel>"))

    load_preview_page(0)  # Load the first page

def edit_file(file_name):
    try:
        logging.info(f'Editing file: {file_name}')
        file_location = file_location_entry.get()
        
        # Verify the total number of rows in the input file
        total_rows = 0
        with open(os.path.join(file_location, file_name), 'r') as csvfile:
            reader = csv.reader(csvfile)
            total_rows = sum(1 for row in reader)

        # Open the file in read mode
        with open(os.path.join(file_location, file_name), 'r') as csvfile:
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
                        pass
                    row_index += 1
            elif 'Longitude' in header and 'Latitude' in header and 'Ellipsoidal height' in header:
                indices = [header.index('Name'), header.index('Longitude'), header.index('Latitude'), header.index('Ellipsoidal height')]
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                for row in reader:
                    if len(row) > max(indices):
                        try:
                            rows_to_write.append([row_count] + [row[i] for i in indices[1:]])
                            row_count += 1
                        except IndexError:
                            pass
                    else:
                        pass
                    row_index += 1
            elif header == ['Name', 'Easting', 'Northing', 'Elevation', 'Description', 'Longitude', 'Latitude', 'Ellipsoidal height', 'Easting RMS', 'Northing RMS', 'Elevation RMS', 'Lateral RMS', 'Antenna height', 'Antenna height units', 'Solution status', 'Averaging start', 'Averaging end', 'Samples', 'PDOP', 'Base easting', 'Base northing', 'Base elevation', 'Base longitude', 'Base latitude', 'Base ellipsoidal height', 'Baseline', 'CS name']:
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                for row in reader:
                    if len(row) >= 8:
                        rows_to_write.append([row_count, row[5], row[6], row[7]])
                        row_count += 1
                    else:
                        pass
                    row_index += 1
            else:
                rows_to_write.append(['Name', 'Longitude', 'Latitude', 'Ellipsoidal height'])
                rows_to_write.append(header)
                for line in reader:
                    rows_to_write.append(line)
                    row_index += 1

        # Write back to the same file
        with open(os.path.join(file_location, file_name), 'w', newline='') as writefile:
            writer = csv.writer(writefile)
            writer.writerows(rows_to_write)

        logging.info(f'File {file_name} processed successfully.')
        refresh_files()  # Refresh the file list after conversion
    except Exception as e:
        logging.error(f"An error occurred while editing the file {file_name}: {e}")

def refresh_files():
    process_files()

def toggle_sorting(sort_by):
    global sort_descending
    sort_descending = not sort_descending
    process_files(sort_by=sort_by, descending=sort_descending)

# Create the main window
root = tk.Tk()
root.title("Tabor's CSV Wizard")
root.geometry('550x800')
root.configure(bg='#000000')

# Create a style
style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", font=("Helvetica", 12), background='#000000', foreground='white')
style.configure("TButton", font=("Helvetica", 12), background='#555555', foreground='white')
style.configure("TEntry", font=("Helvetica", 12), fieldbackground='#3a3a3a', foreground='white')

# Load and place the graphic
image_path = os.path.join(os.path.dirname(__file__), "csvwizard.webp")
try:
    image = Image.open(image_path)
    image = image.resize((550, 150), Image.LANCZOS)
    photo = ImageTk.PhotoImage(image)

    graphic_label = tk.Label(root, image=photo, bg='#000000')
    graphic_label.pack(pady=10)
    logging.info('Graphic loaded successfully.')
except Exception as e:
    logging.error(f"An error occurred while loading the graphic: {e}")

# Create and place the directory entry
file_location_entry = ttk.Entry(root, width=75)
file_location_entry.pack(pady=5)
file_location_entry.insert(0, DEFAULT_FILE_LOCATION)

# Create and place the top frame with all buttons and entry
top_frame = tk.Frame(root, bg='#000000')
top_frame.pack(pady=10, fill=tk.X)

# Center the buttons
buttons = ["Process Files", "Refresh Files", "Sort by Name", "Sort by Date"]
commands = [process_files, refresh_files, lambda: toggle_sorting('name'), lambda: toggle_sorting('date')]
for text, command in zip(buttons, commands):
    button = ttk.Button(top_frame, text=text, command=command)
    button.pack(side=tk.LEFT, padx=5, pady=5, expand=True)

# Create a frame with a scrollbar for the file list
file_list_container = tk.Frame(root, bg='#000000')
file_list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
canvas = tk.Canvas(file_list_container, bg='#000000')
scrollbar = tk.Scrollbar(file_list_container, orient="vertical", command=canvas.yview)
file_list_frame = tk.Frame(canvas, bg='#000000')
file_list_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)
canvas.create_window((0, 0), window=file_list_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Enable scrolling with the mouse wheel
def on_mouse_wheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mouse_wheel))
canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

# Automatically process files from the default directory
sort_descending = True
process_files()

# Run the application
root.mainloop()
