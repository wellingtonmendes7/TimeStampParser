import os
import re
from openpyxl import Workbook

# Define the input and output directory
input_directory = r"D:\Praat"
output_file = os.path.join(input_directory, "output_timestamps.xlsx")

# Regular expressions for parsing the TextGrid file
interval_pattern = re.compile(r'\n(\d+\.\d+)\n(\d+\.\d+)\n(?:"([^"]*)"|"")')

def process_textgrid_file(file_path, file_name):
    """Extracts intervals from a TextGrid file and returns them as a list."""
    intervals = []
    with open(file_path, "r", encoding="utf-8") as file:
        content = file.read()
        matches = interval_pattern.findall(content)
        for match in matches:
            start_time, end_time, label = match
            if label.strip():  # Only include intervals with a non-empty label
                intervals.append((file_name, start_time, end_time, label.strip()))
    return intervals

def main():
    all_intervals = []

    # Process all .textgrid files in the directory
    for file_name in os.listdir(input_directory):
        if file_name.lower().endswith(".textgrid"):
            file_path = os.path.join(input_directory, file_name)
            intervals = process_textgrid_file(file_path, file_name)
            all_intervals.extend(intervals)

    # Create an Excel workbook and add a sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Intervals"

    # Write the header row
    ws.append(["File Name", "Start Time", "End Time", "Label"])

    # Write the intervals to the Excel sheet
    for interval in all_intervals:
        ws.append(interval)

    # Save the workbook to the output file
    wb.save(output_file)

    print(f"Extraction completed. Results saved to {output_file}")

if __name__ == "__main__":
    main()
