# Excel Data Populator

The **Excel Data Populator** is a Python script designed to populate an Excel spreadsheet's columns with data from another spreadsheet while ensuring unique entries. This can be particularly useful for tasks like combining data from different sources. This script uses the `openpyxl` library to achieve this functionality.

## Prerequisites

Before you begin, make sure you have the following:

- Python (3.6 or higher) installed on your system.
- The `openpyxl` library. If you don't have it installed, you can install it using the following command: `pip install openpyxl`


## Usage

1. Clone or download this repository to your local machine.

2. Place your input Excel files in the same directory as the script. The script expects two input files:

 - `example_file1.xlsx`: The spreadsheet to be updated.
 - `example_export.xlsx`: The spreadsheet containing the data to be transferred.

3. Open a terminal or command prompt and navigate to the directory where the script is located.

4. Run the script using the following command: `python [.py file]`

5. The script will populate the columns in `example_file1.xlsx` with the data from `example_export.xlsx` while ensuring unique entries.

6. Once the script has finished running, you'll find the updated data in the `example_file1.xlsx` spreadsheet.

## Script Explanation

The script works as follows:

- It loads the input spreadsheets using the `openpyxl` library.
- It defines the column indexes for "Home #" and "Street Name & Notes" in the `example_file1.xlsx` spreadsheet.
- It iterates over the rows in the `example_export.xlsx` spreadsheet, skipping the header.
- For each row, it checks if the street number has been populated before. If not, it populates the "Home #" and "Street Name & Notes" columns in the `example_file1.xlsx` spreadsheet.
- It ensures that each street number is populated only once using a set called `populated_street_numbers`.
- Finally, it saves the updated `example_file1.xlsx` spreadsheet with the populated data.

---

Feel free to contribute to this project by submitting issues or pull requests. If you have any questions or need further assistance, please don't hesitate to contact the repository owner.
