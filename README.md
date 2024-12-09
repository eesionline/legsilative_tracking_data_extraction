# Congressional Tracking Parser

This Python script extracts and organizes information from a `.docx` document containing details of congressional bills and resolutions. It processes the document, extracts key details, and exports them into an Excel file for further analysis.

## Features

- **Parse DOCX Documents:** Reads and processes `.docx` files containing information on bills and resolutions.
- **Extract Bill Details:** Captures data such as:
  - Section (House or Senate)
  - Bill ID (e.g., H.R.123, S.456)
  - Title and Description
  - Sponsor
  - Number of Cosponsors
  - Committees Involved
  - Latest Action
- **Export to Excel:** Outputs the parsed data into a clean and structured `.xlsx` file for easy access and analysis.

## Requirements

Before running the script, ensure you have the following Python packages installed:

- `pandas`: For data handling and exporting to Excel.
- `python-docx`: For reading `.docx` documents.
- `openpyxl`: For writing Excel files.

Install these dependencies using pip:

```bash
pip install pandas python-docx openpyxl
```

## Usage

### 1. Prepare Your Input Document

Create a `.docx` file (SPECIFICALLY NAMED `congressional_tracking.docx`) containing information about bills and resolutions, following a structure like:

```
House:
H.R.123
Bill Title Here
Sponsor: John Doe
Cosponsors: (10)
Committees: Committee on Transportation
Latest Action: Passed the House

Senate:
S.456
Another Bill Title
Sponsor: Jane Smith
Cosponsors: (5)
Committees: Committee on Finance
Latest Action: Referred to Committee
```

### 2. Run the Script

- Place the script and your `.docx` file in the same directory or update the file paths as needed.
- Execute the script:

```bash
python script_name.py
```

### 3. Exported Data

The script processes the `.docx` file and creates an Excel file (e.g., `congressional_tracking_data.xlsx`) in the same directory, containing a table with all extracted data.

## Functions

### `parse_docx(filename)`

- Parses the input `.docx` file and extracts information about bills and resolutions.
- **Input:** Path to the `.docx` file.
- **Output:** A list of dictionaries containing parsed data.

### `export_to_excel(data, output_filename)`

- Converts the parsed data into an Excel file.
- **Input:** 
  - `data`: List of dictionaries with parsed information.
  - `output_filename`: Name of the output Excel file.

## Example Output

After running the script, the Excel file will contain columns like:

| Section | Bill ID  | Title             | Sponsor     | Cosponsors | Committees              | Latest Action           |
|---------|----------|-------------------|-------------|------------|-------------------------|-------------------------|
| House   | H.R.123  | Bill Title Here   | John Doe    | 10         | Committee on Transportation | Passed the House        |
| Senate  | S.456    | Another Bill Title | Jane Smith  | 5          | Committee on Finance    | Referred to Committee   |

## Notes

- Ensure the `.docx` file follows a consistent format to allow the script to extract information accurately.
- The script currently supports only `.docx` input files; additional formats will require modification.

## License

This project is open-source and available under the MIT License. Contributions and improvements are welcome!
