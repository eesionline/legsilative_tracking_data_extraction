import pandas as pd
from docx import Document
import re

def parse_docx(filename):
    # Load the document
    doc = Document(filename)

    # Initialize data storage
    bills_data = []

    # Set variables to track the current section (House or Senate)
    current_section = None
    current_bill = {}
    expecting_title = False

    # Regular expressions to match bill information
    bill_start_re = re.compile(r'^(H\.R\.\d+|H\.Res\.\d+|S\.\d+|S\.Res\.\d+)')
    sponsor_re = re.compile(r'Sponsor: (.+)')
    cosponsors_re = re.compile(r'Cosponsors: \((\d+)\)')
    committees_re = re.compile(r'Committees: (.+)')
    latest_action_re = re.compile(r'Latest Action: (.+)')

    # Iterate through each paragraph in the document
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        
        # Detect section headers (House and Senate)
        if text == "House:":
            current_section = "House"
            continue
        elif text == "Senate:":
            current_section = "Senate"
            continue
        
        # Skip empty lines
        if not text:
            continue

        # Check if the line starts with a bill identifier
        bill_match = bill_start_re.match(text)
        if bill_match:
            # If we were parsing a previous bill, store it first
            if current_bill:
                bills_data.append(current_bill)
            # Start a new bill entry and expect the title next
            current_bill = {
                "section": current_section,
                "bill_id": bill_match.group(1),
                "title": "",
                "description": "",
                "sponsor": "",
                "cosponsors": 0,
                "committees": "",
                "latest_action": ""
            }
            expecting_title = True
            continue

        # If we are expecting a title, assign the line as the title
        if expecting_title:
            current_bill["title"] = text
            expecting_title = False
            continue

        # Check for other fields based on the current line
        sponsor_match = sponsor_re.search(text)
        if sponsor_match:
            current_bill["sponsor"] = sponsor_match.group(1)
            continue
        
        cosponsors_match = cosponsors_re.search(text)
        if cosponsors_match:
            current_bill["cosponsors"] = int(cosponsors_match.group(1))
            continue
        
        committees_match = committees_re.search(text)
        if committees_match:
            current_bill["committees"] = committees_match.group(1)
            continue

        latest_action_match = latest_action_re.search(text)
        if latest_action_match:
            current_bill["latest_action"] = latest_action_match.group(1)
            continue

    # Add the last bill if any
    if current_bill:
        bills_data.append(current_bill)

    return bills_data

def export_to_excel(data, output_filename):
    # Create a DataFrame from the parsed data
    df = pd.DataFrame(data)

    # Export DataFrame to Excel
    df.to_excel(output_filename, index=False)

# Usage example
filename = "congressional_tracking.docx"
output_filename = "congresional_tracking_data.xlsx"

# Parse the document and export to Excel
bills_data = parse_docx(filename)
export_to_excel(bills_data, output_filename)

print(f"Data successfully exported to {output_filename}")