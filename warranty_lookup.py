import requests
import csv
import os
import sys


"""
This script uses an API to locate warranty information for Lenovo systems. Eventually more than just Lenovo
Supports just serial numbers and csv files with serial numbers.
"""
#Only serial numbers
def query_api(serial, machine_type=None):
    url = "https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/json"
    }

    payload = {"serialNumber": serial}
    if machine_type:
        payload["machineType"] = machine_type
    
    try:
        response = requests.post(url, json=payload, headers=headers)
        if response.status_code == 200:

            # Retrieve Warranty info and Model info.
            data = response.json()

            startDate = data['data']['baseWarranties'][0]['startDate']
            endDate = data['data']['baseWarranties'][0]['endDate']
            productModel = data['data']['machineInfo']['subSeries']

            return {
                "Start/Purchase Date": startDate,
                "Warranty End": endDate,
                "Product Model": productModel}

    except Exception as e:
        return {
                "Start/Purchase Date": "None",
                "Warranty End": "None",
                "Product Model": "None"}
    return None

def process_csv(input_file, progress_callback=None):
    base, ext = os.path.splitext(input_file)
    output_file = f"{base}_updated{ext}"

    with open(input_file, newline='') as infile, open(output_file, mode='w', newline='') as outfile:
        reader = csv.DictReader(infile)
        rows = list(reader)
        total = len(rows)
        fieldnames = reader.fieldnames
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()

        # Iterate through each row in the CSV
        for idx, row in enumerate(rows):
            serial = row['Serial Number'].strip()
            asset_tag = row['Asset Tag'].strip()
            manufacturer = row['Manufacturer'].strip()

            #make search faster
            if serial == '' and asset_tag == '':
                continue
            if manufacturer.lower() not in ['lenovo', None, 'lenovo monitor', 'lenovo laptop', 'lenovo desktop', 'lenovo dock']:
                result = {
                "Start/Purchase Date": "None",
                "Warranty End": "None",
                "Product Model": "None"}
            else:
                result = query_api(serial)
                
            #machine_type = row.get('MTM Number', '').strip()
            #print(f"Searching for warranty info for {serial}...")
            
            row.update(result)
            writer.writerow(row)

            if progress_callback:
                progress_callback(idx + 1, total)

    return output_file

if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        print("Usage:")
        print("python warranty_lookup.py <file.csv>")
        print("python warranty_lookup.py list <serial_number> <serial_number> etc...")
        sys.exit(1)
    
    if args[0].lower().endswith(".csv"):
        process_csv(args[0])
    
    elif args[0].lower() == "list" and len(args) > 1:
        for serial in args[1:]:
            result = query_api(serial)
            print(f"{serial} -> Start Date: {result['Start/Purchase Date']}, End Date: {result['Warranty End']}, Product Model: {result['Product Model']}")

    else:
        print("Usage:")
        print("python warranty_lookup.py <file.csv>")
        print("python warranty_lookup.py list <serial_number> <serial_number> etc...")