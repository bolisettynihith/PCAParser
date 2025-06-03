# References:
# - https://www.sygnia.co/blog/new-windows-11-pca-artifact/
# - https://aboutdfir.com/new-windows-11-pro-22h2-evidence-of-execution-artifact/

# Code inspiration:
# - https://github.com/AndrewRathbun/PCAParser

import os
import csv
import argparse
from datetime import datetime

# Format's date and time and strips the last 3 decimals
def format_timestamp(timestamp_str):
    try:
        dt = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S.%f")
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return timestamp_str

# Parses the PcaAppLaunchDic.txt and converts it to a CSV file
def Parse_PcaAppLaunchDic(pcaapplaunchTXT, csv_path):
    print(f"[+] Parsing - {os.path.abspath(pcaapplaunchTXT)}")
    with open(pcaapplaunchTXT, 'r', encoding='utf-8') as in_file:
        applaunch_lines = in_file.read()
        applaunch_lines = applaunch_lines.splitlines()

        FIELDNAMES = ["Execution Time", "File Path"]

        with open(csv_path, 'w', newline='', encoding='utf-8') as out_file:
            writer = csv.DictWriter(out_file, fieldnames=FIELDNAMES)
            writer.writeheader()

            for line in applaunch_lines:
                parts = line.split('|')

                clean_time = format_timestamp(parts[1])
                row = {
                    "Execution Time": clean_time,
                    "File Path": parts[0]
                }
                writer.writerow(row)
    # Checks whether the CSV file is created and returns a output with the path
    if(os.path.exists(csv_path)):
        print(f"[+] Created - {os.path.abspath(csv_path)}\n")

# Parses the PcaGeneralDb*.txt and converts it to a CSV file
def Parse_PcaGeneralDb(pcaGeneralTXT, csv_path):
    print(f"[+] Parsing - {os.path.abspath(pcaGeneralTXT)}")
    FIELDNAMES = ["Creation Time", "Record Type", "File Path", "Product Name", "Company Name", "Product Version", "Program ID", "Message"]

    RECORD_TYPE_MAP = {
        "0": "Installer failed (0)",
        "1": "Driver was Blocked (1)",
        "2": "Abnormal Process Exit (2)",
        "3": "PCA Resolve is called (3)"
    }
    with open(pcaGeneralTXT, 'r', encoding='utf-16le') as in_file, \
        open(csv_path, 'w', newline='', encoding='utf-8') as out_file:

            writer = csv.DictWriter(out_file, fieldnames=FIELDNAMES)
            writer.writeheader()

            pcaGeneral_lines = in_file.read()
            pcaGeneral_lines = pcaGeneral_lines.splitlines()

            for line in pcaGeneral_lines:
                parts = line.split('|')
                if len(parts) != 8:
                    print(f"Skipping malformed line: {line}")
                    continue

                clean_time = format_timestamp(parts[0])
                record_type = RECORD_TYPE_MAP.get(parts[1], "Unknown")

                row = {
                    "Creation Time": clean_time, 
                    "Record Type": record_type, 
                    "File Path": parts[2], 
                    "Product Name": parts[3], 
                    "Company Name": parts[4], 
                    "Product Version": parts[5], 
                    "Program ID": parts[6], 
                    "Message": parts[7]
                }
                writer.writerow(row)
    if(os.path.exists(csv_path)):
        print(f"[+] Created - {os.path.abspath(csv_path)}\n")

# Generates a sorted timeline of the executions from both the txt files (PcaAppLaunchDic and PcaGeneralDb*.txt) and generates a CSV file
def Parse_PCATimeline(in_folder, out_folder):
    FIELDNAMES = ["Execution Time", "File Path"]
    out_file_path = os.path.join(out_folder, "PCATimeline.csv")

    # Mapping of filename to a tuple: (time index, path index)
    file_parsers = {
        'PcaAppLaunchDic.txt': (1, 0),
        'PcaGeneralDb0.txt': (0, 2),
        'PcaGeneralDb1.txt': (0, 2),
    }

    # Set encoding per file
    encoding_map = {
        'PcaGeneralDb0.txt': 'utf-16le',
        'PcaGeneralDb1.txt': 'utf-16le',
        'PcaAppLaunchDic.txt': 'utf-8'
    }

    all_rows = []

    for filename in os.listdir(in_folder):
        if filename in file_parsers:
            file_path = os.path.join(in_folder, filename)
            if not os.path.isfile(file_path):
                print(f'[-] New file found: {file_path}')
                continue

            time_idx, path_idx = file_parsers[filename]
            with open(file_path, 'r', encoding=encoding_map.get(filename, 'utf-8')) as in_file:
                for line in in_file:
                    parts = line.strip().split('|')
                    if len(parts) > max(time_idx, path_idx):
                        row = {
                            "Execution Time": format_timestamp(parts[time_idx]),
                            "File Path": parts[path_idx]
                        }
                        all_rows.append(row)
    
    def parse_time(row):
        try:
            return datetime.strptime(row["Execution Time"], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return datetime.min
        
    all_rows.sort(key=parse_time)

    with open(out_file_path, 'w', newline='', encoding='utf-8') as out_csv:
        writer = csv.DictWriter(out_csv, fieldnames=FIELDNAMES)
        writer.writeheader()
        for row in all_rows:
            writer.writerow(row)

    if(os.path.exists(out_file_path)):
        print(f"[+] Created timeline at - {os.path.abspath(out_file_path)}\n")

def PCAParser(input_folder, output_folder):
    if(os.path.exists(output_folder)):
        pass
    elif(output_folder == 'Reports'):
        os.mkdir(output_folder)
        print(f"[+] Output folder isn't provided. Creating the default - '{os.path.abspath(output_folder)}' folder.\n")
    else:
        print("[+] Output Folder doesn't exist!")

    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder,filename)

        if os.path.isfile(file_path):
            if filename == 'PcaAppLaunchDic.txt':
                csv_path = os.path.join(output_folder, "PcaAppLaunchDic.csv")
                Parse_PcaAppLaunchDic(file_path, csv_path)
            elif filename == 'PcaGeneralDb0.txt':
                csv_path = os.path.join(output_folder, "PcaGeneralDb0.csv")
                Parse_PcaGeneralDb(file_path, csv_path)
            elif filename == 'PcaGeneralDb1.txt':
                csv_path = os.path.join(output_folder, "PcaGeneralDb1.csv")
                Parse_PcaGeneralDb(file_path, csv_path)
        else:
            print(f"[-] New file found - {filename}")

    Parse_PCATimeline(input_folder, output_folder)
            

def main():
    parser = argparse.ArgumentParser(description='Windows Program Compatibility Assistant (C:\Windows\appcompat\pca) Parser')
    parser.add_argument('-i', '--input_folder', required=True, action="store", help='Location to text files that reside in C:\Windows\appcompat\pca')
    parser.add_argument('-o','--output_folder', required=False, action="store", help='Path to output folder, if not mentioned reports will be generated to default folder (default: Reports)')

    args = parser.parse_args()
    input_path = args.input_folder
    output_path = args.output_folder

    # Checking if output folder is current folder or No output folder
    if (os.path.exists(input_path) and (output_path == None or (os.path.abspath(output_path) == os.getcwd()))):
        PCAParser(input_path, 'Reports')
    elif(os.path.exists(input_path) and os.path.exists(output_path)):
        PCAParser(input_path, output_path)
    elif(not (os.path.exists(output_path))):
        os.mkdir(output_path)
        print(f"[+] Output folder does not exist. Creating the - '{os.path.abspath(output_path)}' folder.\n")
        PCAParser(input_path, output_path)
    else:
        print(parser.print_help())

if __name__ == '__main__':
    main()