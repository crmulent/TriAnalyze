import os
import subprocess
import pandas as pd
from openpyxl import Workbook
import shutil
import argparse

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="Automates the extraction of OLE files and its exif aand file stream information from pcap files.")
    parser.add_argument('-p', '--pcap', required=True, help="Path to the pcap file")
    parser.add_argument('-f', '--format', default="csv", choices=["excel", "csv", "json"], help="Output format (excel, csv, json)")
    parser.add_argument('-o', '--output', default="Output", help="Path for the output")
    args = parser.parse_args()

    pcap_file = args.pcap
    output_format = args.format
    output_directory = args.output

    # Define input and output directories
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # Define paths to executables
    script_directory = os.path.dirname(os.path.abspath(__file__))
    networkminer_cli_path = os.path.join(script_directory, "netminercli")
    networkminer_cli_exe_path = os.path.join(networkminer_cli_path, "NetworkMinerCLI.exe")
    oledump_path = os.path.join(script_directory, "oledump", "oledump.py")
    exiftool_path = os.path.join(script_directory, "exiftool", "exiftool.exe")

    # Step 1: Run NetworkMinerCLI to extract files
    networkminer_output = os.path.join(networkminer_cli_path, 'AssembledFiles')
    if not os.path.exists(networkminer_output):
        os.makedirs(networkminer_output)

    subprocess.run([networkminer_cli_exe_path, pcap_file])

    # Step 2: Analyze OLE files using Oledump
    ole_files = []
    for root, dirs, files in os.walk(networkminer_output):
        for file in files:
            if file.endswith('.doc') or file.endswith('.xls') or file.endswith('.docx'):
                ole_files.append(os.path.join(root, file))

    # Ensure olefile is installed
    subprocess.run(['pip', 'install', 'olefile'])

    print("Oledump.py is processing...")
    oledump_results = []
    for file in ole_files:
        file_path = os.path.join(networkminer_output, file)
        result = subprocess.run(['python', oledump_path, file_path], capture_output=True, text=True)
        oledump_results.append({'File': file, 'Oledump_Output': result.stdout})

    # Step 3: Extract metadata using Exiftool
    print("Exiftool is processing...")
    exiftool_results = []
    for file in ole_files:
        file_path = os.path.join(networkminer_output, file)

        output = ""
        # Check if the filename includes '-k'
        if '-k' in exiftool_path:
            # Use Popen and communicate to handle the Enter keypress
            process = subprocess.Popen([exiftool_path, file_path], stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
            output, _ = process.communicate(input='\n')  # Simulate Enter keypress
        else:
            # Use subprocess.run for files without '-k'
            result = subprocess.run([exiftool_path, file_path], capture_output=True, text=True)
            output = result.stdout
        exiftool_results.append({'File': file, 'Exiftool_Output': output})

    # Step 4: Compile results into a DataFrame
    print("Compilation...")
    df_oledump = pd.DataFrame(oledump_results)
    df_exiftool = pd.DataFrame(exiftool_results)

    # Debug: Print DataFrame columns
    # print("Columns in df_oledump:", df_oledump.columns.tolist())
    # print("Columns in df_exiftool:", df_exiftool.columns.tolist())

    # Ensure 'File' column exists in both DataFrames
    if 'File' not in df_oledump.columns or 'File' not in df_exiftool.columns:
        raise ValueError("One or both DataFrames do not have a 'File' column.")

    # Combine DataFrames
    df_combined = pd.merge(df_oledump, df_exiftool, on='File', how='outer')

    # Step 5: Export to chosen format
    if output_format == 'excel':
        excel_file = os.path.join(output_directory, 'results.xlsx')
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False, sheet_name='Combined_Results')
        print(f"Results have been exported to {excel_file}")
    elif output_format == 'csv':
        csv_file = os.path.join(output_directory, 'results.csv')
        df_combined.to_csv(csv_file, index=False)
        print(f"Results have been exported to {csv_file}")
    elif output_format == 'json':
        json_file = os.path.join(output_directory, 'results.json')
        df_combined.to_json(json_file, orient='records', lines=True)
        print(f"Results have been exported to {json_file}")
    else:
        raise ValueError("Invalid output format specified.")

    # Step 6: Prompt user to decide whether to keep the contents of the networkminer_output folder
    keep_files = input("Keep network miner files? [y/n]: ").strip().lower()
    if keep_files == 'y':
        print(f"The contents of {networkminer_output} will be kept.")
    elif keep_files == 'n':
        if os.path.exists(networkminer_output):
            shutil.rmtree(networkminer_output)
            os.makedirs(networkminer_output)
        print(f"The contents of {networkminer_output} have been deleted and the directory has been recreated.")
    else:
        print("Invalid input. The contents of the directory will be kept by default.")
    
if __name__ == "__main__":
    main()