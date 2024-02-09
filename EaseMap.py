import openpyxl
import time

print("\n\t")
print("""
                        _____                        __  __                 
                        | ____|   __ _   ___    ___  |  \/  |   __ _   _ __  
                        |  _|    / _` | / __|  / _ \ | |\/| |  / _` | | '_ \ 
                        | |___  | (_| | \__ \ |  __/ | |  | | | (_| | | |_) |
                        |_____|  \__,_| |___/  \___| |_|  |_|  \__,_| | .__/ 
                                                                    |_|    

                                                                                Version: v1.0
                                                                                By: LegionSec - Shubham Singh
""")

def parse_nmap_report(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()

    results = []
    current_result = None

    for line in lines:
        if line.startswith("Nmap scan report for "):
            if current_result:
                results.append(current_result)
            host = line.split()[-1].strip()
            current_result = {"Host": host, "Ports": []}
        elif line.startswith("Device type: "):
            current_result["Device-Type"] = line.split(": ", 1)[1].strip()
        elif line.startswith("OS CPE: "):
            current_result["OS-CPE"] = line.split(": ", 1)[1].strip().split()
        elif line.startswith("Running (JUST GUESSING): "):
            current_result["OS-Guesses"] = line.split(": ", 1)[1].strip().split()
        elif "/tcp" in line:
            parts = line.split()
            port, status, service = parts[0].split("/")[0], parts[1], parts[2]
            current_result["Ports"].append({"Port": port, "Status": status, "Service": service})

    if current_result:
        results.append(current_result)

    return results

def create_excel_file(results):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append(['Host', 'Port', 'Service', 'Status', 'Device-Type', 'OS-CPE', 'OS-Guesses'])

    for result in results:
        host = result["Host"]
        device_type = result.get("Device-Type", "")
        os_cpe = result.get("OS-CPE", [])
        os_guesses = result.get("OS-Guesses", [])

        for port_info in result["Ports"]:
            sheet.append([host, port_info["Port"], port_info["Service"], port_info["Status"],
                          device_type, ', '.join(os_cpe), ', '.join(os_guesses)])

    excel_file_path = 'nmap_report.xlsx'
    workbook.save(excel_file_path)
    print("\n\nProcessing your request ..")
    time.sleep(3)
    print(f'\nExcel file created: {excel_file_path}')

if __name__ == "__main__":
    print("\n\n")
    print("NOTE: This script converts your normal nmap output from .txt to .xlsx file. For any bug or query, please contact me!")
    print("\n\n")
    file_path = input("Please enter the file path: ")
    results = parse_nmap_report(file_path)
    create_excel_file(results)
