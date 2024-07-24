import win32print
import socket
from concurrent.futures import ThreadPoolExecutor, as_completed

# Function to get the IP address from the port name
def get_printer_ip(port_name):
    try:
        # Attempt to resolve the port name to an IP address
        ip_address = socket.gethostbyname(port_name)
        return ip_address
    except Exception as e:
        return f"Unable to determine IP address: {e}"

# Function to get the status of the printer
def get_printer_status(status):
    # Dictionary mapping status codes to status messages
    status_dict = {
        0: "Ready",
        1: "Error",
        2: "Pending Deletion",
        3: "Paper Jam",
        4: "Paper Out",
        5: "Manual Feed",
        6: "Paper Problem",
        7: "Offline",
        8: "IO Active",
        9: "Busy",
        10: "Printing",
        11: "Output Bin Full",
        12: "Not Available",
        13: "Waiting",
        14: "Processing",
        15: "Initializing",
        16: "Warming Up",
        17: "Toner Low",
        18: "No Toner",
        19: "Page Punt",
        20: "User Intervention Required",
        21: "Out of Memory",
        22: "Door Open",
        23: "Server_Unknown",
        24: "Power Save"
    }
    return status_dict.get(status, "Unknown")

# Function to get and print detailed information about a printer
def get_printer_info(printer_name):
    try:
        # Open the printer to get a handle
        handle = win32print.OpenPrinter(printer_name)
        # Retrieve detailed properties of the printer
        properties = win32print.GetPrinter(handle, 2)
        print(f"Printer: {printer_name}")
        print(f"    Driver: {properties['pDriverName']}")
        print(f"    Status: {get_printer_status(properties['Status'])}")
        ip_address = get_printer_ip(properties['pPortName'])
        print(f"    IP Address: {ip_address}")
    except Exception as e:
        print(f"Error fetching printer properties: {e}")
    finally:
        # Ensure the printer handle is closed
        win32print.ClosePrinter(handle)

# Function to check if a printer matches the given IP address
def check_printer_ip(printer, ip_address):
    try:
        # Open the printer to get a handle
        handle = win32print.OpenPrinter(printer[2])
        # Retrieve limited properties (PRINTER_INFO_5) of the printer
        properties = win32print.GetPrinter(handle, 5)
        printer_ip = get_printer_ip(properties['pPortName'])
        # Close the printer handle
        win32print.ClosePrinter(handle)
        # Check if the printer IP matches the given IP address
        if printer_ip == ip_address:
            return printer[2]
    except Exception:
        pass
    return None

# Function to get the printer name by IP address
def get_printer_name_by_ip(ip_address, server_name=None):
    if server_name:
        # Enumerate printers on the specified server
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, server_name, 1)
    else:
        # Enumerate local and connected printers
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        
    # Use ThreadPoolExecutor to check multiple printers
    with ThreadPoolExecutor() as executor:
        # Submit tasks to the executor
        future_to_printer = {executor.submit(check_printer_ip, printer, ip_address): printer for printer in printers}
        for future in as_completed(future_to_printer):
            # Get the result of each completed task
            printer_name = future.result()
            if printer_name:
                return printer_name
    return None

# Main function to prompt the user and handle input
def main():
    while True:
        print("Enter 'name' to search by printer name, 'ip' to search by IP address, or 'exit' to close:")
        choice = input("> ").lower()

        if choice == 'exit':
            break
        elif choice == 'name':
            # Prompt the user to enter the printer name
            print("Enter the name of the printer to fetch information:")
            printer_name = input("> ")
            get_printer_info(printer_name)
        elif choice == 'ip':
            # Prompt the user to enter the IP address
            print("Enter the IP address of the printer to fetch information:")
            ip_address = input("> ")
            # Prompt the user to enter the server name
            print("Enter the server name (leave blank for local)")
            server_name = input(">")
            printer_name = get_printer_name_by_ip(ip_address, server_name if server_name else None)
            if printer_name:
                get_printer_info(printer_name)
            else:
                print(f"No printer found with IP address: {ip_address}")
        else:
            print("Invalid choice. Please enter 'name' or 'ip'.")
        print("\n")

if __name__ == "__main__":
    main()
