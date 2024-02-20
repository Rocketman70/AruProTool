import serial
import time
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import subprocess
from Classes import APNamerGUI 

def restart(self):
        # Clear the output textbox
        self.output_textbox.delete(1.0, tk.END)
        # Restart the application with the same Excel file and COM port
        self.start_serial(self.com_port, self.file_path)
        self.check_com_ports()

def select_excel_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.output_textbox.delete(1.0, tk.END)
            self.update_output("Selected Excel file: " + file_path)
            self.file_path = file_path  # Update the stored file path
            self.process_excel_file(file_path)

def update_output(self, message):
        self.output_textbox.insert(tk.END, message + '\n')
        self.output_textbox.see(tk.END)

def check_com_ports(self):
        first_instance_find_com = True
        while True:
            self.com_port = self.find_com_port()
            if self.com_port and first_instance_find_com:
                self.update_output("COM Port found: " + self.com_port)
                first_instance_find_com = False
            if self.com_port:
                break
            else:
                self.update_output("No COM Port found")
            time.sleep(1)

def process_excel_file(self, file_path):
        if self.com_port:
            self.start_serial(self.com_port, file_path)
        else:
            self.update_output("No COM port found. Exiting...")


    #Use powershell script to find COM port and strip result to use 'COMx'
def find_com_port(self):
        powershell_cmd = [
            'powershell.exe',
            '-Command',
            '''
            $5xxcomPort = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Caption -like '*USB Serial Port*' }
            $3xxcomPort = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Caption -like '*Prolific USB-to-Serial Comm Port*' }

            if ($5xxcomPort) {
                $5xxcomPortName = $5xxcomPort.Name
                Write-Output "USB Serial Port found on COM port: $5xxcomPortName"
            } elseif ($3xxcomPort) {
                $3xxcomPortName = $3xxcomPort.Name
                Write-Output "USB Serial Port found on COM port: $3xxcomPortName"
            } else {
                Write-Output "USB Serial Port not found"
            }
            '''
        ]
        #Strip result logic
        try:
            result = subprocess.check_output(powershell_cmd, text=True)
            lines = result.strip().split('\n')
            for line in result.strip().split('\n'):
                if "USB Serial Port" in line:
                    com_port_index = line.find("(COM")
                    if com_port_index != -1:
                        com_port = line[com_port_index:].strip("() ")
                        return com_port
            else:
                return None
        except subprocess.CalledProcessError as e:
            print(f"Error: {e}")
            return None

    #Look for columns MAC, AP Name, and AP Group
    #Look for MAC row, give AP {name} {group} in MAC row, save 
    #Give the user the results
def process_mac(self, mac_to_match, com_port, wb):
        self.excel_button.configure(state=tk.DISABLED) ## Disable user control to prevent unresponsive window
        self.restart_button.configure(state=tk.DISABLED)

        sheet = wb.active
        headers = {cell.value: cell.column for cell in sheet[1]}

        #Edit column names here
        mac_column_index = headers.get("MAC")
        name_column_index = headers.get("AP Name")
        group_column_index = headers.get("AP Group")

        if mac_column_index is None or name_column_index is None or group_column_index is None:
            self.update_output("Error: Required column not found in Excel file.")
            return

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[mac_column_index - 1] == mac_to_match:
                name = row[name_column_index - 1]
                group = row[group_column_index - 1]

                with serial.Serial(com_port, baudrate=9600, timeout=1) as ser:
                    time.sleep(1)

                    ser.write(f"set name {name}\r\n".encode())
                    time.sleep(1)

                    ser.write(f"set group {group}\r\n".encode())
                    time.sleep(1)

                    ser.write(b"saveenv \r\n")
                    time.sleep(2)

                    ser.write(b"printenv \r\n")
                    time.sleep(1)

                    #Give provided name/group/mac : source : AP
                    response = ser.read_all().decode()

                    provided_name = response.find("name=")
                    name_endIndex = response.find("\n", provided_name)
                    actualName = response[provided_name:name_endIndex]
                    
                    provided_group = response.find("group=")
                    group_endIndex = response.find("\n", provided_group)
                    actualGroup = response[provided_group:group_endIndex]
                    
                    provided_mac = response.find("ethaddr=")
                    mac_endIndex = response.find("\n", provided_mac)
                    actualMac = response[provided_mac:mac_endIndex]
                    
                    #Pass name, group, MAC and activate buttons
                    self.update_output("Extracted name: " + actualName + "Extracted group" + actualGroup + "Extracted Mac" + actualMac)
                    self.excel_button.configure(state=tk.ACTIVE)
                    self.restart_button.configure(state=tk.ACTIVE)
                break
        else:
            self.update_output("MAC address not found in the Excel file.")
            self.excel_button.configure(state=tk.ACTIVE)
            self.restart_button.configure(state=tk.ACTIVE)


    #Load Excel file
    #Open COM
    #Stop the autoboot
    #Find MAC and pass 
def start_serial(self, com_port, file_path):
        self.excel_button.configure(state=tk.DISABLED) ## Disable user control to prevent unresponsive window
        self.restart_button.configure(state=tk.DISABLED)
        
        try:
            wb = load_workbook(file_path, read_only=True)
            self.file_path = file_path  # Store the file path
            self.update_output("Selected Excel file: " + file_path)
            self.update_idletasks()  # Force GUI update

            with serial.Serial(com_port, baudrate=9600, timeout=1) as ser:
                no_data_prompt_printed = False

                while True:
                    data = ser.readline()

                    if b"Hit <Enter>" in data:
                        self.update_output(data.decode())
                        ser.write(b"\r\n")
                        break
                    elif not data:
                        if not no_data_prompt_printed:
                            self.update_output("No data received from serial port.")
                            self.update_idletasks()  # Force GUI update
                            no_data_prompt_printed = True
                    else:
                        no_data_prompt_printed = False
                        ser.write(b"\r\n")

                self.update_output("Sending print command...")
                self.update_idletasks()  # Force GUI update

                ser.write(b"printenv \r\n")
                time.sleep(1)

                response = ser.read_all().decode()

                ser.close()  #This will not work without this line. It will 're-open' the already being used COM port.

                mac_start_index = response.find("ethaddr=")
                if mac_start_index != -1:
                    mac_end_index = response.find("\n", mac_start_index)
                    mac_address_line = response[mac_start_index:mac_end_index].strip()
                    mac_address = mac_address_line.split("=")[1].replace(":", "").upper()

                    self.update_output("Extracted MAC Address:" + mac_address)
                    self.process_mac(mac_address, com_port, wb)
                    self.excel_button.configure(state=tk.ACTIVE)
                    self.restart_button.configure(state=tk.ACTIVE)
                else:
                    self.update_output("MAC address not found in the printenv response.")
                    self.excel_button.configure(state=tk.ACTIVE)
                    self.restart_button.configure(state=tk.ACTIVE)

        except Exception as e:
            self.update_output(f"Error reading Excel file: {str(e)}")
            self.excel_button.configure(state=tk.ACTIVE)
            self.restart_button.configure(state=tk.ACTIVE)


if __name__ == "__main__":
        gui = APNamerGUI()
        gui.mainloop()
