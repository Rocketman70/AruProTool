import serial
import time
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import subprocess
import threading

#
# TO DO: strip print env result for name, group and mac
#

# Define CustomTkinter class with necessary methods and attributes
class CustomTkinter:
    @staticmethod
    def set_appearance_mode(mode):
        CustomTkinter.appearance_mode = mode

    @staticmethod
    def set_default_color_theme(theme):
        CustomTkinter.default_color_theme = theme

    @staticmethod
    def get_color(key):
        if CustomTkinter.appearance_mode == "dark":
            if key == "background":
                return "#121212"
            elif key == "foreground":
                return "#FFFFFF"
            elif key == "accent":
                return "#2196F3"
        else:
            if key == "background":
                return "#FFFFFF"
            elif key == "foreground":
                return "#000000"
            elif key == "accent":
                return "#3f51b5"

    appearance_mode = "default"
    default_color_theme = "default"
#
#
#
#
#
#
#
#
#
#
#
#
# Define APNamerGUI class inheriting from ctk.CTk
class APNamerGUI(ctk.CTk):
    WIDTH = 800
    HEIGHT = 900

    def __init__(self):
        super().__init__()
        self.title("AP Namer")
        self.geometry(f"{APNamerGUI.WIDTH}x{APNamerGUI.HEIGHT}")
        CustomTkinter.set_appearance_mode("dark")
        CustomTkinter.set_default_color_theme("red")

        self.tabControl = ctk.CTkTabview(self)
        self.tabControl.pack(expand=1, fill="both")

        tab_1 = self.tabControl.add("Intro")
        tab_2 = self.tabControl.add("Excel/Output")

        self.com_port = None
        self.file_path = None  # Store the current Excel file path

        intro_text = """
        Welcome to AP Namer!

        1. Plug in serial cable to computer
        2. Select Excel file in next tab
        3. Plug serial cable into AP
        4. Follow prompts, plug AP into power
        5. Click restart when you want to provision the next AP
        
        *Once current AP is done provisioning you may select a new Excel file and use that


        P.S. Don't forget to feed your pet hamster while the program runs!🐹
        """

        intro_textbox = ctk.CTkTextbox(tab_1, font=("Lucida Console", 16), wrap=tk.WORD, height=20)
        intro_textbox.insert(tk.END, intro_text)
        intro_textbox.configure(state=tk.DISABLED)
        intro_textbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.excel_button = ctk.CTkButton(tab_2, text="Select Excel File", command=self.select_excel_file)
        self.excel_button.pack(pady=4)
        
        self.restart_button = ctk.CTkButton(tab_2, text="Restart", command=self.restart)
        self.restart_button.pack(pady=4)
        
        self.output_textbox = ctk.CTkTextbox(tab_2, font=("Lucida Console", 16), wrap=tk.WORD, height=20)
        self.output_textbox.configure(state=tk.NORMAL)
        self.output_textbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.com_port_thread = threading.Thread(target=self.check_com_ports)
        self.com_port_thread.daemon = True
        self.com_port_thread.start()
#
#
#
#
#
#
#
#
#
#
#
#

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

    def process_mac(self, mac_to_match, com_port, wb):
        self.excel_button.configure(state=tk.DISABLED) ## Disable user control to prevent unresponsive window
        self.restart_button.configure(state=tk.DISABLED)

        sheet = wb.active
        headers = {cell.value: cell.column for cell in sheet[1]}

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

                    response = ser.read_all().decode()
                    self.update_output(response)
                    self.excel_button.configure(state=tk.ACTIVE)
                    self.restart_button.configure(state=tk.ACTIVE)
                break
        else:
            self.update_output("MAC address not found in the Excel file.")
            self.excel_button.configure(state=tk.ACTIVE)
            self.restart_button.configure(state=tk.ACTIVE)

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

                ser.close()  # Close the serial port properly

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




# Check if the script is executed directly
if __name__ == "__main__":
    gui = APNamerGUI()
    gui.mainloop()
