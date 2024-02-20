import customtkinter as ctk
import tkinter as tk
import threading

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

