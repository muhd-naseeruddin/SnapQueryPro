import tkinter as tk
from tkinter import filedialog, messagebox
# import tkinter.font as tkFont
from ttkbootstrap.constants import *
import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledFrame
import datetime, requests, os, threading
from loguru import logger
import snapquery
import time
from seleniumbase import Driver
from pandas import read_excel
from PIL import Image, ImageTk


class GUI():
    def __init__(self, master, input_field, input_first_name, input_last_name, input_company_name):
        ##Singular variable
        self.input_field = input_field
        self.input_first_name = input_first_name
        self.input_last_name = input_last_name
        self.input_company_name = input_company_name
        self.driver_instances = {}
        self.errors = []

        ##Multiple variable
        self.multi_input_field = input_field
        self.multi_input_first_name = input_first_name
        self.multi_input_last_name = input_last_name
        self.multi_input_company_name = input_company_name
        self.multi_driver_instances = {}
        self.multi_errors = []
        self.processed_name = {}
        self.counter = 0

        self.lock = threading.Lock()
        self.master = master
        self.master.title("SnapQuery Pro TH")
        self.master.iconbitmap("Images\\screenshot_icon.ico")

        self.master.geometry("740x680")
        self.gray = "#a6a6a6"
        self.beige = "#836c58"

        self.first_name_placeholder = "Enter first name here..."
        self.last_name_placeholder = "Enter last name here..."
        self.company_name_placeholder = "Enter company name here..."

        # Create a main notebook
        self.main_tab = tb.Notebook(self.master, bootstyle="default")
        self.main_tab.pack(expand=True, fill='both', padx=10, pady=10)

        # Create singular frame
        self.singular_tab = tb.Frame(self.main_tab)
        self.singular_tab.pack_propagate(True)
        self.main_tab.add(self.singular_tab, text="Single Search")
        self.upper_row = tb.Labelframe(self.singular_tab, border=False)
        self.upper_row.grid(row=0, column=0)

        # Widgets for singular tab
        # First name widget
        self.individual_label = tb.Labelframe(self.upper_row, text="Individual", width=460, height=150)
        self.individual_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.individual_label.grid_propagate(False)
        self.first_name_label = tb.Label(self.individual_label, text="First Name:", bootstyle="default", anchor="center")
        self.first_name_label.grid(row=0, column=0, padx=10, pady=(15,0), sticky="nsew")
        self.first_name_entry = tb.Entry(self.individual_label, width=40, foreground=self.gray)
        self.first_name_entry.placeholder = self.first_name_placeholder
        self.first_name_entry.insert(0, self.first_name_placeholder)
        self.first_name_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.first_name_entry.bind("<FocusOut>", self.on_entry_focus_out)
        self.first_name_entry.grid(row=0, column=1, padx=10, pady=(20,5))
        
        # Last name widget
        self.last_name_label = tb.Label(self.individual_label, text="Last Name:", bootstyle="default", anchor="center")
        self.last_name_label.grid(row=1, column=0, padx=10, pady=(0,20), sticky="nsew")
        self.last_name_entry = tb.Entry(self.individual_label, width=40, foreground=self.gray)
        self.last_name_entry.placeholder = self.last_name_placeholder
        self.last_name_entry.insert(0, self.last_name_placeholder)
        self.last_name_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.last_name_entry.bind("<FocusOut>", self.on_entry_focus_out)
        self.last_name_entry.grid(row=1, column=1, padx=10, pady=(0,20))

        # OR widget
        self.or_label = tb.Label(self.upper_row, text="OR", anchor="center")
        self.or_label.grid(row=1, column=0, sticky="nswe")

        # Company name widget
        self.company_label = tb.Labelframe(self.upper_row, text="Company", width=460, height=80)
        self.company_label.grid(row=2, column=0, padx=10, pady=(0,10), sticky='w')
        self.company_label.grid_propagate(False)
        self.company_name_label = tb.Label(self.company_label, text="Company Name:", bootstyle="default", anchor="center")
        self.company_name_label.grid(row=0, column=0, padx=10, pady=10)
        self.company_name_entry = tb.Entry(self.company_label, foreground=self.gray, width=35)
        self.company_name_entry.placeholder = self.company_name_placeholder
        self.company_name_entry.insert(0, self.company_name_placeholder)
        self.company_name_entry.bind("<FocusIn>", self.on_entry_focus_in)
        self.company_name_entry.bind("<FocusOut>", self.on_entry_focus_out)
        self.company_name_entry.grid(row=0, column=1, padx=12, pady=10)

        #Website status widget
        self.website_status_frame = tb.Labelframe(self.upper_row, text="Web Status", bootstyle="default", width=200, height=150)
        self.website_status_frame.grid(row=0, column=1, padx=20, pady=10, sticky="w")
        self.website_status_frame.grid_propagate(False)

        #Refresh button widget
        self.refresh_button = tb.Button(self.website_status_frame, text="Refresh", bootstyle="default", command=self.refresh_website_status, width=19)
        self.refresh_button.grid(row=0, column=0, columnspan=2, padx=10, pady=(10,0), sticky="nsew")
        self.website_ofac = tb.Label(self.website_status_frame, text="OFAC", anchor="center")
        self.website_ofac.grid(row=1, column=0, padx=10, pady=(10,0), sticky="w")
        self.ofac_connection_status = tb.Label(self.website_status_frame, text="Refresh to check", foreground="grey", anchor="center")
        self.ofac_connection_status.grid(row=1, column=1, padx=10, pady=(10,0), sticky="w")
        self.website_oic = tb.Label(self.website_status_frame, text="OIC", anchor="center")
        self.website_oic.grid(row=2, column=0, padx=10, pady=0, sticky="w")
        self.oic_connection_status = tb.Label(self.website_status_frame, text="Refresh to check", foreground="grey", anchor="center")
        self.oic_connection_status.grid(row=2, column=1, padx=10, pady=0, sticky="w")
        self.last_check_label = tb.Label(self.website_status_frame, text="Last checked: ", anchor="center")
        self.last_check_label.grid(row=3, column=0, columnspan=2, padx=10, pady=(0,10), sticky='w')

        #Website widget
        self.website_frame = tb.Labelframe(self.upper_row, text="Websites", bootstyle="default")
        self.website_frame.grid(row=2, column=1, padx=20, pady=(0,10), sticky="w")
        self.function_checkboxes = {}
        self.function = {
            "ofac": "OFAC Sanctions Search",
            "oic": "OIC e-Services"
        }
        for i, (func_name, func_desc) in enumerate(self.function.items()):
            self.function_checkboxes[func_name] = tb.BooleanVar(value=True)
            checkbox = tb.Checkbutton(self.website_frame, text=func_desc, variable=self.function_checkboxes[func_name], command=lambda name=func_name: self.update_checked_box(name))
            checkbox.grid(row=i, column=0, sticky="w", padx=10, pady=5)

        #Browse widget
        self.browse_frame = tb.Labelframe(self.upper_row, text="Folder")
        self.browse_frame.grid(row=3, column=0, columnspan=1, padx=10, pady=10, sticky="nsew")
        self.browse_label = tb.Label(self.browse_frame, text="Save location:")
        self.browse_label.grid(row=0, column=0, sticky="nsew", padx=(10,0))
        self.initial_folder_path = os.path.expanduser("~\\Desktop")
        self.folder_var = tb.StringVar()
        self.folder_var.set(self.initial_folder_path)
        self.folder_entry = tb.Entry(self.browse_frame, width=29, textvariable=self.folder_var)
        self.folder_entry.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.browse_button = tb.Button(self.browse_frame, text="Browse", command=lambda exec_type = "singular": self.browse_folder(exec_type))
        self.browse_button.grid(row=0, column=3, padx=(0,10), pady=10, sticky="nsew")

        #Search type
        self.singular_search_type_frame = tb.Labelframe(self.upper_row, text="Search type")
        self.singular_search_type_frame.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        self.singular_search = tb.IntVar(value=2)
        self.singular_radio_button_dictionary = {"Normal": 1,
                                                 "Fast": 2}
        for i,(text, value) in enumerate(self.singular_radio_button_dictionary.items()):
            radio_button = tb.Radiobutton(self.singular_search_type_frame, text=text, variable=self.singular_search, value=value)
            radio_button.grid(row=0, column=i, padx=10, pady=15)

        #Execute button widget
        self.execute_button = tb.Button(self.upper_row, text="Execute", width=19, command=self.execute_selenium)
        self.execute_button.grid(row=4, column=0, columnspan=2, padx=10, pady=(5,0))

        # Screenshot status widget
        self.screenshot_status = tb.Labelframe(self.singular_tab, text="Status")
        self.screenshot_status.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky=E+W+N+S)
        self.singular_tab.rowconfigure(2, weight=1)
        self.singular_tab.columnconfigure(0, weight=1)
        self.screenshot_status.columnconfigure(1, weight=1)
        self.screenshot_status.rowconfigure(0, weight=1)
        self.screenshot_label_frame = tb.Labelframe(self.screenshot_status, border=False, width=260, height=200)
        self.screenshot_label_frame.grid(row=0, column=0, padx=5, pady=10, sticky="nw")
        self.screenshot_label_frame.grid_propagate(False)
        self.ofac_screenshot_label = tb.Label(self.screenshot_label_frame, text="OFAC", foreground="#007FFF", font=("Helvetica", 12, "bold"), anchor="center", width=10)
        self.ofac_screenshot_label.grid(row=0, column=0, padx=10, pady=10, sticky=N+S+E+W)
        self.ofac_select_label = tb.Label(self.screenshot_label_frame, text="Selected", foreground="#007FFF", font=("Helvetica", 12))
        self.ofac_select_label.grid(row=1, column=0, padx=10, pady=10)
        self.oic_screenshot_label = tb.Label(self.screenshot_label_frame, text="OIC", foreground="#007FFF", font=("Helvetica", 12, "bold"), anchor="center", width=10)
        self.oic_screenshot_label.grid(row=0, column=1, padx=10, pady=10, sticky=N+S+E+W)
        self.oic_select_label = tb.Label(self.screenshot_label_frame, text="Selected", foreground="#007FFF", font=("Helvetica", 12))
        self.oic_select_label.grid(row=1, column=1, padx=10, pady=10)

        # Status message widget
        self.status_text = tb.ScrolledText(self.screenshot_status, width=56, wrap="none")
        self.status_text.grid(row=0, column=1, padx=(10,0), pady=10, sticky=E+W+N+S)

        # Create multiple frame tab
        self.multiple_tab = tb.Frame(self.main_tab)
        self.multiple_tab.pack_propagate(True)
        self.main_tab.add(self.multiple_tab, text="Multiple Search")

        # Create transparent frame to hold widget together
        self.first_row = tb.Labelframe(self.multiple_tab, border=False)
        self.first_row.grid(row=0, column=0)

        # Create widget under first row label frame
        # Create step one widget
        self.step_one = tb.Labelframe(self.first_row, text="Step 1", width=420, height=90)
        self.step_one.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")
        self.step_one.grid_propagate(False)
        self.step_one_text = "Select your file containing list of subject names:"
        self.step_one_instruction = tb.Label(self.step_one, text=self.step_one_text, anchor="w")
        self.step_one_instruction.grid(row=0, column=0, padx=5, pady=(0,5), sticky="w")
        self.source_file_var = tb.StringVar()
        self.source_file_var.set(self.initial_folder_path)
        self.source_file_entry = tb.Entry(self.step_one, width=39, textvariable=self.source_file_var)
        self.source_file_entry.grid(row=1, column=0, columnspan=1, padx=5, pady=(0,5), sticky="w")
        self.source_browse_button = tb.Button(self.step_one, text="Browse", command=self.source_browse_folder)
        self.source_browse_button.grid(row=1, column=1, padx=5, pady=(0,5), sticky="w")

        # Create step two widget
        self.step_two = tb.Labelframe(self.first_row, text="Step 2", width=420, height=90)
        self.step_two.grid(row=1, column=0, padx=10, pady=(5,5), sticky="w")
        self.step_two.grid_propagate(False)
        # self.step_two_text = "Which sheet is your list of subject names in (e.g Sheet1):"
        self.step_two_sheet = tb.Label(self.step_two, text="Sheet:")
        self.step_two_sheet.grid(row=0, column=0, padx=5, pady=(0,5), sticky="w")
        self.sheet_name = tb.StringVar()
        self.sheet_name.set("Sheet1")
        self.sheet_name_entry = tb.Entry(self.step_two, width=15, textvariable=self.sheet_name)
        self.sheet_name_entry.grid(row=1, column=0, padx=5, pady=(0,5), sticky="w")
        self.step_two_column = tb.Label(self.step_two, text="Column:")
        self.step_two_column.grid(row=0, column=1, padx=5, pady=(0,5), sticky="w")
        self.column_name = tb.StringVar()
        self.column_name.set("Names")
        self.column_name_entry = tb.Entry(self.step_two, width=15, textvariable=self.column_name)
        self.column_name_entry.grid(row=1, column=1, padx=5, pady=(0,5), sticky="w")
        self.step_two_row = tb.Label(self.step_two, text="Row start:")
        self.step_two_row.grid(row=0, column=2, padx=5, pady=(0,5), sticky="w")
        self.row_name = tb.IntVar(value=1)
        self.spin_box = tb.Spinbox(self.step_two, width=6, from_=1, to=1048576, textvariable=self.row_name)
        self.spin_box.grid(row=1, column=2, padx=5, pady=(0,5), sticky="w")

        #Create step three widget
        self.step_three = tb.Labelframe(self.first_row, text="Step 3", width=420, height=60)
        self.step_three.grid(row=2, column=0, padx=10, pady=(5,5), sticky="w")
        self.step_three.grid_propagate(False)
        self.multi_browse_label = tb.Label(self.step_three, text="Save path:")
        self.multi_browse_label.grid(row=0, column=0, sticky="w", padx=5, pady=(0,5))
        self.multi_folder_var = tb.StringVar()
        self.multi_folder_var.set(self.initial_folder_path)
        self.multi_folder_entry = tb.Entry(self.step_three, width=29, textvariable=self.multi_folder_var)
        self.multi_folder_entry.grid(row=0, column=1, columnspan=2, padx=(0,5), pady=(0,10), sticky="nsew")
        self.multi_browse_button = tb.Button(self.step_three, text="Browse", command=lambda exec_type = "multi": self.browse_folder(exec_type))
        self.multi_browse_button.grid(row=0, column=3, padx=(0,10), pady=(0,10), sticky="nsew")

        #Create step four widget
        self.step_four = tb.Labelframe(self.first_row, text="Step 4", width=210, height=90)
        self.step_four.grid(row=0, column=1, padx=10, pady=(10,5), sticky="w")
        self.step_four.grid_propagate(False)
        self.step_four_text = "Timeout to wait before error:"
        self.step_four_instruction = tb.Label(self.step_four, text=self.step_four_text)
        self.step_four_instruction.grid(row=0, column=0, columnspan=2, padx=5, pady=(0,5), sticky="w")
        self.timeout = tb.IntVar(value=120)
        self.spin_box_timeout = tb.Spinbox(self.step_four,increment=10, from_=1, to=3600, textvariable=self.timeout)
        self.spin_box_timeout.grid(row=1, column=0, padx=5, pady=(0,5), sticky="w")

        #Create step five widget
        self.step_five = tb.Labelframe(self.first_row, text="Step 5", width=210, height=90)
        self.step_five.grid(row=1, column=1, padx=10, pady=(5,5), sticky="w")
        self.step_five.grid_propagate(False)
        self.step_five_checkbox = {}
        self.step_five_function = {
            "ofac": "OFAC Sanctions Search",
            "oic": "OIC e-Services"
        }
        for i, (func_name, func_desc) in enumerate(self.step_five_function.items()):
            self.step_five_checkbox[func_name] = tb.BooleanVar(value=True)
            checkbox = tb.Checkbutton(self.step_five, text=func_desc, variable=self.step_five_checkbox[func_name])
            checkbox.grid(row=i, column=0, sticky="w", padx=5, pady=6)

        #Search type widget
        self.multiple_search_type_frame = tb.Labelframe(self.first_row, text="Search type")
        self.multiple_search_type_frame.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
        self.multiple_search_type = tb.IntVar(value=2)
        self.multiple_radio_button_dictionary = {"Normal": 1,
                                                 "Fast": 2}
        for i,(text, value) in enumerate(self.multiple_radio_button_dictionary.items()):
            multiple_radio_button = tb.Radiobutton(self.multiple_search_type_frame, text=text, variable=self.multiple_search_type, value=value)
            multiple_radio_button.grid(row=0, column=i, padx=10, pady=10)

        #Multiple tab execute button widget
        # self.multi_tab_execute_frame = tb.Labelframe(self.first_row, text="", width=200, height=60, border=False)
        # self.multi_tab_execute_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=0, sticky="nsew")
        # self.multi_tab_execute_frame.grid_propagate(False)
        self.multi_execute_button = tb.Button(self.first_row, text="Execute", width=19, command=self.execute_selenium_multi)
        self.multi_execute_button.grid(row=3, column=0, columnspan=2, padx=10, pady=(0,5))

        #Mutli status message
        self.multi_status = tb.Labelframe(self.multiple_tab, text="Status")
        self.multi_status.grid(row=1, column=0, padx=10, pady=(5,10), sticky=N+S+E+W)
        self.multiple_tab.rowconfigure(1, weight=1)
        self.multiple_tab.columnconfigure(0, weight=1)
        self.multi_status.columnconfigure((0,1), weight=1)
        self.multi_status.rowconfigure((1,3), weight=1)
        self.multi_ofac_text = tb.Label(self.multi_status, text="Incomplete OFAC screenshot")
        self.multi_ofac_text.grid(row=0, column=0, padx=(10,0), sticky="w")
        self.multi_ofac_fail_status = tb.ScrolledText(self.multi_status, wrap="none")
        self.multi_ofac_fail_status.grid(row=1, column=0, padx=(10,0), sticky="nsew")
        self.multi_oic_text = tb.Label(self.multi_status, text="Incomplete OIC screenshot")
        self.multi_oic_text.grid(row=0, column=1, padx=(10,0), sticky="w")
        self.multi_oic_fail_status = tb.ScrolledText(self.multi_status, wrap="none")
        self.multi_oic_fail_status.grid(row=1, column=1, padx=10, sticky="nsew")
        self.multi_success_text = tb.Label(self.multi_status, text="Completed Screenshot")
        self.multi_success_text.grid(row=2, column=0, padx=(10,0), pady=(10,0), sticky="w")
        self.multi_success_status = tb.ScrolledText(self.multi_status, wrap="none")
        self.multi_success_status.grid(row=3, column=0, columnspan=2, padx=10, pady=(0,10), sticky="nsew")

        # Create help tab
        self.help_tab = tb.Frame(self.main_tab)
        self.help_tab.pack_propagate(True)
        self.main_tab.add(self.help_tab, text="Help")

        #Create scrolled frame
        self.help_frame = ScrolledFrame(self.help_tab, autohide=True)
        self.help_frame.pack(padx=10, pady=10, fill="both", expand=True, anchor="center")
        self.basic_guide = tb.Label(self.help_frame, text="Basic Guide", font=("Helvetica", 14, "bold"))
        self.basic_guide.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.single_search = tb.Label(self.help_frame, text="Single Search Instructions:", font=("Helvetica", 12, "bold"))
        self.single_search.grid(row=1, column=0, padx=5, pady=(5,2), sticky="w")
        self.single_search_name = tb.Label(self.help_frame, text="1. Fill in either an individual name or company name only.")
        self.single_search_name.grid(row=2, column=0, padx=5, pady=2, sticky="w")
        self.single_search_folder = tb.Label(self.help_frame, text="2. Click 'Browse' button and select your preferred save location.\n\
    Screenshot will be saved in the save location.")
        self.single_search_folder.grid(row=3, column=0, padx=5, pady=2, sticky="w")
        self.single_search_web_refresh = tb.Label(self.help_frame, text="3. [OPTIONAL] If you wish to check both OFAC and OIC connection,\n\
    click 'Refresh' button in the Web Status frame. After few seconds,\n\
    both website status will be shown either 'Online' or 'Offline'.")
        self.single_search_web_refresh.grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.single_search_checkbox = tb.Label(self.help_frame, text="4. Select preferred website(s) for SnapQuery to perform screenshot.")
        self.single_search_checkbox.grid(row=5, column=0, padx=5, pady=2, sticky="w")
        self.single_search_execute = tb.Label(self.help_frame, text="5. Once you have filled all the required informations, click 'Execute'.\n\
    Note: For first execution, please wait for a few seconds as the apps\n\
    will have a delay/lag on the interface.")
        self.single_search_execute.grid(row=6, column=0, padx=5, pady=2, sticky="w")
        self.single_search_status = tb.Label(self.help_frame, text="6. 'Status' shows the progress of the screenshot.\n\
    Once screenshot completed, an alert window will pop up.")
        self.single_search_status.grid(row=7, column=0, padx=5, pady=2, sticky="w")

        self.multiple_search = tb.Label(self.help_frame, text="Multiple Search Instructions:", font=("Helvetica", 12, "bold"))
        self.multiple_search.grid(row=8, column=0, padx=5, pady=(8,2), sticky="w")
        self.multiple_search_source_file = tb.Label(self.help_frame, text="1. Click 'Browse' button to select your excel file containing name list.")
        self.multiple_search_source_file.grid(row=9, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_step_two = tb.Label(self.help_frame, text="2. Enter the 'Sheet' name, 'Column' name, and starting 'Row' number.")
        self.multiple_search_step_two.grid(row=10, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_step_two_extended = tb.Label(self.help_frame, text="[IMPORTANT] Please separate individual and company name into their own column", font=("Helvetica", 10, "bold"), foreground="red")
        self.multiple_search_step_two_extended.grid(row=11, column=0, padx=5, pady=(0,1), sticky="w")
        self.multiple_search_step_two_example = tb.Label(self.help_frame, text="Example:\n\
    Sheet name: Sheet1\n\
    Column name: Company\n\
    Starting row number: 2")
        self.multiple_search_step_two_example.grid(row=12, column=0, padx=5, pady=(1,2), sticky="w")
        self.multi_search_tutorial = Image.open("Images\\multi_search_tutor.png")
        self.multi_search_tutorial = self.multi_search_tutorial.resize((500, 400), Image.Resampling.LANCZOS)
        self.multi_search_tutorial = ImageTk.PhotoImage(self.multi_search_tutorial)
        self.multi_search_label = tb.Label(self.help_frame, image=self.multi_search_tutorial, anchor="center")
        self.multi_search_label.grid(row=13, column=0, padx=5, pady=2)
        self.multiple_search_save_file = tb.Label(self.help_frame, text="3. Click 'Browse' button to select your screenshot save location.")
        self.multiple_search_save_file.grid(row=14, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_timeout = tb.Label(self.help_frame, text="4. Set timeout(in seconds) to wait for website and its component to load.\n\
    Note: This setting will be applied to 'Single Search' tab too.\n\
    Example: When you set 10 seconds, if the website took more than 10 seconds,\n\
                    an error will be received.")
        self.multiple_search_timeout.grid(row=15, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_checkbox = tb.Label(self.help_frame, text="5. Select preferred website(s) for SnapQuery to perform screenshot.")
        self.multiple_search_checkbox.grid(row=16, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_execute = tb.Label(self.help_frame, text="6. Once you have filled all the required informations, click 'Execute'.\n\
    Note: For first execution, please wait for a few seconds as the apps\n\
    will have a delay/lag on the interface.")
        self.multiple_search_execute.grid(row=17, column=0, padx=5, pady=2, sticky="w")
        self.multiple_search_status = tb.Label(self.help_frame, text="Note: If there is any error, the execution for the particular website will be stopped,\n\
    it will be shown in either OFAC or OIC Incomple Screenshot status box.\n\
    In the 'Status' box, it will show the progress of each name in the name list.\n\
    Once completed, a pop up window will appear to notify completion of automation.")
        self.multiple_search_status.grid(row=18, column=0, padx=5, pady=2, sticky="w")

        # Create right-click menu
        self.entry_menu = tb.Menu(master, tearoff=0)
        self.entry_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut_text)
        self.entry_menu.add_command(label="Copy", accelerator="Ctrl+C",  command=self.copy_text)
        self.entry_menu.add_command(label="Paste",  accelerator="Ctrl+V", command=self.paste_text)
        self.entry_menu.add_separator()
        self.entry_menu.add_command(label="Select all",  accelerator="Ctrl+A", command=self.select_all)

        # Bind right-click menu to entry widgets
        self.first_name_entry.bind("<Button-3>", self.show_entry_menu)
        self.last_name_entry.bind("<Button-3>", self.show_entry_menu)
        self.company_name_entry.bind("<Button-3>", self.show_entry_menu)
        self.folder_entry.bind("<Button-3>", self.show_entry_menu)
        # self.master.bind_class("Entry", "<Button-3>", self.show_entry_menu)

        #Logger setting
        today_date = datetime.datetime.now().strftime("%Y-%m-%d")
        logger_file_name = f"SnapQuery_{today_date}.log"
        logger.add(f"Logs\\{logger_file_name}", rotation="1 day", retention="7 days")

        # Close window
        self.master.protocol("WM_DELETE_WINDOW", self.close_window)


    def show_entry_menu(self, event):
        self.entry_menu.post(event.x_root, event.y_root)

    def cut_text(self):
        self.master.focus_get().event_generate("<<Cut>>")

    def copy_text(self):
        self.master.focus_get().event_generate("<<Copy>>")

    def paste_text(self):
        self.master.focus_get().event_generate("<<Paste>>")
        
    def select_all(self):
        focused_widget = self.master.focus_get()
        focused_widget.selection_range(0, tk.END)

    def browse_folder(self, exec_type):
        if exec_type == "singular":
            folder_path = filedialog.askdirectory(initialdir=self.folder_var.get())
            if folder_path:
                self.folder_var.set(folder_path)
        elif exec_type == "multi":
            folder_path = filedialog.askdirectory(initialdir=self.multi_folder_var.get())
            if folder_path:
                self.multi_folder_var.set(folder_path)

    def source_browse_folder(self):
        source_file = filedialog.askopenfilename()
        if source_file:
            self.source_file_var.set(source_file)

    def on_entry_focus_in(self, event):
        entry = event.widget
        placeholder = entry.placeholder
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.configure(foreground=self.beige)

    def on_entry_focus_out(self, event):
        entry = event.widget
        placeholder = entry.placeholder
        if not entry.get():
            entry.insert(0, placeholder)
            entry.configure(foreground=self.gray)

    def refresh_website_status(self):
        ofac_status = self.get_website_status("https://sanctionssearch.ofac.treas.gov/")
        oic_status = self.get_website_status("https://smart.oic.or.th/EService/Menu1")
        self.update_website_status(ofac_status, oic_status)
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.last_check_label.config(text="Last checked: " + current_time)

    def get_website_status(self, url):
        try:
            response = requests.get(url, verify=False)
            if response.status_code == 200:
                return "Online"
            else:
                return "Offline"
        except (requests.ConnectionError, requests.exceptions.HTTPError):
            return "Offline"
        
    def update_website_status(self, ofac_status, oic_status):
        self.ofac_connection_status.config(text=ofac_status, foreground="green" if ofac_status == "Online" else "red")
        self.oic_connection_status.config(text=oic_status, foreground="green" if oic_status == "Online" else "red")

    def update_checked_box(self, name):
        checkbox_value = self.function_checkboxes[name].get()
        if name == "ofac":
            self.ofac_screenshot_label.config(foreground="#007FFF" if checkbox_value else "gray")
            self.ofac_select_label.config(text="Selected" if checkbox_value else "", foreground="#007FFF" if checkbox_value else "gray")
        elif name == "oic":
            self.oic_screenshot_label.config(foreground="#007FFF" if checkbox_value else "gray")
            self.oic_select_label.config(text="Selected" if checkbox_value else "", foreground="#007FFF" if checkbox_value else "gray")

    def excel_to_list(self):
        source_file = str(self.source_file_entry.get().strip())
        sheet_name = str(self.sheet_name_entry.get().strip())
        column_name = str(self.column_name_entry.get().strip())
        try:
            if self.row_name.get() == None:
                header = 0
            else:
                header = self.row_name.get() - 1
            df = read_excel(source_file, sheet_name=sheet_name, header=header)
            name_list = df[column_name].to_list()
            return name_list
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}")
            raise e
        
    def update_multi_status(self, name):
        with self.lock:
            if all(self.processed_name[name].values()):
                self.status_message("multi_success", "success", f"[SUCCESS] {name} has completed screenshot")

    def execute_selenium_multi(self, event=None):
        source_file = str(self.source_file_entry.get().strip())
        sheet_name = str(self.sheet_name_entry.get().strip())
        column_name = str(self.column_name_entry.get().strip())
        multi_folder_entry = str(self.multi_folder_entry.get().strip())
        exec_type = "multi"
        multi_threads = []
        multi_start_event = threading.Event()
        self.counter = 0
        search_type = self.multiple_search_type.get()

        self.multi_success_status.configure(state="normal")
        self.multi_success_status.delete(1.0, tk.END)
        self.multi_success_status.configure(state="disable")
        self.multi_ofac_fail_status.configure(state="normal")
        self.multi_ofac_fail_status.delete(1.0, tk.END)
        self.multi_ofac_fail_status.configure(state="disable")
        self.multi_oic_fail_status.configure(state="normal")
        self.multi_oic_fail_status.delete(1.0, tk.END)
        self.multi_oic_fail_status.configure(state="disable")
        

        selected_websites = [key for key, value in self.step_five_checkbox.items() if value.get()]
        if not selected_websites:
            self.status_message("multi_success", "error", "[error] No websites selected for automation.")
            return
        try:
            if source_file != "" and sheet_name != "" and column_name != "" and multi_folder_entry != "":
                name_list = self.excel_to_list()
                for name in name_list:
                    name = name.strip()
                    snapquery.CreateFolder().create_folder_if_not_exists(multi_folder_entry, name)
                    self.processed_name[name] = {website: False for website in selected_websites}
                for website in selected_websites:
                    if website not in self.multi_driver_instances:
                        self.create_driver_instances(website, "multi")
                        self.status_message("multi_success", "info", f"Driver initialized for {website}")
                    if search_type == 2:
                        thread = threading.Thread(target=self.run_selenium_multi, args=(website, name_list, multi_folder_entry))
                        multi_threads.append(thread)
                        thread.start()
                    elif search_type == 1:
                        self.run_selenium_multi(website, name_list, multi_folder_entry)
                        self.status_message("multi_success", "info", f"Complete screenshot on {website} for all list.")

                multi_start_event.set()
                monitoring_thread = threading.Thread(target=self.monitor_threads_periodically, args=(multi_threads, exec_type))
                monitoring_thread.start()


                # Close unused driver instances
                for website in list(self.multi_driver_instances.keys()):
                    if website not in selected_websites:
                        driver_instance = self.multi_driver_instances[website]
                        driver_instance.quit()
                        del self.multi_driver_instances[website]
                        if website == "ofac":
                            self.multi_input_field = None 
                        if website == "oic":
                            self.multi_input_first_name = None
                            self.multi_input_last_name = None
                            self.multi_input_company_name = None
            else:
                self.status_message("multi_success", "error", "[ERROR] Please fill in all fields")
        except Exception as e:
            self.status_message("multi_success", "error",f"Error parsing Excel file: {e}")
            return

    def execute_selenium(self, event=None):
        threads = []
        start_event = threading.Event()
        exec_type = "singular"
        
        first_name = str(self.first_name_entry.get().strip())
        last_name = str(self.last_name_entry.get().strip())
        full_name = str(first_name + " " +last_name).strip()
        company_name = str(self.company_name_entry.get().strip())
        name = None
        first_name, last_name, company_name = self.check_name(first_name, last_name, company_name)
        full_name = str(first_name + " " +last_name).strip()
        folder = self.folder_entry.get().strip()
        search_type = self.singular_search.get()

        self.status_text.configure(state="normal")
        self.status_text.delete(1.0, tk.END)
        self.status_text.configure(state="disable")

        if (first_name == "") and \
           (last_name == "") and \
           (company_name == ""):
            self.status_message("status_text", "error", "[ERROR] Please enter candidate/subject name...")
            return
        elif first_name != "" and last_name != "" and company_name != "":
            self.status_message("status_text", "error", "[ERROR] Please enter either individual or company only")
            return
        elif (first_name != "" and last_name == "") and company_name != "":
            self.status_message("status_text", "error", "[ERROR] Please enter either individual or company only")
            return
        elif (first_name == "" and last_name != "") and company_name != "":
            self.status_message("status_text", "error", "[ERROR] Please enter either individual or company only")
            return
        if not folder:
            self.status_message("status_text", "error", "[ERROR] Please enter folder path...")
            return
        
        selected_functions = [func_name for func_name, var in self.function_checkboxes.items() if var.get()]
        for web in selected_functions:
            if web == "ofac":
                self.ofac_select_label.config(text="In Progress", foreground="orange")
            if web == "oic":
                self.oic_select_label.config(text="In Progress", foreground="orange")
        if (company_name == "") and (first_name == "" and last_name == ""):
            self.status_message("status_text", "error", "[ERROR] Please enter either company name or individual name only.")
            return
        elif (company_name != ""):
            snapquery.CreateFolder().create_folder_if_not_exists(folder, company_name)
            name = company_name
            self.status_message("status_text", "info", f"Company {name} folder is created.")
        elif (first_name != "") and \
             (last_name != ""):
            snapquery.CreateFolder().create_folder_if_not_exists(folder, full_name)
            name = full_name
            self.status_message("status_text", "info", f"Candidate {name} folder is created.")
        elif (first_name != "") and \
             (last_name == ""):
            snapquery.CreateFolder().create_folder_if_not_exists(folder, first_name)
            name = first_name
            self.status_message("status_text", "info", f"Candidate with first name {name} is created")
        elif (first_name == "") and \
             (last_name != ""):
            snapquery.CreateFolder().create_folder_if_not_exists(folder, last_name)
            name = last_name
            self.status_message("status_text", "info", f"Candidate with last name {name} is created")
        
        if not selected_functions:
            self.status_message("status_text", "error", "[error] No websites selected for automation.")
            return

        try:
            self.status_message("status_text", "info", f"Automation started for {name}")
            if search_type == 2:
                for selected_func in selected_functions:
                    if selected_func not in self.driver_instances:
                        driver_thread = threading.Thread(target=self.create_driver_instances, args=(selected_func, "singular"))
                        driver_thread.start()
                        driver_thread.join()
                        self.status_message("status_text", "info", f"Driver initialized for {selected_func}")
                    thread = threading.Thread(target=self.run_selenium, args=(selected_func, name, first_name, last_name, company_name, folder ))
                    threads.append(thread)
                    thread.start()
                    self.status_message("status_text", "info", f"Screenshot on {selected_func.upper()} is running...")

                start_event.set()

                monitoring_thread = threading.Thread(target=self.monitor_threads_periodically, args=(threads, exec_type))
                monitoring_thread.start()

                # Close unused driver instances
                for website in list(self.driver_instances.keys()):
                    if website not in selected_functions:
                        driver_instance = self.driver_instances[website]
                        driver_instance.quit()
                        del self.driver_instances[website]
                        if website == "ofac":
                            self.input_field = None 
                        if website == "oic":
                            self.input_first_name = None
                            self.input_last_name = None
                            self.input_company_name = None

            elif search_type == 1:
                for selected_func in selected_functions:
                    if selected_func not in self.driver_instances:
                        self.create_driver_instances(selected_func, "singular")
                        self.status_message("status_text", "info", f"Driver initialized for {selected_func}")
                    self.run_selenium(selected_func, name, first_name, last_name, company_name, folder)
                for website in list(self.driver_instances.keys()):
                    if website not in selected_functions:
                        driver_instance = self.driver_instances[website]
                        driver_instance.quit()
                        del self.driver_instances[website]
                        if website == "ofac":
                            self.input_field = None 
                        if website == "oic":
                            self.input_company_name = None
                            self.input_first_name = None
                            self.input_last_name = None
                self.monitor_threads_periodically(threads, exec_type)


        except Exception as e:
            self.status_message("status_text", "error", f"[ERROR] Error during process: {e}")
    
    def run_selenium_multi(self, website, name_list, multi_folder_entry , multi_start_event = None):
        timeout = self.timeout.get()
        for name in name_list:
            name = name.strip()
            try:
                self.status_message("multi_success", "info", f"[INFO] {name} is being searched on {website}.")
                with self.lock:
                    driver = self.multi_driver_instances[website]
                if multi_start_event:
                    multi_start_event.wait()
                if website == "ofac":
                    chev_ofac = snapquery.OFACAutoScreenshot(name, multi_folder_entry, timeout)
                    self.multi_input_field = chev_ofac.search_candidate(driver, self.multi_input_field)
                    self.processed_name[name][website] = True
                elif website == "oic":
                    if self.column_name_entry.get().lower().strip() == "company":
                        chev_oic = snapquery.OICAutoScreenshotV2(name, multi_folder_entry, timeout)
                        self.multi_input_company_name = chev_oic.search_candidate(driver, self.multi_input_company_name)
                        self.processed_name[name][website] = True
                    else:
                        try:
                            first_name, last_name = name.split(maxsplit=1)
                        except ValueError:
                            first_name = name
                            last_name = ""
                        chev_oic = snapquery.OICAutoScreenshotV1(first_name, last_name, multi_folder_entry, timeout)
                        self.multi_input_first_name, self.multi_input_last_name = chev_oic.search_candidate(driver, self.multi_input_first_name, self.multi_input_last_name)
                        self.processed_name[name][website] = True
            except Exception as e:
                if website == "ofac":
                    self.status_message("multi_ofac", "error", f"[ERROR] Automation on {name} with error: {e}")
                    self.processed_name[name][website] = False
                elif website == "oic":
                    self.status_message("multi_oic", "error", f"[ERROR] Automation on {name} with error: {e}")
                    self.processed_name[name][website] = False
                self.multi_errors.append(e)
                raise e
            
            if all(self.processed_name[name].values()):
                self.update_multi_status(name)
                self.counter += 1
                self.status_message("multi_success","info",f"[INFO] {self.counter}/{len(name_list)} subjects have completed")

    def run_selenium(self, selected_func, name, first_name, last_name, company_name, folder, start_event = None):
        timeout = self.timeout.get()
        try:
            with self.lock:
                driver = self.driver_instances[selected_func]
            if start_event:
                start_event.wait()
            if selected_func == "ofac":
                chev_ofac = snapquery.OFACAutoScreenshot(name, folder, timeout)
                self.input_field = chev_ofac.search_candidate(driver, self.input_field)
                self.ofac_select_label.config(text="Completed", foreground="green")
            elif selected_func == "oic":
                if (company_name != "" and company_name != self.company_name_placeholder):
                    self.status_message("status_text", "info",f"{company_name} is being searched.")
                    chev_oic = snapquery.OICAutoScreenshotV2(name, folder, timeout)
                    self.input_company_name = chev_oic.search_candidate(driver, self.input_company_name)
                else:
                    self.status_message("status_text", "info", f"{first_name} {last_name} is being searched")
                    chev_oic = snapquery.OICAutoScreenshotV1(first_name, last_name, folder, timeout)
                    self.input_first_name, self.input_last_name = chev_oic.search_candidate(driver, self.input_first_name, self.input_last_name)
                self.status_message("status_text", "info", f"Screenshot for {selected_func.upper()} has completed.")
                self.oic_select_label.config(text="Completed", foreground="green")
        except Exception as e:
            self.status_message("status_text", "info", f"Automation interrupted with error: {e}")
            self.errors.append(e)
            # Identify which instance encountered the error and handle it accordingly
            if selected_func == "ofac":
                self.handle_error(chev_ofac, e, "singular")
                self.ofac_select_label.config(text="Error", foreground="red")
            elif selected_func == "oic":
                self.handle_error(chev_oic, e, "singular")
                self.oic_select_label.config(text="Error", foreground="red")
            raise e

    def create_driver_instances(self, website, exec_type):
        driver = Driver(uc=True, headless=True, d_width=1920, d_height=1080)
        driver.set_window_size(1920, 1080)
        if exec_type == "singular":
            self.driver_instances[website] = driver
        elif exec_type == "multi":
            self.multi_driver_instances[website] = driver

    def status_message(self, status_type, type, message):
        time_stamp = datetime.datetime.now()
        formatted_time = time_stamp.strftime("%H:%M:%S")
        if status_type == "status_text":  
            if type == "error":
                self.status_text.configure(foreground="red")
                logger.error(message)
            elif type == "warning":
                self.status_text.configure(foreground="yellow")
                logger.warning(message)
            elif type == "info":
                logger.info(message)
                self.status_text.configure(foreground="grey")
            elif type == "success":
                logger.success(message)
                self.status_text.configure(foreground="green")
            self.status_text.configure(state="normal")
            self.status_text.insert(tk.END, f"[{formatted_time}]: {message}\n")
            self.status_text.configure(state="disabled")
            self.status_text.yview_pickplace("end")

        elif status_type == "multi_success":
            if type == "error":
                self.multi_success_status.configure(foreground="red")
                logger.error(message)
            elif type == "warning":
                self.multi_success_status.configure(foreground="yellow")
                logger.warning(message)
            elif type == "info":
                logger.info(message)
                self.multi_success_status.configure(foreground="grey")
            elif type == "success":
                logger.success(message)
                self.multi_success_status.configure(foreground="green")
            self.multi_success_status.configure(state="normal")
            self.multi_success_status.insert(tk.END, f"[{formatted_time}]: {message}\n")
            self.multi_success_status.configure(state="disabled")
            self.multi_success_status.yview_pickplace("end")

        elif status_type == "multi_ofac":
            if type == "error":
                self.multi_ofac_fail_status.configure(foreground="red")
                logger.error(message)
            self.multi_ofac_fail_status.configure(state="normal")
            self.multi_ofac_fail_status.insert(tk.END, f"[{formatted_time}]: {message}\n")
            self.multi_ofac_fail_status.configure(state="disabled")
            self.multi_ofac_fail_status.yview_pickplace("end")

        elif status_type == "multi_oic":
            if type == "error":
                self.multi_oic_fail_status.configure(foreground="red")
                logger.error(message)
            self.multi_oic_fail_status.configure(state="normal")
            self.multi_oic_fail_status.insert(tk.END, f"[{formatted_time}]: {message}\n")
            self.multi_oic_fail_status.configure(state="disabled")
            self.multi_oic_fail_status.yview_pickplace("end")

    def monitor_threads_periodically(self, threads, execution_type):
        
        if self.singular_search.get() == 2:
            while any(thread.is_alive() for thread in threads):
                time.sleep(1)  # Wait for 1 second before checking again
        elif self.multiple_search_type.get == 2:
            while any(thread.is_alive() for thread in threads):
                time.sleep(1)
            
        error_count = len(self.errors) 
        multi_error_count = len(self.multi_errors)
        time.sleep(1)
        if execution_type == "singular":
            self.status_message("status_text", "info", f"[INFO] Automation completed, {error_count} process(es) with errors.")
            messagebox.showinfo("[INFO]", f"Automation completed, {error_count} process(es) with errors.")
        elif execution_type == "multi":
            self.status_message("multi_success", "info", f"[INFO] Automation completed, {multi_error_count} process(es) with errors.")
            messagebox.showinfo("[INFO]", f"Automation completed, {multi_error_count} process(es) with errors")
        self.errors = []
        self.multi_errors = []


    def check_name(self, first_name, last_name, company_name):
        first_name = "" if first_name == self.first_name_placeholder else first_name
        last_name = "" if last_name == self.last_name_placeholder else last_name
        company_name = "" if company_name == self.company_name_placeholder else company_name
        return first_name, last_name, company_name
    
    def handle_error(self, instance, error, exec_type):
        if isinstance(instance, snapquery.OFACAutoScreenshot):
            if exec_type == "singular":
                self.status_message("status_text", "error", f"Error in OFACAutoScreenshot instance: {error}")
        elif isinstance(instance, snapquery.OICAutoScreenshotV1) or isinstance(instance, snapquery.OICAutoScreenshotV2):
            if exec_type == "singular":
                self.status_message("status_text", "error", f"Error in OICAutoScreenshot instance: {error}")

    def run(self):
        self.master.mainloop()

    def close_window(self):
        for driver_instance in self.driver_instances.values():
            driver_instance.quit()
        self.master.destroy()
        

        
def main():
    input_field = None
    input_first_name = None
    input_last_name = None
    input_company_name = None
    root = tb.Window(themename="morph")
    app = GUI(root, input_field=input_field, input_first_name=input_first_name, input_last_name=input_last_name, input_company_name=input_company_name)
    # root.mainloop()
    app.run()

if __name__ == "__main__":
    main()
