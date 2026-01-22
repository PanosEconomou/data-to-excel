# Vassilis Economou  16/01/2025 v.02
#                   20/01/2026 v.2.1
#                   22/01/2026 v.2.2 (Added Language Toggle)

import openpyxl
from openpyxl import Workbook
import csv
import serial
import serial.tools.list_ports as list_ports
from datetime import datetime
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import warnings
import os
import requests  
from itertools import zip_longest


warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

class SerialDataLogger:
    def __init__(self, root):
        self.root = root
        # Αρχική Γλώσσα
        self.current_lang = "EL" 
        
        # Λεξικό Μεταφράσεων
        self.translations = {
            "EL": {
                "title": "Serial Data Logger [Βασίλης Οικονόμου v.2.2]",
                "instructions": "  Οδηγίες  ",
                "port_label": "Θα διαβάσω από τη Θύρα:",
                "refresh": "Ανανέωση",
                "baud_label": "...με ρυθμό (Baudrate):",
                "file_label": "Θα αποθηκεύσω στο αρχείο (.xlsx ή .csv):",
                "browse": "Επιλογή άλλου αρχείου",
                "col_titles": "Ονόματα στηλών στο .xlsx:",
                "thingspeak": "Αποστολή και στο ThingSpeak,  με API Key:",
                "ts_interval": " και συχνότητα αποστολής (σε δευτερόλεπτα):",
                "sampling": "Καθυστέρηση σε προβολή & απεικόνηση (ms):",
                "start": "Έναρξη",
                "stop": "Τερματισμός",
                "save": "Αποθήκευση στο αρχείο",
                "clear": "Καθαρισμός",
                "graph_win": "Διάγραμμα [με ανώτατο όριο τιμών (στον άξονα y):",
                "scroll": "Scrolling προς τα αριστερά",
                "points": "σημεία].",
                "log_win": "Kαταγραφή τιμών από τη θύρα",
                "copy": "Αντιγραφή",
                "export_csv": "Εξαγωγή επιλεγμένων σε .csv",
                "export_xlsx": "Εξαγωγή επιλεγμένων σε .xlsx",
                "lang_btn": "🇬🇧 English",
                "listbox_limit": "με όριο γραμμών προβολής:"
                
            },
            "EN": {
                "title": "Serial Data Logger [Vassilis Economou v.2.2]",
                "instructions": " Instructions ",
                "port_label": "Read from Port:",
                "refresh": "Refresh",
                "baud_label": "...with Baudrate:",
                "file_label": "Save to file (.xlsx or .csv):",
                "browse": "Browse File",
                "col_titles": "Column titles in .xlsx:",
                "thingspeak": "Send to ThingSpeak with API Key:",
                "ts_interval": " and interval (sεcond):",
                "sampling": "View & Plot delay (ms):",
                "start": "Start",
                "stop": "Stop",
                "save": "Save to File",
                "clear": "Clear",
                "graph_win": "Graph Window [Upper limit threshold:",
                "scroll": "Scroll to the left",
                "points": "points].",
                "log_win": "Serial port data log",
                "copy": "Copy",
                "export_csv": "Export selected to .csv",
                "export_xlsx": "Export selected to .xlsx",
                "lang_btn": "🇬🇷 Ελληνικά",
                "listbox_limit": "Listbox line limit:"
                

            }
        }

        self.root.title(self.translations[self.current_lang]["title"])

        # Αρχικοποίηση μεταβλητών
        self.serial_port = None
        self.baudrate = tk.IntVar(value=9600)
        self.max_val_limit = tk.IntVar(value=1024)
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "BO_SDL.xlsx"))
        self.times = []
        self.values = []
        self.data_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.sampling_rate = tk.IntVar(value=0)
        self.send_to_thingspeak = tk.BooleanVar(value=False)
        self.thingspeak_api_key = tk.StringVar(value="0J62FHGN0IS42VNQ")
        self.scroll_mode = tk.BooleanVar(value=False)
        self.scroll_window_size = tk.IntVar(value=500)
        self.actual_timestamps = []
        self.listbox_limit = tk.IntVar(value=80000) # Προεπιλεγμένο όριο έχει δοκιμαστεί 50000 γραμμές

        
        self.lines = [] 
        self.ts_interval = tk.IntVar(value=15)
        self.last_ts_send = datetime.min  
        
        self.create_widgets()

    def toggle_language(self):
        """Εναλλαγή μεταξύ Ελληνικών και Αγγλικών"""
        self.current_lang = "EN" if self.current_lang == "EL" else "EL"
        t = self.translations[self.current_lang]
        
        # Ενημέρωση κειμένων
        self.root.title(t["title"])
        self.title_label.config(text=t["title"])
        self.instr_btn.config(text=t["instructions"])
        self.lang_btn.config(text=t["lang_btn"])
        self.port_lbl.config(text=t["port_label"])
        self.refresh_btn.config(text=t["refresh"])
        self.baud_lbl.config(text=t["baud_label"])
        self.file_lbl.config(text=t["file_label"])
        self.browse_btn.config(text=t["browse"])
        self.col_titles_lbl.config(text=t["col_titles"])
        self.tspeak_chk.config(text=t["thingspeak"])
        self.sampling_lbl.config(text=t["sampling"])
        self.start_btn.config(text=t["start"])
        self.stop_btn.config(text=t["stop"])
        self.save_btn.config(text=t["save"])
        self.clear_btn.config(text=t["clear"])
        self.graph_win_lbl.config(text=t["graph_win"])
        self.scroll_chk.config(text=t["scroll"])
        self.points_lbl.config(text=t["points"])
        self.log_win_lbl.config(text=t["log_win"])
        self.listbox_lbl.config(text=" | " + t["listbox_limit"])
        self.ts_interval_lbl.config(text=t["ts_interval"])
       
        # Ενημέρωση Context Menu
        self.context_menu.entryconfigure(0, label=t["copy"])
        self.context_menu.entryconfigure(1, label=t["export_csv"])
        self.context_menu.entryconfigure(2, label=t["export_xlsx"])

    def create_widgets(self):
        t = self.translations[self.current_lang]
        
        # Header
        self.title_label = ttk.Label(self.root, text=t["title"], font=("Arial", 16, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Buttons Top Right
        btn_frame = ttk.Frame(self.root)
        btn_frame.grid(row=0, column=1, sticky="ne", pady=10, padx=5)
        
        self.lang_btn = ttk.Button(btn_frame, text=t["lang_btn"], command=self.toggle_language)
        self.lang_btn.pack(side=tk.RIGHT, padx=2)
        
        self.instr_btn = ttk.Button(btn_frame, text=t["instructions"], command=self.open_instructions_window)
        self.instr_btn.pack(side=tk.RIGHT, padx=2)
        
        # Connection Line
        conn_frame = ttk.Frame(self.root)
        conn_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        self.port_lbl = ttk.Label(conn_frame, text=t["port_label"])
        self.port_lbl.pack(side=tk.LEFT, padx=5)
        
      

        self.ports_combobox = ttk.Combobox(conn_frame, state="readonly", width=20)
        self.ports_combobox.pack(side=tk.LEFT, padx=5)
        self.refresh_ports()
        
        self.refresh_btn = ttk.Button(conn_frame, text=t["refresh"], command=self.refresh_ports)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        
        self.baud_lbl = ttk.Label(conn_frame, text=t["baud_label"])
        self.baud_lbl.pack(side=tk.LEFT, padx=(20, 5))
        
        baudrate_combobox = ttk.Combobox(conn_frame, textvariable=self.baudrate, state="readonly", width=10)
        baudrate_combobox["values"] = [9600, 19200, 38400, 57600, 115200]
        baudrate_combobox.pack(side=tk.LEFT, padx=5)

        # File Selection
        file_frame = ttk.Frame(self.root)
        file_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        
        self.file_lbl = ttk.Label(file_frame, text=t["file_label"])
        self.file_lbl.pack(side=tk.LEFT, padx=5)
        
        ttk.Entry(file_frame, width=60, textvariable=self.output_path).pack(side=tk.LEFT, padx=5)
        self.browse_btn = ttk.Button(file_frame, text=t["browse"], command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=5)

        # Column Titles
        self.col_titles_lbl = ttk.Label(self.root, text=t["col_titles"])
        self.col_titles_lbl.grid(row=4, column=0, padx=5, pady=3, sticky="w")
        
        fields_frame = ttk.Frame(self.root)
        fields_frame.grid(row=4, column=1, columnspan=1, padx=5, pady=(10, 3), sticky="ew")
        self.extra_text_vars = [tk.StringVar(value=f"Col{i+1}") for i in range(8)]
        for i in range(8):
            ttk.Entry(fields_frame, textvariable=self.extra_text_vars[i], width=7).grid(row=0, column=i, padx=2, sticky="ew")

        # ThingSpeak
        self.tspeak_chk = ttk.Checkbutton(self.root, text=t["thingspeak"], variable=self.send_to_thingspeak)
        self.tspeak_chk.grid(row=5, column=0, padx=5, pady=3, sticky="w")

        ts_frame = ttk.Frame(self.root)
        ts_frame.grid(row=5, column=1, padx=5, pady=3, sticky="w")
        ttk.Entry(ts_frame, textvariable=self.thingspeak_api_key, width=20).pack(side=tk.LEFT)
        #ttk.Label(ts_frame, text=" (sec):").pack(side=tk.LEFT)
        #ttk.Entry(ts_frame, textvariable=self.ts_interval, width=5).pack(side=tk.LEFT, padx=5)
        # Μέσα στο ts_settings_frame
        self.ts_interval_lbl = ttk.Label(ts_frame, text=t["ts_interval"])
        self.ts_interval_lbl.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Entry(ts_frame, textvariable=self.ts_interval, width=5).pack(side=tk.LEFT, padx=5)


        # Sampling Rate
        self.sampling_lbl = ttk.Label(self.root, text=t["sampling"])
        self.sampling_lbl.grid(row=6, column=0, padx=5, pady=3, sticky="w")
        
        slider_frame = ttk.Frame(self.root)
        slider_frame.grid(row=6, column=1, padx=5, pady=3, sticky="ew")
        self.sampling_rate_slider = ttk.Scale(slider_frame, from_=0, to=1000, variable=self.sampling_rate, orient=tk.HORIZONTAL, command=self.update_sampling_rate_label)
        self.sampling_rate_slider.pack(side=tk.LEFT)
        self.sampling_rate_value_label = ttk.Label(slider_frame, text="0 ms")
        self.sampling_rate_value_label.pack(side=tk.LEFT, padx=5)


        # Control Buttons 
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.grid(row=6, column=1, pady=5, sticky="e")
        # Τα υπάρχοντα CONTROL κουμπιά σου ακολουθούν μετά (Start, Stop, κτλ)
        self.start_btn = ttk.Button(ctrl_frame, text=t["start"], command=self.start_logging)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn = ttk.Button(ctrl_frame, text=t["stop"], command=self.stop_logging)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        self.save_btn = ttk.Button(ctrl_frame, text=t["save"], command=self.save_data)
        self.save_btn.pack(side=tk.LEFT, padx=2)
        self.clear_btn = ttk.Button(ctrl_frame, text=t["clear"], command=self.clear_data)
        self.clear_btn.pack(side=tk.LEFT, padx=2)

        # Threshold & Scroll
        thresh_frame = ttk.Frame(self.root)
        thresh_frame.grid(row=10, column=1, columnspan=2, padx=5, pady=3, sticky="w")
        
        self.graph_win_lbl = ttk.Label(thresh_frame, text=t["graph_win"])
        self.graph_win_lbl.pack(side=tk.LEFT)
        ttk.Entry(thresh_frame, textvariable=self.max_val_limit, width=7).pack(side=tk.LEFT, padx=2)
        
        self.scroll_chk = ttk.Checkbutton(thresh_frame, text=t["scroll"], variable=self.scroll_mode)
        self.scroll_chk.pack(side=tk.LEFT, padx=5)
        ttk.Entry(thresh_frame, textvariable=self.scroll_window_size, width=5).pack(side=tk.LEFT)
        self.points_lbl = ttk.Label(thresh_frame, text=t["points"])
        self.points_lbl.pack(side=tk.LEFT)

        # Data Area
        data_label_frame = ttk.Frame(self.root)
        data_label_frame.grid(row=10, column=0, padx=5, sticky="w")
        self.log_win_lbl = ttk.Label(data_label_frame, text=t["log_win"])
        self.log_win_lbl.pack(side=tk.LEFT)
        # Προσθήκη του ορίου 
        self.listbox_lbl = ttk.Label(data_label_frame, text=" | " + t["listbox_limit"])
        self.listbox_lbl.pack(side=tk.LEFT, padx=(5, 2))
        self.listbox_entry = ttk.Entry(data_label_frame, textvariable=self.listbox_limit, width=8)
        self.listbox_entry.pack(side=tk.LEFT)
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.grid(row=11, column=0, columnspan=5, padx=5, pady=3, sticky="nsew")

        # Listbox
        list_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(list_frame, weight=1)
        self.data_listbox = tk.Listbox(list_frame, height=10, selectmode=tk.EXTENDED)
        self.data_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scbr = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.data_listbox.yview)
        scbr.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_listbox.config(yscrollcommand=scbr.set)

        # Plot
        fig = Figure(dpi=100)
        self.ax = fig.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(fig, master=self.paned_window)
        self.paned_window.add(self.canvas.get_tk_widget(), weight=3)

        # Context Menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label=t["copy"], command=self.copy_to_clipboard)
        self.context_menu.add_command(label=t["export_csv"], command=self.export_selected_to_csv)
        self.context_menu.add_command(label=t["export_xlsx"], command=self.export_selected_to_xlsx)
        self.data_listbox.bind("<Button-3>", self.show_context_menu)
        self.data_listbox.bind("<Button-2>", self.show_context_menu)

        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(11, weight=1)

    # --- Παραμένουν οι υπόλοιπες συναρτήσεις (record_data, save_data, κτλ) ίδιες ---
    def open_instructions_window(self):
        instructions_window = tk.Toplevel(self.root)
        instructions_window.title("Οδηγίες / Instructions")
        instructions_window.geometry("750x700")
        
        text_el = (
            "Καταγραφή δεδομένων από serial (Serial Data Logger).\n\n\n"
            "Μπορείτε να:\n\n" 
            "1. Ορίστε τη θύρα από την οποία θα διαβάσετε δεδομένα.\n"
            "   (με [Aνανέωση] διαβάζονται ξανά οι διαθέσιμες θύρες, \n"
            "   σε περίπτωση που συνδέσατε τον μικροεπεξεργαστή μετά το άνοιγμα αυτής εδώ της εφαρμογής)\n\n"
            "2. Ορίσετε το Baudrate για τη σύνδεση (Παράδειγμα: 9600 για Mind+ ή 115200 για MakeCode).\n\n"
            "3. Επιλέξετε το όνομα του αρχείου και τον τύπο του (.xlsx ή .csv), για αποθήκευση των μετρήσεων.\n\n"
            "4. Ορίστε τους τίτλους των στηλών στο .xlsx (μέχρι 8)\n\n"
            "5. Επιλέξετε αν οι μετρήσεις (μέχρι 8) θα εξάγονται ταυτόχρονα στο ThinkSpeeak το οποίο δέχεται τιμές κάθε 15''.\n"
            "   (Θα χρειαστεί να oρίσετε και το API Key που θα βρείτε στην αντίστοιχη επιλογή της διαδικτυακής εφαρμογής ThinkSpeak).\n\n"
            "6. Επιλέξετε την καθυστέρηση μεταξύ των δειγματοληψιών (καλό είναι να ρυθμίζεται από το πρόγραμμα που τις εξάγει)\n\n"
            "7. Επιλέξετε αν θα scrollάρει το διάγραμμα και για πόσα σημεία\n\n"
            "8. Επιλέξετε το άνω όριο των τιμών που θα εμφανίζονται  στο διάγραμμα\n\n\n\n"
            
            "Λειτουργίες:\n"
            "_______________________\n\n"
            "Πατήστε [Έναρξη] για να ξεκινήσετε τη καταγραφή.\n"
            "Πατήστε [Τερματισμός] για να σταματήσετε την καταγραφή.\n"
            "Πατήστε [Αποθήκευση στο αρχείο] για να αποθηκεύσετε τις μετρήσεις στο αρχείο που ήδη έχετε επιλέξει.\n"
            "   (μπορείτε και πριν τον τερματισμό να αποθηκεύετε τιμές στο αρχείο, οι οποίες θα προστεθούν σ' αυτό)\n"
            "Εναλλακτικά μπορείτε να αποθηκεύσετε στη μνήμη, σε άλλο αρχείο (.xlsx, .csv) ...και με δεξί κλικ πάνω στο παράθυρο των τιμών \n"
            "   (επιλέγοντας κάποιες aαπό αυτές ή/και όλες τις γραμμές που έχουν καταγραφεί).\n\n"
            "Πατήστε [Καθαρισμός] για να καθαρίσετε το διάγραμμα και τις τρέχουσες τιμές\n"
            "   (Δεν διαγράγονται τιμές από το .xlsx που ήδη έχετε αποθηκεύσει από προηγούμενη φορά). \n"
            "Μπορείτε να ρυθμίσετε σε μήκος το παράθυρο καταγραφής τιμών ...και του διαγράμματος,  \n" 
            "   σύρροντας την ενδιάμεση διαχωριστική μπάρα δεξιά ή αριστερά.\n\n\n"
           
          
            "Ελπίζω να σας φανεί χρήσιμη η εφαρμογή αυτή.\n"
        )
        
        text_en = (
            "Serial Data Logger.\n\n\n"
            "You can:\n\n" 
            "1. Set the port from which data will be read.\n"
            "   (use [Refresh] to scan for available ports again, \n"
            "   in case you connected the microprocessor after opening this application)\n\n"
            "2. Set the Baudrate for the connection (Example: 9600 for Mind+ or 115200 for MakeCode).\n\n"
            "3. Choose the file name and type (.xlsx or .csv) to save the measurements.\n\n"
            "4. Set the column titles in the .xlsx file (up to 8)\n\n"
            "5. Choose if the measurements (up to 8) will be exported simultaneously to ThingSpeak, which accepts values every 15''.\n"
            "   (You will also need to provide the API Key found in the corresponding option of the ThingSpeak web application).\n\n"
            "6. Select the delay between samples (it is recommended to be regulated by the source program exporting them)\n\n"
            "7. Choose whether the chart will scroll and for how many points\n\n"
            "8. Select the upper limit for the values displayed on the chart\n\n\n\n"
            
            "Functions:\n"
            "_______________________\n\n"
            "Press [Start] to begin logging.\n"
            "Press [Stop] to stop logging.\n"
            "Press [Save to file] to save the measurements to the file you have already selected.\n"
            "   (you can save values to the file even before stopping, which will be appended to it)\n"
            "Alternatively, you can save to memory, to another file (.xlsx, .csv) ...and by right-clicking on the values window \n"
            "   (by selecting some or all of the recorded lines).\n\n"
            "Press [Clear] to clear the chart and the current values\n"
            "   (Values already saved in the .xlsx file from previous times will not be deleted). \n"
            "You can adjust the width of the log window ...and the chart, \n" 
            "   by dragging the middle separator bar right or left.\n\n\n"
           
            "I hope you find this application useful.\n"
        )
        
        display_text = text_el if self.current_lang == "EL" else text_en
        tk.Label(instructions_window, text=display_text, justify=tk.LEFT, font=("Arial", 11)).pack(padx=10, pady=10)
        ttk.Button(instructions_window, text="OK", command=instructions_window.destroy).pack(pady=5)

    def update_sampling_rate_label(self, value):
        self.sampling_rate_value_label.config(text=f"{int(float(value))} ms")

    def show_context_menu(self, event):
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def refresh_ports(self):
        ports = [port.device for port in list_ports.comports()]
        self.ports_combobox["values"] = ports
        if ports: self.ports_combobox.current(0)

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if file_path: self.output_path.set(file_path)

    def copy_to_clipboard(self, event=None):
        selected_indices = self.data_listbox.curselection()
        if not selected_indices: return
        selected_text = "\n".join([self.data_listbox.get(i) for i in selected_indices])
        self.root.clipboard_clear()
        self.root.clipboard_append(selected_text)

    def export_selected_to_csv(self): self._export_selected_logic(".csv")
    def export_selected_to_xlsx(self): self._export_selected_logic(".xlsx")

    def _export_selected_logic(self, extension):
        selected_indices = self.data_listbox.curselection()
        if not selected_indices: return
        file_path = filedialog.asksaveasfilename(defaultextension=extension)
        if not file_path: return
        try:
            headers = ["Timestamp"] + [v.get() for v in self.extra_text_vars if v.get()]
            rows = []
            for i in selected_indices:
                raw_line = self.data_listbox.get(i)
                if ": " in raw_line:
                    ts, vals = raw_line.split(": ", 1)
                    rows.append([ts] + vals.split(", "))
            if extension == ".csv":
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    csv.writer(f).writerow(headers)
                    csv.writer(f).writerows(rows)
            else:
                nb = Workbook()
                ws = nb.active
                ws.append(headers)
                for r in rows: ws.append(r)
                nb.save(file_path)
            messagebox.showinfo("Export", "Done!")
        except Exception as e: messagebox.showerror("Error", str(e))

    def connect_to_serial(self):
        try:
            return serial.Serial(self.ports_combobox.get(), baudrate=self.baudrate.get(), timeout=1)
        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
            return None

    def start_logging(self):
        self.serial_port = self.connect_to_serial()
        if not self.serial_port: return
        self.stop_event.clear()
        threading.Thread(target=self.record_data, daemon=True).start()
        self.update_plot()

    def stop_logging(self):
        if self.serial_port:
            self.stop_event.set()
            self.serial_port.close()
            self.serial_port = None

    def record_data(self):
            try:
                while not self.stop_event.is_set():
                    line = self.serial_port.readline().decode('utf-8', errors='ignore').strip() 
                    if line:
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        
                        # Λήψη του ορίου από το GUI
                        current_max = self.max_val_limit.get()
                        
                        # Αντικατάσταση διαχωριστικών
                        line = line.replace(';', ',').replace(':', ',')
                        raw_items = [item.strip() for item in line.split(',') if item.strip()]
                        
                        clean_numeric_values = []
                        for item in raw_items:
                            try:
                                val = float(item)
                                # Χρήση της μεταβλητής current_max
                                if val > current_max:
                                    clean_numeric_values.append(0.0)
                                else:
                                    clean_numeric_values.append(val)
                            except ValueError:
                                clean_numeric_values.append(0.0)

                        # 1. Εμφάνιση στην ουρά
                        self.data_queue.put((timestamp, clean_numeric_values, raw_items))

                        # 2. Αποθήκευση στο Excel (όλα τα raw δεδομένα)
                        #if self.get_file_extension() == ".xlsx":
                        #    excel_row = raw_items[:8]
                        #    padding = [None] * (8 - len(excel_row))
                        #    self.sheet.append([timestamp, *excel_row, *padding])
                        
                        # 3. Αποστολή στο ThingSpeak
                        self.send_to_thingspeak_api(clean_numeric_values)
                        threading.Event().wait(self.sampling_rate.get() / 1000)

            except Exception as e:
                if not self.stop_event.is_set():
                    messagebox.showerror("Σφάλμα", str(e))


    #def send_to_thingspeak_api(self, values):
    #    if self.send_to_thingspeak.get():
    #        threading.Thread(target=self._async_ts, args=(values,), daemon=True).start()


    def send_to_thingspeak_api(self, values):
        if self.send_to_thingspeak.get():
            now = datetime.now()
            # Υπολογισμός διαφοράς χρόνου σε δευτερόλεπτα
            diff = (now - self.last_ts_send).total_seconds()
            
            if diff >= self.ts_interval.get():
                self.last_ts_send = now
                threading.Thread(target=self._async_ts, args=(values,), daemon=True).start()

    def _async_ts(self, values):
        try:
            url = "https://api.thingspeak.com/update"
            params = {"api_key": self.thingspeak_api_key.get()}
            for i, v in enumerate(values[:8]): params[f"field{i+1}"] = v
            requests.get(url, params=params, timeout=5)
        except: pass

   
           
    def update_plot(self):
        # 1. Διαβάζουμε όλα τα νέα δεδομένα από την ουρά
        while not self.data_queue.empty():
            timestamp, numeric_values, raw_items = self.data_queue.get()
            self.times.append(len(self.times) + 1) # Παραμένει αύξοντας αριθμός για το γράφημα
            self.actual_timestamps.append(timestamp) # Αποθήκευση της ώρας για το Excel
            self.values.append(numeric_values)
            self.data_listbox.insert(tk.END, f"{timestamp}: {', '.join(raw_items)}")
            self.data_listbox.see(tk.END)
            
            # Λαμβάνουμε το όριο που έγραψε ο χρήστης στο GUI
            current_limit = self.listbox_limit.get()
            if self.data_listbox.size() > current_limit:
                # Διαγράφουμε το 10% των παλαιότερων τιμών για να μην τρέχει συνέχεια η διαγραφή
                delete_count = max(1, int(current_limit * 0.1))
                self.data_listbox.delete(0, delete_count)
            # ------------------------------


        #  2. Σχεδιασμός του διαγράμματος
        if self.times:
            self.ax.clear()
            
            # Υπολογισμός του "παραθύρου" εμφάνισης
            if self.scroll_mode.get():
                window = self.scroll_window_size.get()
                # Παίρνουμε μόνο τα τελευταία N στοιχεία
                plot_times = self.times[-window:]
                plot_values = self.values[-window:]
            else:
                plot_times = self.times
                plot_values = self.values

            # Οργάνωση των δεδομένων σε στήλες
            data = list(zip_longest(*plot_values, fillvalue=0.0))
            data = [list(col) for col in data]
            
            for i, col in enumerate(data):
                if i < 8:
                    label = self.extra_text_vars[i].get()
                    self.ax.plot(plot_times, col, label=label)
            
            self.ax.set_xlabel("Αριθμός μετρήσεων" + (" (Τελευταίες)" if self.scroll_mode.get() else ""))
            self.ax.set_ylabel("Τιμή")
            self.ax.legend()
            self.canvas.draw()

        if not self.stop_event.is_set():
            self.root.after(100, self.update_plot)
    
    

    def save_data(self):
        path = self.output_path.get()
        ext = ".xlsx" if path.endswith(".xlsx") else ".csv"
        if not self.times: return
        headers = ["Time"] + [v.get() for v in self.extra_text_vars if v.get()]
        rows = [[t] + list(v) for t, v in zip(self.actual_timestamps, self.values)]
        try:
            if ext == ".xlsx":
                wb = openpyxl.load_workbook(path) if os.path.exists(path) else Workbook()
                ws = wb.active
                if wb.get_sheet_names() == ['Sheet']: ws.append(headers) # Simple check for new file
                for r in rows: ws.append(r)
                wb.save(path)
            else:
                exists = os.path.exists(path)
                with open(path, "a", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    if not exists: writer.writerow(headers)
                    writer.writerows(rows)
            messagebox.showinfo("Save", "Success!")
        except Exception as e: messagebox.showerror("Error", str(e))

    def clear_data(self):
        if messagebox.askyesno("Clear", "Delete all data?"):
            self.times, self.values, self.actual_timestamps = [], [], []
            self.data_listbox.delete(0, tk.END)
            self.ax.clear()
            self.canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialDataLogger(root)
    root.mainloop()
