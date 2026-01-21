#Vassilis Economou  16/01/2025 v.02
#                   20/01/2026 v.2.1    

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

# Απενεργοποίηση προειδοποιήσεων από matplotlib
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

class SerialDataLogger:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Data Logger")

        # Προσθήκη εικονιδίου και τίτλου
        #self.root.iconbitmap("icon.ico")  # Αντικαταστήστε με το όνομα του αρχείου εικονιδίου
        title_label = ttk.Label(self.root, text="Serial Data Logger  [Βασίλης Οικονόμου v.2.1]", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # Αρχικοποίηση μεταβλητών
        self.serial_port = None
        self.baudrate = tk.IntVar(value=9600)
        self.max_val_limit = tk.IntVar(value=1024)  # Νέα μεταβλητή με προεπιλογή το 1024
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "BO_SDL.xlsx"))
        self.times = []
        self.values = []
        self.data_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.sampling_rate = tk.IntVar(value=0)  # Ταχύτητα δειγματοληψίας σε ms

        # Επιλογή ThingSpeak
        self.send_to_thingspeak = tk.BooleanVar(value=False)
        self.thingspeak_api_key = tk.StringVar(value="0J62FHGN0IS42VNQ")  # Αρχικό API Key
        
        # Επιλογή scroling
        self.scroll_mode = tk.BooleanVar(value=False)
        self.scroll_window_size = tk.IntVar(value=500) # Προεπιλογή τελευταία 500 σημεία

        # για τo timestamp στην πρώτη στήλη του .xlsx
        self.times = []
        self.actual_timestamps = [] # Νέα λίστα για το Excel
        self.values = []



        self.create_widgets()

    # Προσθήκη της συνάρτησης για το παράθυρο οδηγών
    def open_instructions_window(self):
        # Δημιουργία νέου παραθύρου
        instructions_window = tk.Toplevel(self.root)
        instructions_window.title("Οδηγίες")
        instructions_window.geometry("750x700")

        # Εισαγωγή κειμένου με οδηγίες
        instructions_text = (
            "Καταγραφή δεδομένων από serial (Serial Data Logger).\n\n\n"
            "Μπορείτε να:\n\n" 
            "1. Ορίστε τη θύρα από την οποία θα διαβάσετε δεδομένα.\n"
            "   (με [Aνανέωση] διαβάζονται ξανά οι διαθέσιμες θύρες, \n"
            "   σε περίπτωση που συνδέσατε τον μικροεπεξεργαστή μετά το άνοιγμα αυτής εδώ της εφαρμογής)\n\n"
            "2. Ορίσετε το Baudrate για τη σύνδεση (Παράδειγμα: 9600 για Mind+ ή 115200 για MakeCode).\n\n"
            "3. Επιλέξετε το όνομα του αρχείου και τον τύπο του (.xlsx ή .csv), για αποθήκευση των μετρήσεων.\n\n"
            "4. Ορίστε τους τίτλους των στηλών στο .xlsx (μέχρι 8)\n\n"
            "5. Επιλέξετε αν οι μετρήσεις (μέχρι 8) θα εξάγονται ταυτόχρονα στο ThinkSpeeak το οποίο δέχεται τιμές κάθε 15''.\n"
            "   (Θα χρειαστεί να oρίσετε και το API Key που θα βρείτε στην αντίστοιχη επιλογή τηςδιαδικτυακής εφαρμογής ThinkSpeak).\n\n"
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
        # Εμφάνιση κειμένου
        text_widget = tk.Label(instructions_window, text=instructions_text, justify=tk.LEFT, font=("Arial", 12))
        text_widget.pack(padx=10, pady=10)
        # Προσθήκη κουμπιού για κλείσιμο
        close_button = ttk.Button(instructions_window, text="Κλείσιμο", command=instructions_window.destroy)
        close_button.pack(pady=5)

    def create_widgets(self):
        # Προσθήκη κουμπιού για οδηγίες χρήσης
        instructions_button = ttk.Button(self.root, text="  Οδηγίες  ", command=self.open_instructions_window)
        instructions_button.grid(row=0, column=2, pady=3)
        
        # --- ΕΝΙΑΙΑ ΓΡΑΜΜΗ ΡΥΘΜΙΣΕΩΝ ΣΥΝΔΕΣΗΣ ---
        connection_frame = ttk.Frame(self.root)
        connection_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        # Στοιχεία Θύρας
        ttk.Label(connection_frame, text="Θα διαβάσω από τη Θύρα:").pack(side=tk.LEFT, padx=5)
        self.ports_combobox = ttk.Combobox(connection_frame, state="readonly", width=20)
        self.ports_combobox.pack(side=tk.LEFT, padx=5)
        self.refresh_ports()
        ttk.Button(connection_frame, text="Ανανέωση", command=self.refresh_ports).pack(side=tk.LEFT, padx=5)
        # Στοιχεία Baudrate
        ttk.Label(connection_frame, text="...με ρυθμό (Baudrate):").pack(side=tk.LEFT, padx=(20, 5))
        baudrate_combobox = ttk.Combobox(connection_frame, textvariable=self.baudrate, state="readonly", width=10)
        baudrate_combobox["values"] = [9600, 19200, 38400, 57600, 115200]
        baudrate_combobox.pack(side=tk.LEFT, padx=5)

        # --- ΕΝΙΑΙΑ ΓΡΑΜΜΗ ΕΠΙΛΟΓΗΣ ΑΡΧΕΙΟΥ ---
        file_selection_frame = ttk.Frame(self.root)
        file_selection_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        ttk.Label(file_selection_frame, text="Θα αποθηκεύσω στο αρχείο (.xlsx ή .csv):").pack(side=tk.LEFT, padx=5)
        output_entry = ttk.Entry(file_selection_frame, width=60, textvariable=self.output_path)
        output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=False)
        ttk.Button(file_selection_frame, text="Επιλογή άλλου αρχείου", command=self.browse_file).pack(side=tk.LEFT, padx=5)


        # 8 fields for the column titles - Όλα σε μία γραμμή
        sampling_rate_label = ttk.Label(self.root, text="Ονόματα στηλών στο .xlsx:")
        sampling_rate_label.grid(row=4, column=0, padx=5, pady=3, sticky="w")
        fields_frame = ttk.Frame(self.root)
        fields_frame.grid(row=4, column=1, columnspan=1, padx=5, pady=(10, 3), sticky="ew")
        # Ρύθμιση 8 στηλών με ίσο βάρος
        for c in range(8): 
            fields_frame.grid_columnconfigure(c, weight=1)
    
        self.extra_text_vars = [tk.StringVar() for _ in range(8)]
        self.extra_entries = []
    
        for i in range(8):
            # width=7 κάνει το κουτάκι μικρό οπτικά
            e = ttk.Entry(fields_frame, textvariable=self.extra_text_vars[i], width=7)
            e.grid(row=0, column=i, padx=2, pady=3, sticky="ew")
            self.extra_entries.append(e)

        for i, v in enumerate(self.extra_text_vars): 
            v.set(f"Στήλη{i+1}")



        # Επιλογή για ThingSpeak
        thingspeak_check = ttk.Checkbutton(self.root, text="Αποστολή και στο ThingSpeak με API Key:", variable=self.send_to_thingspeak)
        thingspeak_check.grid(row=5, column=0, padx=5, pady=3, sticky="w")
        api_key_entry = ttk.Entry(self.root, textvariable=self.thingspeak_api_key)
        api_key_entry.grid(row=5, column=1, padx=5, pady=3, sticky="w")


        # --- Ταχύτητα Δειγματοληψίας και Όριο Τιμής ---  
        # 1. Sampling Rate Elements
        sampling_rate_label = ttk.Label(self.root, text="Καθυστέρηση σε προβολή & απεικόνηση (ms):")
        sampling_rate_label.grid(row=6, column=0, padx=5, pady=3, sticky="w")
        
        # Δημιουργούμε ένα μικρό frame για να βάλουμε το slider και την τιμή του μαζί στη στήλη 
        slider_frame = ttk.Frame(self.root)
        slider_frame.grid(row=6, column=1, padx=5, pady=3, sticky="ew")
        
        self.sampling_rate_slider = ttk.Scale(slider_frame, from_=0, to=1000, variable=self.sampling_rate, orient=tk.HORIZONTAL)
        self.sampling_rate_slider.pack(side=tk.LEFT, fill=tk.X, expand=False)
        
        self.sampling_rate_value_label = ttk.Label(slider_frame, text=f"{self.sampling_rate.get()} ms")
        self.sampling_rate_value_label.pack(side=tk.LEFT, padx=5)
        self.sampling_rate_slider.config(command=self.update_sampling_rate_label)

    
 
        # Κουμπιά ελέγχου
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.grid(row=6, column=1, columnspan=4, pady=5)
        ttk.Button(ctrl_frame, text="Έναρξη", command=self.start_logging).pack(side=tk.LEFT, padx=4)
        ttk.Button(ctrl_frame, text="Τερματισμός", command=self.stop_logging).pack(side=tk.LEFT, padx=4)
        ttk.Button(ctrl_frame, text="Αποθήκευση στο αρχείο", command=self.save_data).pack(side=tk.LEFT, padx=4)
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.grid(row=6, column=2, columnspan=4, pady=5)
        ttk.Button(ctrl_frame, text="Καθαρισμός", command=self.clear_data).pack(side=tk.LEFT, padx=4)
     
        

        # άνω όριο τιμής Elements (Στην ίδια γραμμή, στήλη 2)
        threshold_frame = ttk.Frame(self.root)
        threshold_frame.grid(row=10, column=1, padx=5, pady=3, sticky="w")
        threshold_label = ttk.Label(threshold_frame, text="Παράθυρο διαγράμματος  [Άνώτατο αποδεκτό όριο τιμών:") 
        threshold_label.pack(side=tk.LEFT, padx=2)
        threshold_entry = ttk.Entry(threshold_frame, textvariable=self.max_val_limit, width=7)
        threshold_entry.pack(side=tk.LEFT, padx=2)

        # --- Επιλογή Scrolling Mode ---
        #scroll_frame = ttk.Frame(self.root)
        #scroll_frame.grid(row=10, column=2, padx=5, pady=3, sticky="w") 
        scroll_check = ttk.Checkbutton(threshold_frame, text="Scrolling προς τα αριστερά ", variable=self.scroll_mode)
        scroll_check.pack(side=tk.LEFT, padx=2)
        scroll_entry = ttk.Entry(threshold_frame, textvariable=self.scroll_window_size, width=5)
        scroll_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(threshold_frame, text="σημεία].").pack(side=tk.LEFT)

        # --- Περιοχή εμφάνισης δεδομένων
        sampling_rate_label = ttk.Label(self.root, text="Παράθυρο καταγραφής τιμών από τη θύρα")
        sampling_rate_label.grid(row=10, column=0, padx=5, pady=3, sticky="w")
        self.data_listbox = tk.Listbox(self.root, height=10)
        self.data_listbox.grid(row=11, column=0, columnspan=4, padx=5, pady=3, sticky="nsew")


        # Διάγραμμα
        figure = Figure(figsize=(7, 5), dpi=100)
        self.ax = figure.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(figure, master=self.root)
        self.canvas.get_tk_widget().grid(row=11, column=1, columnspan=4, pady=3, sticky="nsew")


        # Ρύθμιση διαστάσεων πλέγματος
        #self.root.columnconfigure(1, weight=1)
        #self.root.rowconfigure(9, weight=1)

        # --- Περιοχή εμφάνισης δεδομένων με PanedWindow ---
        # Δημιουργία του PanedWindow (οριζόντιο)
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        
        # Το sticky="nsew" λέει στο widget να "κολλήσει" σε όλες τις πλευρές (North, South, East, West)
        self.paned_window.grid(row=11, column=0, columnspan=5, padx=5, pady=3, sticky="nsew")

        # 1. Frame για το Listbox και το Scrollbar (αριστερά)
        listbox_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(listbox_frame, weight=1)

        self.data_listbox = tk.Listbox(listbox_frame, height=10)
        self.data_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.data_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_listbox.config(yscrollcommand=scrollbar.set)

        # 2. Το Διάγραμμα (δεξιά)
        # Αφαιρούμε το σταθερό figsize ή το βάζουμε ως ελάχιστο, 
        # το fill=BOTH θα αναλάβει τα υπόλοιπα.
        figure = Figure(dpi=100) 
        self.ax = figure.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(figure, master=self.paned_window)
        plot_widget = self.canvas.get_tk_widget()
        self.paned_window.add(plot_widget, weight=3)

       
        # --- ΡΥΘΜΙΣΗ ΔΥΝΑΜΙΚΟΥ ΜΕΓΕΘΟΥΣ (CRITICAL) ---
        # Λέμε στο κεντρικό παράθυρο (root) ότι η γραμμή 11 (που έχει το PanedWindow) 
        # και η στήλη 1 πρέπει να επεκτείνονται όταν μεγαλώνει το παράθυρο.
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(11, weight=1)




        #------COPY PASTE

        # Επιτρέπουμε την επιλογή πολλαπλών γραμμών
        self.data_listbox.config(selectmode=tk.EXTENDED)

       # Ενημέρωση του μενού δεξιού κλικ
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Αντιγραφή", command=self.copy_to_clipboard)
        self.context_menu.add_command(label="Εξαγωγή επιλεγμένων σε .csv", command=self.export_selected_to_csv)
        self.context_menu.add_command(label="Εξαγωγή επιλεγμένων σε .xlsx", command=self.export_selected_to_xlsx)

        # Σύνδεση δεξιού κλικ (Button-3 για Windows/Linux, Button-2 για macOS)
        self.data_listbox.bind("<Button-3>", self.show_context_menu)
        self.data_listbox.bind("<Button-2>", self.show_context_menu)
        
        # Σύνδεση Ctrl+C για ευκολία
        self.data_listbox.bind("<Control-c>", self.copy_to_clipboard)


    def show_context_menu(self, event):
        """Εμφανίζει το μενού στη θέση του κέρσορα"""
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def copy_to_clipboard(self, event=None):
        """Αντιγράφει τις επιλεγμένες γραμμές στο πρόχειρο"""
        selected_indices = self.data_listbox.curselection()
        if not selected_indices: return
        selected_text = "\n".join([self.data_listbox.get(i) for i in selected_indices])
        self.root.clipboard_clear()
        self.root.clipboard_append(selected_text)

    def export_selected_to_csv(self):
        """Εξαγωγή σε CSV"""
        self._export_selected_logic(".csv")

    def export_selected_to_xlsx(self):
        """Εξαγωγή σε XLSX"""
        self._export_selected_logic(".xlsx")

    def _export_selected_logic(self, extension):
        """Κοινή λογική εξαγωγής για CSV και XLSX"""
        selected_indices = self.data_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Εξαγωγή", "Παρακαλώ επιλέξτε πρώτα τις γραμμές.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=extension,
                                                 filetypes=[("Excel/CSV Files", f"*{extension}")],
                                                 title=f"Αποθήκευση επιλεγμένων σε {extension}")
        if not file_path: return

        try:
            # Προετοιμασία δεδομένων
            headers = ["Timestamp"] + [v.get() for v in self.extra_text_vars if v.get()]
            rows_to_save = []
            for i in selected_indices:
                raw_line = self.data_listbox.get(i)
                if ": " in raw_line:
                    timestamp, vals = raw_line.split(": ", 1)
                    rows_to_save.append([timestamp] + vals.split(", "))
                else:
                    rows_to_save.append([raw_line])

            if extension == ".csv":
                with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(rows_to_save)
            else:
                from openpyxl import Workbook
                new_wb = Workbook()
                ws = new_wb.active
                ws.append(headers)
                for r in rows_to_save: ws.append(r)
                new_wb.save(file_path)

            messagebox.showinfo("Εξαγωγή", f"Η αποθήκευση σε {extension} ολοκληρώθηκε!")
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Αποτυχία: {str(e)}")

#____________






    def update_sampling_rate_label(self, value):
        self.sampling_rate_value_label.config(text=f"{int(float(value))} ms")



    
    #def send_to_thingspeak_api(self, values):
    #    if self.send_to_thingspeak.get():  # Έλεγχος αν το checkbox είναι ενεργοποιημένο
    #        try:
    #            api_key = self.thingspeak_api_key.get()
    #            if not api_key:
    #                messagebox.showwarning("API Key", "Παρακαλώ εισάγετε το API Key.")
    #                return
    #            url = f"https://api.thingspeak.com/update?api_key={api_key}"
    #            payload = {label: value for label, value in zip([f"field{i+1}" for i in range(len(values))], values)}
    #            response = requests.get(url, params=payload)
    #            # print(payload, response)
    #            if response.status_code == 200:
    #                print(values)
    #            else:
    #                messagebox.showerror("Σφάλμα αποστολής στο ThingSpeak.")
    #        except Exception as e:
    #            messagebox.showerror("Σφάλμα σύνδεσης στο ThingSpeak", str(e))

    def send_to_thingspeak_api(self, values):
        if self.send_to_thingspeak.get():
            # Δημιουργούμε ένα thread για να τρέξει η αποστολή στο background
            threading.Thread(target=self._async_thingspeak_request, args=(values,), daemon=True).start()

    def _async_thingspeak_request(self, values):
        try:
            api_key = self.thingspeak_api_key.get()
            if not api_key:
                return
            
            url = f"https://api.thingspeak.com/update?api_key={api_key}"
            payload = {label: value for label, value in zip([f"field{i+1}" for i in range(len(values))], values)}
            
            # Εδώ γίνεται η κλήση που καθυστερεί, αλλά πλέον τρέχει σε δικό της thread
            response = requests.get(url, params=payload, timeout=5)
            if response.status_code == 200:
                print(values)
            else:
                messagebox.showerror("Σφάλμα αποστολής στο ThingSpeak.")
        except Exception as e:
            messagebox.showerror("Σφάλμα σύνδεσης στο ThingSpeak", str(e))
    
    
    
    def refresh_ports(self):
        ports = [port.device for port in list_ports.comports()]
        self.ports_combobox["values"] = ports
        if ports:
            self.ports_combobox.current(0)

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
        if file_path:
            self.output_path.set(file_path)

    
    #def save_data(self):
    #    file_extension = self.get_file_extension()
    #    if file_extension == ".xlsx":
    #        try:
    #            self.wb.save(self.output_path.get()) #type:ignore
    #            messagebox.showinfo("Αποθήκευση", "Τα δεδομένα αποθηκεύτηκαν με επιτυχία σε .xlsx.")
    #        except Exception as e:
    #            messagebox.showerror("Σφάλμα αποθήκευσης", str(e))
    #    elif file_extension == ".csv":
    #        try:
    #            with open(self.output_path.get(), mode="w", newline="", encoding="utf-8") as file:
    #                writer = csv.writer(file)
    #                labels = [v.get() for v in self.extra_text_vars if v.get() != ""]
    #                writer.writerow(["Time"] + labels)
    #                data = list(zip_longest(*self.values, fillvalue=0.))
    #                data = [list(col) for col in data]
    #                writer.writerows(zip(self.times, *data))
    #            messagebox.showinfo("Αποθήκευση", "Τα δεδομένα αποθηκεύτηκαν με επιτυχία σε .csv.")
    #        except Exception as e:
    #            messagebox.showerror("Σφάλμα αποθήκευσης", str(e))

    def save_data(self):
            file_path = self.output_path.get()
            file_extension = self.get_file_extension()
            
            if not self.times:
                messagebox.showwarning("Αποθήκευση", "Δεν υπάρχουν δεδομένα για αποθήκευση.")
                return

            # Προετοιμασία δεδομένων (μετατροπή στηλών σε γραμμές)
            headers = ["Time"] + [v.get() for v in self.extra_text_vars if v.get() != ""]
            data_to_save = list(zip_longest(*self.values, fillvalue=0.0))
            data_to_save = [list(col) for col in data_to_save]
            #full_rows = [[t] + list(v) for t, *v in zip(self.times, *data_to_save)]
            # Χρησιμοποιούμε actual_timestamps αντί για times
            full_rows = [[t] + list(v) for t, *v in zip(self.actual_timestamps, *data_to_save)]

            try:
                if file_extension == ".xlsx":
                    if os.path.exists(file_path):
                        # Αν το αρχείο υπάρχει, το ανοίγουμε για προσθήκη
                        wb = openpyxl.load_workbook(file_path)
                        ws = wb.active
                    else:
                        # Αν δεν υπάρχει, δημιουργούμε νέο και βάζουμε κεφαλίδες
                        wb = Workbook()
                        ws = wb.active
                        ws.append(headers)
                    
                    # Προσθήκη των δεδομένων
                    for row in full_rows:
                        ws.append(row)
                    wb.save(file_path)

                elif file_extension == ".csv":
                    file_exists = os.path.exists(file_path)
                    # "a" σημαίνει append (προσθήκη στο τέλος)
                    with open(file_path, mode="a", newline="", encoding="utf-8-sig") as f:
                        writer = csv.writer(f)
                        # Αν το αρχείο είναι νέο, γράφουμε πρώτα τις κεφαλίδες
                        if not file_exists:
                            writer.writerow(headers)
                        writer.writerows(full_rows)

                messagebox.showinfo("Αποθήκευση", f"Τα δεδομένα προστέθηκαν με επιτυχία στο {file_extension}!")
                
                # ΠΡΟΑΙΡΕΤΙΚΟ: Καθαρισμός των λιστών μετά την αποθήκευση 
                # ώστε την επόμενη φορά να αποθηκευτούν μόνο τα *νέα* δεδομένα
                # self.times = []
                # self.values = []
                
            except Exception as e:
                messagebox.showerror("Σφάλμα αποθήκευσης", str(e))




    def setup_file(self):
        file_extension = self.get_file_extension()
        if file_extension == ".xlsx":
            path = self.output_path.get()
            try:
                wb = openpyxl.load_workbook(path)
                sheet = wb.active
            except (FileNotFoundError, KeyError):
                wb = Workbook()
                sheet = wb.active
                sheet.append(["Time"] + [v.get() for v in self.extra_text_vars if v.get() != ""]) #type:ignore
                wb.save(path)
            return wb, sheet
        elif file_extension == ".csv":
            return None, None

    def get_file_extension(self):
        # Λαμβάνουμε την επέκταση από το path του αρχείου
        file_path = self.output_path.get()
        if file_path.endswith(".xlsx"):
            return ".xlsx"
        elif file_path.endswith(".csv"):
            return ".csv"
        else:
            messagebox.showerror("Σφάλμα", "Ακατάλληλος τύπος αρχείου!")
            return None

    def connect_to_serial(self):
        port = self.ports_combobox.get()
        baudrate = self.baudrate.get()
        try:
            ser = serial.Serial(port, baudrate=baudrate, timeout=1)
            return ser
        except Exception as e:
            messagebox.showerror("Σφάλμα σύνδεσης", str(e))
            return None

    def start_logging(self):
        self.serial_port = self.connect_to_serial()
        if not self.serial_port:
            return

        #self.wb, self.sheet = self.setup_file()    #εδώ γίνεται apend στο .xlsx 
        self.stop_event.clear()

        record_thread = threading.Thread(target=self.record_data)
        record_thread.daemon = True
        record_thread.start()

        self.update_plot()

    def stop_logging(self):
        if self.serial_port:
            if self.serial_port.is_open:
                self.stop_event.set()
                self.serial_port.close()
            self.serial_port = None
            messagebox.showinfo("Τερματισμός", "Η καταγραφή δεδομένων ολοκληρώθηκε.")
        else:
            messagebox.showinfo("Τερματισμός", "Η καταγραφή είχε ήδη διακοπεί.")

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

    
    
    
    def update_plot(self):
        # 1. Διαβάζουμε όλα τα νέα δεδομένα από την ουρά
        #while not self.data_queue.empty():
        #    timestamp, numeric_values, raw_items = self.data_queue.get()
        #    self.times.append(len(self.times) + 1)
        #    self.values.append(numeric_values)
        #    self.data_listbox.insert(tk.END, f"{timestamp}: {', '.join(raw_items)}")
        #    self.data_listbox.see(tk.END)
        while not self.data_queue.empty():
            timestamp, numeric_values, raw_items = self.data_queue.get()
            self.times.append(len(self.times) + 1) # Παραμένει αύξοντας αριθμός για το γράφημα
            self.actual_timestamps.append(timestamp) # Αποθήκευση της ώρας για το Excel
            self.values.append(numeric_values)
            self.data_listbox.insert(tk.END, f"{timestamp}: {', '.join(raw_items)}")
            self.data_listbox.see(tk.END)

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
    
    
    
    
    
    
    
    def clear_data(self):
        # Επιβεβαίωση από τον χρήστη
        if messagebox.askyesno("Καθαρισμός", "Θέλετε σίγουρα να διαγράψετε το γράφημα και το ιστορικό καταγραφής;"):
            # Μηδενισμός λιστών δεδομένων
            self.times = []
            self.values = []
            self.actual_timestamps = [] # Καθαρισμός και εδώ
            
            # Καθαρισμός του Listbox
            self.data_listbox.delete(0, tk.END)
            
            # Καθαρισμός του γραφήματος
            self.ax.clear()
            self.ax.set_xlabel("Αριθμός μετρήσεων")
            self.ax.set_ylabel("Τιμή")
            self.canvas.draw()
            
            messagebox.showinfo("Καθαρισμός", "Τα δεδομένα καθαρίστηκαν.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialDataLogger(root)
    root.mainloop()
