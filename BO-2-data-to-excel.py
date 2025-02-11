#Vassilis Economou 16/01/2025 v.02

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

# Απενεργοποίηση προειδοποιήσεων από matplotlib
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

class SerialDataLogger:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Data Logger")

        # Προσθήκη εικονιδίου και τίτλου
        #self.root.iconbitmap("icon.ico")  # Αντικαταστήστε με το όνομα του αρχείου εικονιδίου
        title_label = ttk.Label(self.root, text="Serial Data Logger  [Βασίλης Οικονόμου v.02]", font=("Arial", 15, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # Αρχικοποίηση μεταβλητών
        self.serial_port = None
        self.baudrate = tk.IntVar(value=9600)
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "BO_SDL.xlsx"))
        self.times = []
        self.values = []
        self.data_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.sampling_rate = tk.IntVar(value=0)  # Ταχύτητα δειγματοληψίας σε ms

        # Επιλογή ThingSpeak
        self.send_to_thingspeak = tk.BooleanVar(value=False)
        self.thingspeak_api_key = tk.StringVar(value="8Q9GSIRNOAP2FXDY")  # Αρχικό API Key

        self.create_widgets()

    # Προσθήκη της συνάρτησης για το παράθυρο οδηγών
    def open_instructions_window(self):
        # Δημιουργία νέου παραθύρου
        instructions_window = tk.Toplevel(self.root)
        instructions_window.title("Οδηγίες")
        instructions_window.geometry("500x600")

        # Εισαγωγή κειμένου με οδηγίες
        instructions_text = (
            "Καταγραφή δεδομένων από serial (Serial Data Logger).\n\n\n"
            "Μπορείτε να:\n\n" 
            "1. Eπιλέξετε τη θύρα από την οποία θα διαβάσετε δεδομένα.\n"
            "    (με [Aνανέωση] διαβάζονται ξανά οι διαθέσιμες θύρες)\n\n"
            "2. Ορίσετε το Baudrate για τη σύνδεση.\n"
            "    (η τιμή που προτείνεται είναι αρκετή)\n\n"
            "3. Επιλέξετε αν οι μετρήσεις θα εξάγονται στο ThinkSpeeak.\n"
            "    (οπότε θα χρειαστεί να oρίσετε και το API Key)\n\n"
            "4. Επιλέξετε την καθυστέρηση μεταξύ των δειγματοληψιών\n\n\n\n"
            "Πρέπει να:\n\n"
            "Επιλέξετε το όνομα του αρχείου και τον τύπο του (.xlsx ή .csv), για αποθήκευση.\n\n\n"
            "_______________________\n\n"
            "Πατήστε [Έναρξη καταγραφής] για να ξεκινήσετε τη καταγραφή.\n\n"
            "Πατήστε [Τερματισμός] για να σταματήσετε την καταγραφή.\n\n"
            "Πατήστε [Αποθήκευση] για να αποθηκεύσετε τα δεδομένα.\n"
            "    (μπορείτε να αποθηκεύετε και πριν τον τερματισμό τιμές στο αρχείο\n"
            "    ...όσες φορές θέλετε/χρειαστεί)\n\n\n\n"
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
        instructions_button.grid(row=1, column=2, pady=3)
        # Επιλογή θύρας
        ports_label = ttk.Label(self.root, text="Θα διαβάσω από τη Θύρα:")
        ports_label.grid(row=2, column=0, padx=5, pady=3, sticky="w")
        self.ports_combobox = ttk.Combobox(self.root, state="readonly")
        self.ports_combobox.grid(row=2, column=1, padx=5, pady=3, sticky="ew")
        self.refresh_ports()
        refresh_button = ttk.Button(self.root, text="Ανανέωση", command=self.refresh_ports)
        refresh_button.grid(row=2, column=2, padx=5, pady=3)

        # Επιλογή baudrate
        baudrate_label = ttk.Label(self.root, text="...με ρυθμό (Baudrate):")
        baudrate_label.grid(row=3, column=0, padx=5, pady=3, sticky="w")
        baudrate_combobox = ttk.Combobox(self.root, textvariable=self.baudrate, state="readonly")
        baudrate_combobox["values"] = [9600, 19200, 38400, 57600, 115200]
        baudrate_combobox.grid(row=3, column=1, padx=5, pady=3, sticky="ew")

        # Επιλογή τοποθεσίας εξόδου
        output_label = ttk.Label(self.root, text="Θα αποθηκεύσω στο αρχείο (.xlsx ή .csv):")
        output_label.grid(row=4, column=0, padx=5, pady=3, sticky="w")
        output_entry = ttk.Entry(self.root, textvariable=self.output_path)
        output_entry.grid(row=4, column=1, padx=0, pady=3, sticky="ew")
        browse_button = ttk.Button(self.root, text="Επιλογή άλλου", command=self.browse_file)
        browse_button.grid(row=4, column=2, padx=5, pady=3)

        # Επιλογή για ThingSpeak
        thingspeak_check = ttk.Checkbutton(self.root, text="Αποστολή τιμής και στο ThingSpeak με API Key:", variable=self.send_to_thingspeak)
        thingspeak_check.grid(row=5, column=0, padx=5, pady=3, sticky="w")
        api_key_entry = ttk.Entry(self.root, textvariable=self.thingspeak_api_key)
        api_key_entry.grid(row=5, column=1, padx=5, pady=3, sticky="ew")

        # Slider για την ταχύτητα δειγματοληψίας
        sampling_rate_label = ttk.Label(self.root, text="Καθυστέρηση μεταξύ των δειγματοληψιών (ms):")
        sampling_rate_label.grid(row=6, column=0, padx=5, pady=3, sticky="w")
        self.sampling_rate_slider = ttk.Scale(self.root, from_=0, to=5000, variable=self.sampling_rate, orient=tk.HORIZONTAL)
        self.sampling_rate_slider.grid(row=6, column=1, padx=5, pady=3, sticky="ew")
        self.sampling_rate_value_label = ttk.Label(self.root, text=f"{self.sampling_rate.get()} ms")
        self.sampling_rate_value_label.grid(row=6, column=2, padx=5, pady=3, sticky="w")
        self.sampling_rate_slider.config(command=self.update_sampling_rate_label)

        # Κουμπιά έναρξης, τερματισμού και αποθήκευσης
        start_button = ttk.Button(self.root, text="Έναρξη καταγραφής", command=self.start_logging)
        start_button.grid(row=7, column=0, pady=3)
        stop_button = ttk.Button(self.root, text="Τερματισμός καταγραφής", command=self.stop_logging)
        stop_button.grid(row=7, column=1, pady=3)
        save_button = ttk.Button(self.root, text="Αποθήκευση", command=self.save_data)
        save_button.grid(row=7, column=2, pady=3)

        # Περιοχή εμφάνισης δεδομένων
        self.data_listbox = tk.Listbox(self.root, height=10)
        self.data_listbox.grid(row=8, column=0, columnspan=3, padx=5, pady=3, sticky="nsew")

        # Διάγραμμα
        figure = Figure(figsize=(7, 4), dpi=100)
        self.ax = figure.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(figure, master=self.root)
        self.canvas.get_tk_widget().grid(row=9, column=0, columnspan=3, pady=3, sticky="nsew")

        # Ρύθμιση διαστάσεων πλέγματος
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(8, weight=1)

    def update_sampling_rate_label(self, value):
        self.sampling_rate_value_label.config(text=f"{int(float(value))} ms")

    def send_to_thingspeak_api(self, value):
        if self.send_to_thingspeak.get():  # Έλεγχος αν το checkbox είναι ενεργοποιημένο
            try:
                api_key = self.thingspeak_api_key.get()
                if not api_key:
                    messagebox.showwarning("API Key", "Παρακαλώ εισάγετε το API Key.")
                    return
                url = f"https://api.thingspeak.com/update?api_key={api_key}"
                payload = {'field1': value}
                response = requests.get(url, params=payload)
                if response.status_code == 200:
                    print(value)
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

    def save_data(self):
        file_extension = self.get_file_extension()
        if file_extension == ".xlsx":
            try:
                self.wb.save(self.output_path.get())
                messagebox.showinfo("Αποθήκευση", "Τα δεδομένα αποθηκεύτηκαν με επιτυχία σε .xlsx.")
            except Exception as e:
                messagebox.showerror("Σφάλμα αποθήκευσης", str(e))
        elif file_extension == ".csv":
            try:
                with open(self.output_path.get(), mode="w", newline="", encoding="utf-8") as file:
                    writer = csv.writer(file)
                    writer.writerow(["Time", "Value"])
                    writer.writerows(zip(self.times, self.values))
                messagebox.showinfo("Αποθήκευση", "Τα δεδομένα αποθηκεύτηκαν με επιτυχία σε .csv.")
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
                sheet.append(["Time", "Value"])
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

        self.wb, self.sheet = self.setup_file()
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
                    
                    # Ελέγχουμε αν η γραμμή περιέχει μόνο αριθμούς
                    if line.replace('.', '', 1).isdigit():  # Ελέγχει αν είναι αριθμός
                        value = float(line)  # Αν είναι αριθμός, τον αποθηκεύουμε ως float
                    else:
                        value = line  # Αν περιέχει γράμματα, το κρατάμε ως κείμενο
                    
                    # Αν το αρχείο εξόδου είναι .xlsx, καταγράφουμε τη γραμμή
                    if self.get_file_extension() == ".xlsx":
                        self.sheet.append([timestamp, value])
                    
                    # Βάζουμε τα δεδομένα στην ουρά για την απεικόνιση
                    self.data_queue.put((timestamp, value))

                    # Στέλνουμε τα δεδομένα στο ThingSpeak
                    self.send_to_thingspeak_api(value)
                    
                    # Προσθήκη καθυστέρησης ανάλογα με την ταχύτητα δειγματοληψίας
                    threading.Event().wait(self.sampling_rate.get() / 1000)
        except Exception as e:
            if not self.stop_event.is_set():
                messagebox.showerror("Σφάλμα καταγραφής", str(e))

    def update_plot(self):
        while not self.data_queue.empty():
            timestamp, value = self.data_queue.get()
            self.times.append(len(self.times) + 1)
            self.values.append(value)

            self.data_listbox.insert(tk.END, f"{timestamp}: {value}")
            self.data_listbox.see(tk.END)

        self.ax.clear()
        self.ax.plot(self.times, self.values, label="Μέτρηση")
        self.ax.set_xlabel("Αριθμός μετρήσεων")
        self.ax.set_ylabel("Μέτρηση")
        self.ax.legend()
        self.canvas.draw()


        if not self.stop_event.is_set():
            # Μειώνουμε την καθυστέρηση για πιο συχνή ανανέωση
            self.root.after(50, self.update_plot)

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialDataLogger(root)
    root.mainloop()
