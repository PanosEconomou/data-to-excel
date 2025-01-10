# Data From Serial to .xlsx (Vassilis Economou v.1.0.0)
# 2025-01-09

import openpyxl
from openpyxl import Workbook
import serial
import inquirer
import serial.tools.list_ports as list_ports
from datetime import datetime
import sys
import time
from tqdm import tqdm
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import threading
import warnings
import time


# Απενεργοποίηση προειδοποιήσεων από matplotlib
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

# Επιλογή θύρας σειριακής σύνδεσης
def pick_port():
    ports = list_ports.comports()
    if len(ports) == 0:
        print("Δεν βρέθηκαν διαθέσιμες σειριακές θύρες.")
        exit()
    
    questions = [
        inquirer.List('port',
                      message="Επιλέξτε θύρα σειριακής σύνδεσης",
                      choices=[port.device for port in ports],
                      ),
    ]
    answers = inquirer.prompt(questions)
    return [port for port in ports if port.device == answers['port']][0]

# Επιλογή baudrate
def pick_baud_rate():
    questions = [
        inquirer.List('baudrate',
                      message="Επιλέξτε baudrate",
                      choices=[9600, 19200, 38400, 57600, 115200],
                      default=9600,
                      ),
    ]
    answers = inquirer.prompt(questions)
    return answers['baudrate']

# Εκτιμώμενος αριθμός εγγραφών
def pick_record_number():
    questions = [
        inquirer.List('records',
                      message="Επιλέξτε εκτιμώμενο πλήθος εγγραφών",
                      choices=[5000, 10000, 50000, 100000, 200000, 500000],
                      default=10000
                      ),
    ]
    rate_selected = inquirer.prompt(questions)
    return rate_selected['records']

# Επιλογή χρόνου για αποθήκευση
def pick_time_save():
    questions = [
        inquirer.List('time_save',
                      message="Κάθε πόσα λεπτά να αποθηκεύεται το αρχείο;",
                      choices=[1, 5, 10, 30, 60, 90],
                      default=5
                      ),
    ]
    answers = inquirer.prompt(questions)
    return int(answers['time_save'])


# Σύνδεση στη σειριακή θύρα
def connect_to_serial(port, baudrate):
    try:
        ser = serial.Serial(port.device, baudrate=baudrate, timeout=1)
        print("Σύνδεση επιτυχής στη θύρα!")
        return ser
    except Exception as e:
        print(f"Αδυναμία σύνδεσης στη θύρα! Λεπτομέρειες: {e}")
        sys.exit()

# Δημιουργία ή Άνοιγμα Αρχείου Excel
def setup_excel():
    try:
        wb = openpyxl.load_workbook('data_from_serial.xlsx')
        sheet = wb.active
        print("Το αρχείο Excel βρέθηκε και άνοιξε.")
    except (FileNotFoundError, KeyError):
        wb = Workbook()
        sheet = wb.active
        sheet.append(["Time", "Value"])  # Επικεφαλίδες στηλών
        wb.save('data_from_serial.xlsx')
        print("Δημιουργήθηκε νέο αρχείο Excel.")
    return wb, sheet

# Επικύρωση δεδομένων πριν την καταγραφή
def validate_and_save_data(sheet, data, wb):
    try:
        if data:
            sheet.append([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), data])
        else:
            print("Δεν ελήφθησαν δεδομένα.")
    except Exception as e:
        print(f"Σφάλμα κατά την καταγραφή δεδομένων: {e}")

# Αποθήκευση αρχείου περιοδικά
def save_periodically(wb):
    wb.save('data_from_serial.xlsx')

# Δημιουργία γραφήματος σε πραγματικό χρόνο
def animate(i, times, values, ax):
    ax.clear()
    ax.plot(times, values, label="Τιμή καταγραφής")
    ax.set_xlabel("Χρόνος (δευτερόλεπτα)")
    ax.set_ylabel("Τιμή")
    ax.set_title("Διάγραμμα Καταγραφής Δεδομένων")
    plt.xticks(rotation=45)
    ax.legend()

# Καταγραφή δεδομένων σε ξεχωριστό νήμα
def record_data(ser, sheet, wb, times, values, saving_time, stop_event):
    last_save_time = time.time()  # Ξεκινάμε το χρονόμετρο από την αρχή
    records_saved = 0
    progress_bar = tqdm(total=records, desc="Ποσοστό καταγραφής", ncols=150, position=1)
    last_display_time = time.time()

    print("\nΠατήστε Ctrl + C για να σταματήσετε την καταγραφή.")

    try:
        while not stop_event.is_set():  # Συνεχίζει μέχρι να λάβει σήμα για να σταματήσει
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line:  # Αν δεν είναι κενό το data
                validate_and_save_data(sheet, line, wb)
                
                # Πρόσθεση δεδομένων για το διάγραμμα
                times.append(time.time())
                values.append(line)
                
                # Αποθήκευση περιοδικά κάθε 20 εγγραφές
                #if records_saved >= 20:
                #    save_periodically(wb)
                #    records_saved = 0
                #else:
                #    records_saved += 1

                #Αποθήκευση κάθε X δευτερόλεπτα
                if time.time() - last_save_time >= saving_time * 60:
                    save_periodically(wb)  # Αποθήκευση
                    last_save_time = time.time()  # Ενημέρωση του χρόνου τελευταίας αποθήκευσης
                
                # Εμφάνιση της τρέχουσας μέτρησης στην κονσόλα κάθε 0.5 δευτερόλεπτο
                if time.time() - last_display_time > 0.5:
                    sys.stdout.write(f"\rΤρέχουσα μέτρηση: {line}")
                    sys.stdout.flush()
                    last_display_time = time.time()      
                
                progress_bar.update(1)  # Ενημέρωση μπάρας προόδου
    except KeyboardInterrupt:
        print("\nΔιακοπή από τον χρήστη.")
    except Exception as e:
        print(f"Σφάλμα: {e}")
    finally:
        # Εξασφαλίζουμε ότι η αποθήκευση θα ολοκληρωθεί
        print("\nΑποθήκευση τελευταίων δεδομένων πριν τον τερματισμό...")
        save_periodically(wb)  # Διασφαλίζουμε ότι το τελευταίο αρχείο αποθηκεύεται
        print("\nΤα δεδομένα αποθηκεύτηκαν στο αρχείο: data_from_serial.xlsx (στο φάκελο που βρίσκεται και η εφαρμογή αυτή). ")
        print("Αν δεν διαγράψετε και δεν μετακινήσετε το αρχείο αυτό, η επόμενη καταγραφή θα συνεχιστεί από εκεί που σταμάτησε η τρέχουσα")


# Κύριο πρόγραμμα
if __name__ == "__main__":
    print("\n\n\nData From Serial to .xlsx (Vassilis Economou v.1.0.0)\n")

    port = pick_port()
    baudrate = pick_baud_rate()
    records = pick_record_number()
    saving_time = pick_time_save()
    ser = connect_to_serial(port, baudrate)
    wb, sheet = setup_excel()

    # Στοιχεία για το διάγραμμα
    times = []
    values = []

    # Δημιουργία του διαγράμματος
    fig, ax = plt.subplots()
    ani = FuncAnimation(fig, animate, fargs=(times, values, ax), interval=100)  # Αυξήσαμε το διάστημα ανανέωσης

    # Δημιουργία event για τον συγχρονισμό των νημάτων
    stop_event = threading.Event()

    # Ξεκινάμε τη διαδικασία καταγραφής δεδομένων σε νέο νήμα
    record_thread = threading.Thread(target=record_data, args=(ser, sheet, wb, times, values, saving_time, stop_event))
    record_thread.daemon = True
    record_thread.start()

    try:
        plt.show()  # Εμφανίζει το διάγραμμα και περιμένει για νέες ενημερώσεις
    except KeyboardInterrupt:
        stop_event.set()  # Σήμα για να σταματήσει το νήμα καταγραφής
        print("\nΤο πρόγραμμα σταμάτησε.\n")
