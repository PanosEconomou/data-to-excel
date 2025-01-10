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
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import threading
import warnings
import os

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

# Επιλογή χρόνου για αποθήκευση
def pick_time_save():
    questions = [
        inquirer.List('time_save',
                      message="Κάθε πόσα λεπτά να αποθηκεύονται οι μετρήσεις στ αρχείο;",
                      choices=[1, 5, 10, 30, 60, 90],
                      default=5
                      ),
    ]
    answers = inquirer.prompt(questions)
    return int(answers['time_save'])


# Επιλογή τοποθεσίας εξόδου
def pick_output_location():
    questions = [inquirer.Path('file',
                    message='Αρχείο καταγραφής',
                    default=os.path.join(os.getcwd(),"data_from_serial.xlsx"),
                    path_type=inquirer.Path.FILE
                ),]
    path_selected = inquirer.prompt(questions)
    path = os.path.abspath(os.path.expanduser(path_selected['file']))
    if not path.endswith('.xlsx'):
        if os.path.isdir(path):
            path = os.path.join(path,"data_from_serial.xlsx") 
        else:
            path += '.xlsx'
    if not os.path.exists(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    return path

# Δημιουργία ή Άνοιγμα Αρχείου Excel
def setup_excel(path):
    try:
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        print(f"Το αρχείο Excel '{path}' βρέθηκε και άνοιξε.")
    except (FileNotFoundError, KeyError):
        wb = Workbook()
        sheet = wb.active
        sheet.append(["Time", "Value"])  # Επικεφαλίδες στηλών
        wb.save(path)
        print(f"Δημιουργήθηκε νέο αρχείο Excel: '{path}'.")
    return wb, sheet

# Σύνδεση στη σειριακή θύρα
def connect_to_serial(port, baudrate):
    try:
        ser = serial.Serial(port.device, baudrate=baudrate, timeout=1)
        print("Σύνδεση επιτυχής στη θύρα!")
        return ser
    except Exception as e:
        print(f"Αδυναμία σύνδεσης στη θύρα! Λεπτομέρειες: {e}")
        sys.exit()

# Επικύρωση δεδομένων πριν την καταγραφή
def validate_and_save_data(sheet, data):
    if data:
        sheet.append([datetime.now().strftime('%Y-%m-%d %H:%M:%S'), data])
    else:
        print("Δεν ελήφθησαν δεδομένα.")

# Αποθήκευση αρχείου περιοδικά
def save_periodically(wb, path):
    try:
        wb.save(path)
        print(f"Το αρχείο '{path}' αποθηκεύτηκε.")
    except Exception as e:
        print(f"Σφάλμα κατά την αποθήκευση του αρχείου '{path}': {e}")

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
def record_data(ser, sheet, wb, path, times, values, saving_time, stop_event):
    last_save_time = time.time()
    print("\nΠατήστε Ctrl + C για να σταματήσετε την καταγραφή.")
    try:
        while not stop_event.is_set():
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line:
                validate_and_save_data(sheet, line)
                times.append(time.time())
                values.append(line)
                if time.time() - last_save_time >= saving_time * 60:
                    save_periodically(wb, path)
                    last_save_time = time.time()
                print(f"Τρέχουσα μέτρηση: {line}")
    except KeyboardInterrupt:
        print("\nΔιακοπή από τον χρήστη.")
    except Exception as e:
        print(f"Σφάλμα: {e}")
    finally:
        save_periodically(wb, path)
        print(f"Ευχαριστούμε για τη χρήση της εφαρμογής αυτής.")

# Κύριο πρόγραμμα
if __name__ == "__main__":
    print("\n\n\nData From Serial to .xlsx (Vassilis Economou v.1.0.0)")
    print("______________________________________________________________\n")

    port = pick_port()
    baudrate = pick_baud_rate()
    saving_time = pick_time_save()
    path = pick_output_location()
    wb, sheet = setup_excel(path)
    ser = connect_to_serial(port, baudrate)


    times = []
    values = []

    fig, ax = plt.subplots()
    ani = FuncAnimation(fig, animate, fargs=(times, values, ax), interval=100)

    stop_event = threading.Event()
    record_thread = threading.Thread(target=record_data, args=(ser, sheet, wb, path, times, values, saving_time, stop_event))
    record_thread.daemon = True
    record_thread.start()

    try:
        plt.show()
    except KeyboardInterrupt:
        save_periodically(wb, path)
        print("_____________")
        print("Ευχαριστούμε.")
        stop_event.set()
