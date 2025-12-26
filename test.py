import serial
import serial.tools.list_ports as list_ports

brate = 9600
port = '/dev/ttyACM0'

print([port.device for port in list_ports.comports()])

ser = serial.Serial(port, brate, timeout=1)

def convert(x):
    try:
        return float(x)
    except:
        return -1.0

while True:
    line = ser.readline().decode('utf-8', errors='ignore').strip()
    if line:
        print([convert(val) for val in line.split(',')])

