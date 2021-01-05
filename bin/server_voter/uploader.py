import sys
import os

try:
    import serial
except:
    print("Error!")
    print("No serial module available. Please install it with")
    print("\t[py[thon] -m] pip install pyserial")
    sys.exit(1)

try:
    import esptool
except:
    print("Error!")
    print("No esptool module available. Please install it with")
    print("\t[py[thon] -m] pip install esptool")
    sys.exit(2)

# -----------------------  Globals -----------------------

# Serial port for uploading firmware / configuring device
port = 'COM3'

# --------------- Functions that you can use -------------

def upload(firmware_file):
    '''
    firmware_file should be a file name of a binary that was exported from Arduino IDE
    '''
    
    esptool.main([
        '--p',
        port,
        'erase_flash'
    ])
    
    esptool.main([
        '--p',
        port,
        '--baud',
        '460800',
        'write_flash',
        '--flash_size=detect',
        '0',
        firmware_file
    ])


def query(setting = 'all'):
    '''
    Query the setting (or all of them) of a connected device
    '''
    ser = serial.Serial(port, 115200)
    if not ser.is_open:
        ser.open()
    ser.write(b'g')
    respond = ser.readline()
    print(respond)
    ser.close()
    

def devices():
    '''
    Starts device manager
    Handy to find out which serial port to use
    '''
    os.system("devmgmt.msc")

if __name__ == '__main__':
    print("Can not be main. Must work in a shell or as a module.")
    
