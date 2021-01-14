import sys
import os
import io
from time import sleep

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
port = 'COM1'

def setPort(p):
    global port
    port = p

def getPort():
    return port

# --------------- Functions that you can use -------------

def upload(firmware_file):
    '''
    firmware_file should be a file name of a binary
    that was exported from Arduino IDE
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



settings_dict = {
    'ssid':        's',
    'password':    'p',
    
    'ip':          'i',
    'subnet':      'n',
    'gateway':     'g',
    
    'server_ip':   'I',
    
    'type':        't'
}

TYPE_VOTER = 1
TYPE_SERVER = 2 # External Wi-Fi network
TYPE_SERVER_AP = 3 # AccessPoint - create its own Wi-Fi network

def translateSetting(s):
    keys = list(settings_dict.keys())
    assert keys.count(s) != 0, "'%s' is not a setting!" % s
    return bytes(settings_dict[s], encoding = 'ASCII')

ser = serial.Serial()

def beginSettings(freq = 115200):
    if ser.is_open:
        ser.close()
    ser.baudrate = freq
    ser.port = port
    ser.timeout = 5 # seconds

    ser.open()
    ser.reset_input_buffer()
    sleep(0.5)
    query(b's') # s -- Setting mode
        
def endSettings():
    if ser.is_open:
        ser.write(b'e')    # End
        ser.close()

def query(q, wait = False):
    '''If wait is True, waits for respond and prints it'''
    assert ser.is_open, 'Serial port %s is not opened!' % port

    ser.write(q)

    if wait:
        sleep(0.5) # sec
        while ser.in_waiting == 0:
            sleep(0.5)
        while ser.in_waiting > 0:
            res = ser.read(size = ser.in_waiting)
            # backslashreplace -- replace all non-ASCII symbols with "\\xYZ"
            # where YZ is a hex code
            print(res.decode(encoding='ASCII', errors='backslashreplace'))

            sleep(0.5)

def readSettings():
    '''
    Read settings from EEPROM to RAM
    '''
    query(b'r') # Read

def writeSettings():
    '''
    Write settings from RAM to EEPROM
    '''
    query(b'w') # Write

def getSettings():
    '''
    Query all the settings of a connected device
    '''
    query(b'g', True) # Get

def setSetting(setting, value):
    '''
    Set the setting (or all of them) of a connected device
    '''

    if type(value) != bytes:
        if type(value) == str:
            value = bytes(value, encoding = 'ASCII') + bytes([0])
        elif type(value) in [int, bool]:
            value = bytes([value])
        else:
            value = bytes(value)

    # s - Set
    q = b's' + translateSetting(setting) + value
    query(q)
    

def devices():
    '''
    Starts device manager
    Handy to find out which serial port to use
    '''
    os.system("devmgmt.msc")

if __name__ == '__main__':
    print("Can not be main. Must work in a shell or as a module.")
    
