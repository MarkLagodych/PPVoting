import sys
import os
import io

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

settings = {
    'all':         'a',
    
    'ssid':        's',
    'password':    'p',
    
    'ip':          'i',
    'full_ip':     'f',
    'subnet':      'b',
    'gateway':     'g',
    
    'server_ip':   'v',
    'server_port': 'o',
    'AP_mode':     'm', # Access Point mode (autonomic)
    
    'type':        't'
}

def translateSetting(s):
    keys = list(settings.keys())
    assert keys.count(s) != 0, "'%s' is not a setting!" % s
    return bytes(settings[s], encoding = 'ASCII')

ser = serial.Serial()

def queryBegin(freq = 115200):
    if ser.is_open():
        ser.close()
    ser.baudrate = freq
    ser.port = port
    ser.open()
    ser.write(b'q')        # Query mode
    print(ser.readline())

def queryEnd():
    if ser.is_open():
        ser.write(b'e')    # End
        ser.close()

def queryRaw(q):
    if not ser.is_open():
        print('Serial port %s is not opened!' % port)
        return None
    ser.write(q)
    return ser.readline()

def query(setting = 'all'):
    '''
    Query the setting (or all of them) of a connected device
    '''
        
    query = b'g' + translateSetting(setting) # Get
    print(queryRaw(query))

def querySet(setting, value):
    '''
    Set the setting (or all of them) of a connected device
    '''

    if type(value) != bytes:
        if type(value) == str:
            value = bytes(value, encoding = 'ASCII')
        elif type(value) in [int, bool]:
            value = bytes([value])
        else:
            value = bytes(value)
            
    query = b's' + translateSetting(setting) + value # Set
    print(queryRaw(query))

def queryReset(setting = 'all'):
    '''
    Reset the setting (or all of them) of a connected device
    '''
    ser = serial.Serial(port, 115200)
    if not ser.is_open:
        ser.open()

    query = b'r' + translateSetting(setting) # Reset
    print(queryRaw(query))
    

def devices():
    '''
    Starts device manager
    Handy to find out which serial port to use
    '''
    os.system("devmgmt.msc")

if __name__ == '__main__':
    print("Can not be main. Must work in a shell or as a module.")
    
