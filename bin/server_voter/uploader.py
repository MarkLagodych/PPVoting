"""
Здесь есть все функции для прошивки и настройки сервера/пультов
"""

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

# --------------------------------------------------------

# Последовательный порт, на котором весит ESP8266
port = 'COM1'

def setPort(p):
    global port
    port = p

def getPort():
    return port

def upload(firmware_file):
    '''
    Загружает прошивку
    firmware_file -- название бинарного файла прошивки
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
TYPE_SERVER = 2 # Внешняя Wi-Fi сеть
TYPE_SERVER_AP = 3 # AccessPoint - создать собственную сеть

def translateSetting(s):
    keys = list(settings_dict.keys())
    assert keys.count(s) != 0, "'%s' is not a setting!" % s
    return bytes(settings_dict[s], encoding = 'ASCII')

ser = serial.Serial()

def beginSettings(freq = 115200):
    """
    Вводит модуль в режим настроек
    """
    if ser.is_open:
        ser.close()
    ser.baudrate = freq
    ser.port = port
    ser.timeout = 5 # сукунд

    ser.open()
    ser.reset_input_buffer()
    ser.reset_output_buffer()
    sleep(0.5)
    query(b's') # s -- Setting mode
        
def endSettings():
    """
    Выводит модуль из режима настроек
    """
    if ser.is_open:
        ser.write(b'e')    # End
        ser.close()

def query(q, wait = False):
    '''
    Отправляет q модулю 
    Если wait = True, ждёт ответа и выводит его
    '''
    assert ser.is_open, 'Serial port %s is not opened!' % port

    ser.write(q)

    if wait:
        sleep(0.5) # sec
        while ser.in_waiting == 0:
            sleep(0.5)
        while ser.in_waiting > 0:
            res = ser.read(size = ser.in_waiting)
            # backslashreplace -- заменить все non-ASCII символы на "\\xYZ"
            # где YZ -- шестнадцатиричный код
            print(res.decode(encoding='ASCII', errors='backslashreplace'))

            sleep(0.5)

def readSettings():
    '''
    Даёт команду модулю прочитать настройки из EEPROM в RAM
    '''
    query(b'r') # Read

def writeSettings():
    '''
    Даёт команду модулю записаить настройки из RAM в EEPROM
    '''
    query(b'w') # Write

def getSettings():
    '''
    Запросить и напечатать все настройки из RAM
    '''
    query(b'g', True) # Get

def setSetting(setting, value):
    '''
    Установить одну конкретную настройку в RAM
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
    Запускает device manager (менеджер устройств)
    Легко узнать номер COM-порта подключённого модуля
    '''
    os.system("devmgmt.msc")

if __name__ == '__main__':
    print("Can not be main. Must work in a shell or as a module.")
    
