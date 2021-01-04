###########################################################
# СКРИПТ СРАВНЕНИЯ И ЗАПИСИ ПРОШИВОК ДЛЯ ГОЛОСОВАЛОК      #
# ------------------------------------------------------- #
# Читает voter_2.bin и voter_3.bin, ищет различия         #
# побайтово. Бинарные файды должны быть                   #
# скомпилированными в Arduino IDE прошивками из src/voter #
# с только одним различием - последним байтом IP-адреса.  #
# Считается, что первое различие в байтах прошивок - это  #
# и есть те самые IP-адреса (остальные отличия хоть и     #
# возможны, но ни на что не влияют).                      #
###########################################################

import sys

try:
    import esptool
except:
    print("No esptool module found!")

# Change if need
port = 'COM7'

file1 = open('voter_2.bin', 'rb')
file2 = open('voter_3.bin', 'rb')

array1 = list(file1.read())
array2 = list(file2.read())

file1.close()
file2.close()

print('\n--------------- ANALIZING ---------------\n')

l1 = len(array1)
l2 = len(array2)
print('Lengths (bytes):')
print('    1st file: %d' % l1)
print('    2nd file: %d' % l2)

positions = []

if l1 == l2:
    print('    Equal: OK\n')
    
    i = 0
    while i < l1:
        if array1[i] != array2[i]:
            positions.append(i)
            print('Difference:')
            print('    position: %d' % i)
            print('    1st value: %d' % array1[i])
            print('    2nd value: %d' % array2[i])
        i += 1
    print('')

    
    print('\n--------------- PRODUCING ---------------\n')

    while True:
        ip = -1
        while ip < 2  or ip > 254:
            try:
                ip = int(input('IP: '))
            except:
                print('Bad IP')
        result = open('processed.bin', 'wb')
        array1[positions[0]] = ip
        result.write(bytes(array1))
        result.close()
        print('Wrote to result.bin')
        esptool.main(['--p',
                      port,
                      'erase_flash'])
        esptool.main(['--p',
                      port,
                      '--baud',
                      '460800',
                      'write_flash',
                      '--flash_size=detect',
                      '0',
                      'processed.bin'])
        


else:
    print('    FAIL\n')

input('\nExecution finished')

