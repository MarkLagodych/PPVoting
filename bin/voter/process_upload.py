# Reads 2voter.ino.[platform].bin
# and   voter.ino.[platform].bin,
# finds a byte that is different.
#
# Modifying that byte in .bin file
# is faster than modifying and compiling 
# .ino files.
#########################################################

import sys
import esptool

# Change if need
platform = 'generic'
port = 'COM7'

file1 = open('2voter.ino.%s.bin' % platform, 'rb')
file2 = open('voter.ino.%s.bin' % platform, 'rb')

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

    if len(positions) == 2:
        print('Count: 2 OK')
        print('\n--------------- PRODUCING ---------------\n')

        while True:
            ip = -1
            while ip < 2  or ip > 254:
                try:
                    ip = int(input('IP: '))
                except:
                    print('Bad IP')
            result = open('result.bin', 'wb')
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
                          'result.bin'])
        
    else:
        print('Count: %d FAIL' % len(positions))

else:
    print('    FAIL\n')

input('\nExecution finished')

