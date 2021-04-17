from common import *

setPort('COM4')

print('Uploading firmware...')
upload('voter.bin')

print()
print('Entering settings mode...')
beginSettings()
readSettings()

print('Current settings:')
getSettings()

print('Writing settings...')
setSetting('ssid',      'ArbUZ361')
setSetting('password',  'gumANOId1')
setSetting('ip',        [192, 168, 1, 2])
setSetting('subnet',    [255, 255, 255, 0])
setSetting('gateway',   [192, 168, 1, 1])
setSetting('server_ip', [192, 168, 1, 1])
setSetting('type',      TYPE_VOTER)
writeSettings()

print('Current settings:')
readSettings()
getSettings()

print('Leaving settings mode...')
endSettings()

print('Done!!!')
