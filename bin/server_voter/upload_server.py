from uploader import *

setPort('COM3')

print('Uploading firmware...')
upload('server.bin')

print()
print('Entering settings mode...')
beginSettings()

readSettings()

print('Current settings:')
getSettings()

print('Writing settings...')
setSetting('ssid',      'ArbUZ361')
setSetting('password',  'gumANOId1')
setSetting('ip',        [192, 168, 1, 1])
setSetting('subnet',    [255, 255, 255, 0])
setSetting('gateway',   [192, 168, 1, 1])
setSetting('server_ip', [192, 168, 1, 1])
setSetting('type',      TYPE_SERVER_AP)
writeSettings()

print('Current settings:')
readSettings()
getSettings()

endSettings()
