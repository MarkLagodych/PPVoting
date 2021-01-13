from uploader import *

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
setSetting('type',      3)
writeSettings()

print('Current settings:')
readSettings()
getSettings()

endSettings()
