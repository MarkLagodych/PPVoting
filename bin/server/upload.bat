py -m esptool --p COM3 erase_flash
py -m esptool --p COM3 --baud 460800 write_flash --flash_size=detect 0 server.ino.generic.bin