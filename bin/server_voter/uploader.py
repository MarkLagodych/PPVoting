import sys
import re

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

# a word e.g.  a, b, abc, red54yij%4j+,q
# or an empty string e.g. ""
# or a double-quoted string (inner quotes can be escaped) e.g. "a", "\"", " a b "
token_re = re.compile(r'(\b\S{1,}\b)|""|(".*?[^\\]")')

def noQuotes(s):
    if s.startswith('"') and s.endswith('"'):
        s = s[1:-1]
    return s.replace('\\"', '"')

def parseLine(line):
    args = re.split(token_re, line)
    parsed = [noQuotes(arg) for arg in args if arg!=None and arg!='' and arg!='\n' and arg!=' ']
    return parsed


def printHelp(subject = 'all'):
    if subject in ['commands', 'all']:
        print("Commands (multiple commands can be separated by a semicolon ';'):")
        print("  exit    -- Exit program")
        print("  help    -- Show this message again")
        print("  help X  -- Show help message about subject X (variables/commands) ")
        print("  server  -- Upload server firmware")
        print("  voter   -- Upload voter firmware")
        print("  com X   -- Set COM port for uploading")
        print("  show    -- Show current COM port")
        print("  q       -- Query all local variables of a connected device")
        print("  q X     -- Query only X")
        print("  qs X Y  -- Query Set - set X to Y")
        print("  qrst    -- Query reset of all variables to the defaults")
        print("  qrst X  -- Query reset of X to its default")
        
    if subject == 'all':
        print()
        
    if subject in ['variables', 'all']:
        print("Variables (permanently-stored settings):")
        print("  ip      -- the last part in 192.168.1.x")
        print("  net     -- Wi-Fi network name (max. 32 characters)")
        print("  pwd     -- password for the network (max. 64 chars)")
        print("  port    -- Set server port")
        
    if not subject in ['variables', 'commands', 'all']:
        print("Unknown subject '%s', use 'variables' or 'commands' or 'all'" % subject)

# Globals
server_firmware = 'server.bin'
voter_firmware = 'voter.bin'
com_port = 1
        

print("PPVoting firmware uploading script. Type 'help' to see command list.")
print()

while True:
    try:
        line = input('@ ')

        singleLines = line.split(';')

        for singleLine in singleLines:
            tokens = parseLine(line)

            if len(tokens) == 0: continue

            print()
            
            command = tokens[0]

            if command == 'exit':
                sys.exit()
            elif command == 'help':
                if len(tokens) == 1:
                    printHelp()
                else:
                    printHelp(tokens[1])
            else:
                print("Unknown command: '%s'. Type 'help'." % command)

            print()
        
    except KeyboardInterrupt:
        sys.exit()
    
