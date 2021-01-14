#include "settings.h"

Settings settings;

void Settings::start() {
    EEPROM.begin(512);
}


void Settings::printArray(byte *arr) {
    for (int i=0; i<4; i++) {
        Serial.print(arr[i], DEC);
        Serial.print('.');
    }
    Serial.println();
}

void Settings::settingsMode() {
    
    while (1) 
        if (Serial.available()) {
            
            char action = Serial.read();
            
            switch (action) {
              
                case 's': {  // Set

                    // Небольщой таймаут, чтобы загрузился следующий символ
                    for (int i=0; i<10; i++) {
                        if (Serial.available())
                            break;
                        delay(100);
                    }

                    if (!Serial.available()) {
                        Serial.println("Variable name is not given");
                        continue;
                    }

                    // Читаем символ -- название переменной для настройки
                    char var = Serial.read();
                    
                    switch (var) {

                        // Ip
                        case 'i':
                            Serial.readBytes(ip, 4);
                            break;

                        // subNet
                        case 'n':
                            Serial.readBytes(subnet, 4);
                            break;

                        // Gateway
                        case 'g': 
                            Serial.readBytes(gateway, 4);
                            break;

                        // server Ip
                        case 'I':
                            Serial.readBytes(server_ip, 4);
                            break;

                        // Type
                        case 't':
                            type = Serial.read();
                            break;  

                        // Ssid
                        case 's': {
                            size_t written = Serial.readBytesUntil(0, ssid, 31);
                            ssid[written] = 0;
                            break;
                        }

                        // Password
                        case 'p': {
                            size_t written = Serial.readBytesUntil(0, password, 63);
                            password[written] = 0;
                            break;
                        }

                        default:
                            Serial.print("Unknown variable ");
                            Serial.print(var, DEC);
                            Serial.print(" \"");
                            Serial.print(var);
                            Serial.println("\"");
                            break;
                        
                    } // switch (var)
                    
                    break;
                    
                } // case 's'  -- action

                case 'w':   // Write
                    save();
                    break;

                case 'g':   // Get
                    Serial.print("Type: ");
                    
                    switch (type) {
                        case VoterDevice:
                            Serial.print("[voter] ");
                            break;
                            
                        case ServerDevice:
                            Serial.print("[server] ");
                            break;
                            
                        case ServerAPDevice:
                            Serial.print("[server, AccessPoint] ");
                            break;
                            
                        default:
                            Serial.print("[unknown] ");
                            break;
                    }
                    
                    Serial.println(type, DEC);
                    
                    Serial.print("SSID: ");
                    Serial.println(ssid);

                    Serial.print("Password: ");
                    Serial.println(password);

                    Serial.print("IP: ");
                    printArray(ip);

                    Serial.print("Server IP: ");
                    if (type != VoterDevice);
                        Serial.print("[unused] ");
                    printArray(server_ip);

                    Serial.print("Gateway: ");
                    printArray(gateway);

                    Serial.print("Subnet: ");
                    printArray(subnet);
                    break;

                case 'r':   // Read
                    load();
                    break;

                case 'e':  // End
                    return;

                default: {
                    Serial.print("Unknown action ");
                    Serial.print(action, DEC);
                    Serial.print(" \"");
                    Serial.print(action);
                    Serial.println("\"");
                    break;
                }
            } // switch (action)
            
        } // if (Serial.available())
}

void Settings::load() {
    int addr = 0;

#define NEXT EEPROM.read(addr++)
    
    type = NEXT;

    for (int i=0; i<4; i++)
        ip[i] = NEXT;

    for (int i=0; i<4; i++)
        subnet[i] = NEXT;

    for (int i=0; i<4; i++)
        gateway[i] = NEXT;

  for (int i=0; i<4; i++)
        server_ip[i] = NEXT;
    
    for (int i=0; i<31; i++) 
        ssid[i] = NEXT;
    ssid[31] = 0;

    for (int i=0; i<63; i++)
        password[i] = NEXT;
    password[63] = 0;

#undef NEXT
    
}

void Settings::save() {
    int addr = 0;

#define NEXT(x) EEPROM.write(addr++, (x));
    
    NEXT(type);

    for (int i=0; i<4; i++)
        NEXT(ip[i]);

    for (int i=0; i<4; i++)
        NEXT(subnet[i]);

    for (int i=0; i<4; i++)
        NEXT(gateway[i]);

    for (int i=0; i<4; i++)
        NEXT(server_ip[i]);

    for (int i=0; i<31; i++)
        NEXT(ssid[i]);    

    for (int i=0; i<63; i++)
        NEXT(password[i]);

    EEPROM.commit();

#undef NEXT

}