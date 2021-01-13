/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *                        СЕРВЕР ДЛЯ ГОЛОСОВАЛОК                       *
 * ------------------------------------------------------------------- *
 *                                                                     *
 * - Эта программа должна запускаться на ESP8266.                      *
 *   Задачи программы:                                                 *
 *       0. хранить настройки в (эмулированной) EEPROM;                *
 *       1. получать данные от пультов слушателей по WiFi;             *
 *       2. передавать данные ПК докладчика.                           *
 *                                                                     *
 * - Передача данных ПК возможна при помощи драйвера CH341SER.         *
 *   Его можно скачать с оф. сайта разработчика:                       *
 *       http://www.wch.cn/downloads/CH341SER_EXE.html                 *
 *   или взять из папки PPVoting/bin/driver                            *
 *                                                                     *
 * - Сеть WiFi может быть создана самим ESP8266 (макс. 8 пультов) или  *
 * любым другим устройством.                                           *
 * ------------------------------------------------------------------- *
 *                                                                     *
 *                            КАК РАБОТАЕТ КОД                         *
 * ------------------------------------------------------------------- *
 *                           * Общение по WiFi *                       *
 * 0. Ожидание настроек: если сразу после включения пришло "s",        *
 *    войти в режим настройки, если нет, продолжить работу             *
 * 1. Поключиться к / Создать сеть WiFi.                               *
 * 2. Создать веб-сервер.                                              *
 * 3. При каждом запросе веб-страницы вида /Х, где Х - число от 0 до 6 * 
 *    записать голос за Х-й вариант.                                   *
 * 4. Если по последовательному порту пришло "c", очистить все ответы, * 
 *    а если "g", то выслать JSON-строку с результатом                 *
 *                                                                     *
 *                              РЕЖИМ НАСТРОЕК                         *
 * ------------------------------------------------------------------- *
 * Формат общения в режиме настроек такой: A[VD...]                    *
 * A (action) - действие,                                              *
 *    r - прочитать из EEPROM,                                         *
 *    w - записать в EEPROM,                                           *
 *    g - получить текущие настройки,                                  *
 *    s - установить настройку. Для этого нужны следующие аргументы:   *
 * V (variable) - короткое название настройки, один символ             *
 *    (см. uploader.py)                                                *
 * D... (data) - последующие данные, 1Б, 4Б или строка c NULL вконце   *
 *                                                                     *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

#include <map>
#include <ESP8266WiFi.h>
#include <ESP8266WebServer.h>
#include <EEPROM.h>

// ======================== Настройки сети ========================

enum DeviceType {
    VoterDevice = 1,
    ServerDevice = 2,
    ServerAPDevice = 3 // Access Point
};

class Settings {
public: 
    byte type;
    
    char ssid[32];
    char password[64];
    
    byte ip[4];
    byte subnet[4];
    byte gateway[4];
    byte server_ip[4];

    void start();
    void settingsMode();
    void load();
    void save();

    void printArray(byte *arr);             // Для IP, шлюза и маски -- только 4 элемента
    
} settings;


// ======================== Веб-общением ========================

ESP8266WebServer server;

// Функция ответа клиенту
void respond() {
    server.send(200, "text/html", "<h1 align=\"center\">Success :D</h1>");
}

// ================ Процессом голосования =================

class Voting {
    int total[7];                // Общий результат (кол-во голосов за каждый вариант)
    std::map<int, int> detailed; // Подробный (IP - вариант ответа)
    // Не путать класс std::map со стандартной Ардуиновской функцией map() 

public:
    void update(int variant);
    void send();
    void clear();
} voting;

// ===================================================================

void setup() {  
    Serial.begin(115200, SERIAL_8N1);
    settings.start();

    // ----------- Режим настроек -----------
    
    delay(5000);
    if (Serial.available())
        if (Serial.read() == 's') 
            settings.settingsMode();
            
    settings.load();

    if (settings.type == ServerAPDevice) {
        WiFi.mode(WIFI_AP);                                     // AP - Access Point
        WiFi.softAPConfig(settings.ip, settings.gateway, settings.subnet); 
        WiFi.softAP(settings.ssid, settings.password, 1, 0, 8); // 8 - максимальное кол-во подключённых к локальной сети
    } else {
        WiFi.mode(WIFI_STA);                                    // STA - STAtion
        WiFi.config(settings.ip, settings.gateway, settings.subnet);
        
        Serial.print("Connecting");
        WiFi.begin(settings.ssid, settings.password);
        while (WiFi.status() != WL_CONNECTED) {
            Serial.print(".");
            delay(700);
        }
        Serial.println();
    }
  
    server.on("/", respond);

    #define F(n) [](){respond(); voting.update(n);}
    server.on("/0", F(0));
    server.on("/1", F(1));
    server.on("/2", F(2));
    server.on("/3", F(3));
    server.on("/4", F(4));
    server.on("/5", F(5));
    server.on("/6", F(6));
    #undef F

    server.begin();
  
}


void loop() {
  
    server.handleClient();

    if (Serial.available()) {
    
        switch (Serial.read()) {
    
        case 'c':           // 'c' - clear results
            voting.clear();
            break;
    
        case 'g':           // 'g' - get results
            voting.send();
            break;
      
        }
    
    }  
  
}



// ========================= Голосование =======================

void Voting::update(int variant) {
  
    char ip = server.client().remoteIP()[3]; // Последняя часть IP

    // Если устройство с данным IP ещё не голосовало
    if (detailed.find(ip) == detailed.end()) {
        total[variant]++;
        detailed.insert( std::make_pair(ip, variant) );
    }
  
}


/* Выдаёт в последовательный порт JSON-строку с результатами
 * Пример: 
 * {
 *    "total": [17, 18, 56, 2, 55, 100, 6],
 *    
 *    "detailed": [
 *      [2,   0],
 *      [3,   5],
 *      [76,  6],
 *      [101, 8],
 *      ...
 *    ]
 * }
 */
void Voting::send() {   
  
    Serial.print("{\"total\":[");
  
    for (int i=0; i<7; i++) {
        if (i) Serial.print(',');
        Serial.print(total[i]);
    }
  
    Serial.print("],\"votes\":[");
  
    for (auto i = detailed.begin(); i != detailed.end(); i++) {
        if(i != detailed.begin()) 
            Serial.print(',');
      
        Serial.print('[');
        Serial.print(i->first);
        Serial.print(',');
        Serial.print(i->second);
        Serial.print(']');
    }
  
    Serial.println("]}"); // -ln -- конец строки
  
}

void Voting::clear() {
   
    for(int i=0; i<7; i++) 
        total[i] = 0; 
    
    detailed.clear();
  
}

// ======================== Настройки ======================

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
                    Serial.println(type, DEC);
                    
                    Serial.print("SSID: ");
                    Serial.println(ssid);

                    Serial.print("Password: ");
                    Serial.println(password);

                    Serial.print("IP: ");
                    printArray(ip);

                    Serial.print("Server IP [not used]: ");
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
