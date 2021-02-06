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
 * - Сеть WiFi может быть создана самим ESP8266 (макс. 8 пультов) или  *
 * любым другим устройством.                                           *
 * ------------------------------------------------------------------- *
 *                                                                     *
 *                            КАК РАБОТАЕТ КОД                         *
 * ------------------------------------------------------------------- *
 *                           * Общение по WiFi *                       *
 * 0. Ожидание настроек: если сразу после включения пришло "s",        *
 *    войти в режим настройки, если нет, продолжить работу             *
 * 1. Подключиться к / Создать сеть WiFi.                              *
 * 2. Создать веб-сервер.                                              *
 * 3. При каждом запросе веб-страницы вида /Х, где Х - число от 0 до 6 * 
 *    записать голос за Х-й вариант.                                   *
 * 4. Если по последовательному порту пришло "c", очистить все ответы, * 
 *    если "g", то выслать JSON-строку с результатом, если "k", то     *
 *    ответить "OK" -- это специальный валидационный запрос            *
 *                                                                     *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

#include <map>
#include <ESP8266WiFi.h>
#include <ESP8266WebServer.h>
#include "settings.h"

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
    
    // При первом включении в памяти не будет правильных настроек
    // Тогда продолжать нет смысла, всё равно не выйдет
    // Значит, всё, что остаётся делать - ждать настроек
    while (settings.type != ServerDevice && settings.type != ServerAPDevice)
        settings.settingsMode();

    if (settings.type == ServerAPDevice) {
      
        WiFi.mode(WIFI_AP);                                     // AP - Access Point
        WiFi.softAPConfig(settings.ip, settings.gateway, settings.subnet); 
        WiFi.softAP(settings.ssid, settings.password, 1, 0, 8); // 8 - максимальное кол-во подключённых к локальной сети
        
    } else if (settings.type == ServerDevice) {
      
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

        case 'k':           // 'k' - validation request
            Serial.print("OK");
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
