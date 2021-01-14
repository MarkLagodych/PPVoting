/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *                       КОД ДЛЯ ГОЛОСОВАЛОК (ПУЛЬТОВ)                 *
 * ------------------------------------------------------------------- *
 *                                                                     *
 * - Эта программа должна запускаться на ESP8266.                      *
 *   Задачи программы:                                                 *
 *       1. проверять, нажата ли одна из 7 кнопок;                     *
 *       2. посылать результат на веб-сервер через WiFi.               *
 * ------------------------------------------------------------------- *
 *                                                                     *
 *                                                                     *
 *                            КАК РАБОТАЕТ КОД                         *
 * ------------------------------------------------------------------- *
 *                           * Общение по WiFi *                       *
 * 0. Если при включении нажата средняя кнопка, войти в режим настроек *
 * 1. Поключиться к сети WiFi, используя статический IP (быстрее, чем  *
 *    автоматический).                                                 *
 * 2. Проверить кнопки (7).                                            *
 * 3. Если какая-либо нажата, то запросить веб-страницу вида           *
 *    http://A.B.C.D/Х                                                 *
 *    где A.B.C.D - IP сервера, Х - номер варианта (от 0 до 6)         *
 *                                                                     *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

#include <cstring>
#include <ESP8266WiFi.h>
#include <ESP8266HTTPClient.h>
#include "settings.h"


HTTPClient http;

const char *reqf = "http://%i.%i.%i.%i/%s";
char request[64];

void vote(int variant);      // Посылает веб-серверу выбранный вариант ответа ("голосует")

const int ports[] = {        // Порты кнопок
  16, 14, 12, 13, 5, 4, 15
};


void blink(int delay_time) {
    digitalWrite(LED_BUILTIN, HIGH);
    delay(delay_time);
    digitalWrite(LED_BUILTIN, LOW);
}


void setup() {

    Serial.begin(115200, SERIAL_8N1);
  
    // Настройка портов для кнопок
    for(int port : ports) {
        pinMode(port, INPUT); 
        digitalWrite(port, LOW);
    }

    pinMode(LED_BUILTIN, OUTPUT);
    digitalWrite(LED_BUILTIN, LOW);

    // "Настройка настраивания"
    settings.start();

    if (digitalRead(ports[3]))
        settings.settingsMode();

    settings.load();

    // При первом включении в памяти не будет правильных настроек
    // Тогда продолжать нет смысла, всё равно не выйдет
    // Значит, всё, что остаётся делать - ждать настроек
    while (settings.type != VoterDevice)
        settings.settingsMode();

    Serial.end();

    blink(500);
  
    // Настройка и подключение к WiFi-сети.
    // Если успешно подключено, моргнуть светодиодом.
    WiFi.begin(settings.ssid, settings.password);
    WiFi.config(settings.ip, settings.gateway, settings.subnet);
    while (WiFi.status() != WL_CONNECTED) 
        delay(1000);    
  
    blink(1000);
}


void loop() {
    
    for(int i=0; i<7; i++) {
        
        if(digitalRead(ports[i])) {
            vote(i); 
            break;
        }
      
    }
    
    delay(100);
  
}


void vote(int variant) {
    
    if(WiFi.status() == WL_CONNECTED) {

        sprintf(request, reqf,
        (int) settings.ip[0],
        (int) settings.ip[1],
        (int) settings.ip[2],
        (int) settings.ip[3],
        (int) variant);
        
        if(http.begin(request)) {
          http.GET();
          http.end();
    
          blink(700);
        }
  
    }
    
}
