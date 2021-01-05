/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *                        СЕРВЕР ДЛЯ ГОЛОСОВАЛОК                       *
 * ------------------------------------------------------------------- *
 *                                                                     *
 * - Эта программа должна запускаться на ESP8266.                      *
 *   Задачи программы:                                                 *
 *       1. получать данные от пультов слушателей по WiFi;             *
 *       2. передавать данные ПК докладчика.                           *
 *                                                                     *
 * - Передача данных ПК возможна при помощи драйвера CH341SER.         *
 *   Его можно скачать с оф. сайта разработчика:                       *
 *       http://www.wch.cn/downloads/CH341SER_EXE.html                 *
 *   или взять из папки PPVoting/bin/driver                            *
 *                                                                     *
 * - Сеть WiFi может быть создана самим ESP8266 или любым другим       * 
 *   устройством. В первом случае макс. кол-во пультов - 8.            *
 *   В обоих случаях нужно изменить код в соответствии с настройками   *
 *   сети и выбранным режимом.                                         *
 * ------------------------------------------------------------------- *
 *                                                                     *
 *                                                                     *
 *                            КАК РАБОТАЕТ КОД                         *
 * ------------------------------------------------------------------- *
 *                           * Общение по WiFi *                       *
 * 1. Поключиться к / Создать сеть WiFi.                               *
 * 2. Создать веб-сервер.                                              *
 * 3. При каждом запросе веб-страницы вида /Х, где Х - число от 0 до 6 * 
 *    записать голос за Х-й вариант.                                   *
 *                                                                     *
 *                           * Общение по USB *                        *
 * 1. Начать общение по USB.                                           *
 * 2. При каждом запросе "c" очистить текущие результаты голосования.  *
 * 3. При каждом запросе "g" выдать JSON-строку с общим результатом    *
 *    голосования и подробным.                                         *
 *    Общий результат - это список вида                                *
 *        [N0, N1, ..., N6], где Ni - кол-во голосов за i-й вариант.   *
 *    Подробный результат - это словарь вида                           *
 *        "IP-адрес" - "Выбраный варианта".                            *
 *                                                                     *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

#include <map>
#include <ESP8266WiFi.h>
#include <ESP8266WebServer.h>


// ======================== Настройки сети ========================

#define WIFI_MODE_OWN      1             // Собственная WiFi-сеть
#define WIFI_MODE_EXTERNAL 2             // Внешняя WiFi-сеть
#define WIFI_MODE          WIFI_MODE_OWN // Режим WiFi

const String
  ssid     = "ArbUZ361",    // Название сети
  password = "gumANOId1";   // Пароль
  
const IPAddress
  ip(192, 168, 1, 1),       // IP-адрес этого устройства в сети
  gateway(192, 168, 1, 1),  // Шлюз
  subnet(255, 255, 255, 0); // Маска подсети

const int
  webServerPort = 80;       // Порт веб-сервера

ESP8266WebServer server(webServerPort);

// Функция ответа клиенту (изменения целесообразны для отладки)
void respond() {
  server.send(200, "text/html", "<h1 align=\"center\">Success</h1>");
}


// ================ Связанное с процессом голосования =================

int result_total[7];                // Общий результат
std::map<int, int> result_detailed; // Подробный результат
// Не путать класс std::map со стандартной Ардуиновской функцией map() 

void updateResult(int variant);
void sendResult();
void clearResult();


// ========================== Основной код ==========================

void setup() {
  
  /* Модуль ESP8266 для перепрошивки требует подачи тока на специальный
   * контакт при включении. Если модуль идёт вместе с USB-адаптером,
   * данное действие пропускается. Однако после запуска/подключения
   * WiFi-сети ESP8266 не может прошиться без замыкания. Задержка 
   * при запуске режает проблему.
   */
  delay(7000);
  
  Serial.begin(115200, SERIAL_8N1); /* 8-N-1 - стандартаная конфигурация
                                     * последовательного порта
                                     */

#if WIFI_MODE == WIFI_MODE_OWN
    WiFi.mode(WIFI_AP);                   // AP - Access Point
    WiFi.softAPConfig(ip, gateway, subnet); 
    WiFi.softAP(ssid, password, 1, 0, 8); /* 1, 0 - значения по умолчанию
                                           * 8 - максимальное кол-во подключённых
                                           * к локальной сети
                                           */
#else
    WiFi.mode(WIFI_STA);                  // STA - STAtion
    WiFi.config(ip, gateway, subnet);
    WiFi.begin(ssid, password);
#endif
  
  server.on("/", respond);

  #define F(n) [](){respond(); updateResult(n);}
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
      clearResult();
      break;
    
    case 'g':           // 'g' - get results
      sendResult();
      break;
      
    }
    
  }
  
}



// =================== Реализация обмена ответами ==================

void updateResult(int variant) {
  
  char ip = server.client().remoteIP()[3]; // Последняя часть IP

  // Если устройство с данным IP ещё не голосовало
  if (result_detailed.find(ip) == result_detailed.end()) {
    result_total[variant]++;
    result_detailed.insert( std::make_pair(ip, variant) );
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
void sendResult() {   
  
  Serial.print("{\"total\":[");
  
  for (int i=0; i<7; i++) {
    if (i) Serial.print(',');
    Serial.print(result_total[i]);
  }
  
  Serial.print("],\"votes\":[");
  
  for (auto i = result_detailed.begin(); i != result_detailed.end(); i++) {
    if(i != result_detailed.begin()) 
      Serial.print(',');
      
    Serial.print('[');
    Serial.print(i->first);
    Serial.print(',');
    Serial.print(i->second);
    Serial.print(']');
  }
  
  Serial.println("]}"); // -ln -- конец строки
  
}

void clearResult() {
   
  for(int i=0; i<7; i++) 
    result_total[i] = 0; 
    
  result_detailed.clear();
  
}
