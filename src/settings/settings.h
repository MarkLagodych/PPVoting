/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
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
 * D... (data) - последующие данные, 1Б, 4Б или строка c NULL в конце  *
 *                                                                     *
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
 
#ifndef SETTINGS_H_INCLUDED
#define SETTINGS_H_INCLUDED

// В скетче подключается автоматически, а здесь (в библиотеке) нужно отдельно это указать
#include <Arduino.h>

#include <EEPROM.h>

enum DeviceType {
    VoterDevice = 1,
    ServerDevice = 2,
    ServerAPDevice = 3 // Access Point
};

class Settings {
public: 
    byte type; // Значения брать из enum DeviceType
    
    char ssid[32];
    char password[64];
    
    byte ip[4];
    byte subnet[4];
    byte gateway[4];
    byte server_ip[4];

    // Инициализация модуля
    void start();
    
    // Войти в режим настроек
    void settingsMode();

    void load();
    void save();
    
private:
    void printArray(byte *arr); // Для IP, шлюза и маски -- только 4 элемента
    
};

extern Settings settings;

#endif // SETTINGS_H_INCLUDED