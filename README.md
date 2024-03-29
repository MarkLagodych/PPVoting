# PPVoting
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/MarkLagodych/PPVoting)](https://github.com/MarkLagodych/PPVoting/releases/latest) ![Platforms](https://img.shields.io/badge/platform-Windows-blue) ![Target](https://img.shields.io/badge/target-Microsoft_PowerPoint-orange)

Обратная связь типа "вопрос"-"статистика ответов" во время презентаций в PowerPoint, а также небольшой таймер в углу экрана.

## Структура

Проект состоит из трёх частей:

1. **Голосовалки** (пульты голосования) с 7-ю кнопками. На них нажимают слушатели. Одна кнопка - один вариант ответа.

    ![Voter](https://github.com/MarkLagodych/assets/blob/main/PPVoting/voter.jpg?raw=true)
    
2. **Сервер** - аппарат, втыкающийся в компьютер докладчика. К нему по Wi-Fi подключаются голосовалки и передают нажатия кнопок.

    ![Server](https://github.com/MarkLagodych/assets/blob/main/PPVoting/server.jpg?raw=true)
    
    Для корректной работы нужен драйвер CH341SER. Его можно взять из bin/driver или с оф. сайта разработчика: [CH341SER](http://www.wch.cn/downloads/CH341SER_EXE.html)

3. **Плагин** для Microsoft PowerPoint - состоит из нескольких модулей:

    1. **Вкладка меню** - нужна для настройки всех частей плагина. Все настройки сохраняются на компьютере в реестре. В любой момент их можно удалить кнопкой "Remove settings".
    
        ![PluginMenu](https://github.com/MarkLagodych/assets/blob/main/PPVoting/plugin.PNG?raw=true)
        
        Любой другой компонент плагина можно включить/выключить установив соответствующий флажок в группе "Usage".

    2. **Диаграмма** - отображает статистику ответов.

        ![Diagram](https://github.com/MarkLagodych/assets/blob/main/PPVoting/diagram.png?raw=true)
    
    3. **Таймер** - небольшой прямоугольник в углу слайда с текущим временем и временем до конца доклада. За заданное кол-во минут до конца он должен сменить цвет на красноватый.

        ![Normal](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer1.PNG?raw=true)
        ![Red](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer2.PNG?raw=true)
        
        После окончания показа слайдов все созданные прямоугольники удаляются.
        
    4. **Логгер** (модуль протоколирования) - записывает всю информацию, ошибки выполнения и результаты голосования в отдельный текстовый файл. Файл (log, журнал, протокол) нужно создать самостоятельно и указать путь к нему в меню PPVoting.
        
        ![Log](https://github.com/MarkLagodych/assets/blob/main/PPVoting/log.png?raw=true)
        
# Запуск

Вот общий порядок действий, чтобы заставить работать бинарники проекта у себя на компьютере. Подробнее в [Примечания: проверка и настройка](#примечания-проверка-и-настройка).

1. Установите плагин (надстройку PowerPoint) `bin/plugin/PPVoting.ppam`. Для этого в настройках PowerPoint'а в разделе "Настроить ленту" поставьте флажок напротив пункта "Разработчик", во вкладке меню, которая отобразится, нажмите на кнопку "Надстройки PowerPoint" и загрузите `PPVoting.ppam`. Отобразится ещё одна вкладка - "PPVoting"
    
    ![MenuTab](https://github.com/MarkLagodych/assets/blob/main/PPVoting/menutab.png?raw=true)
    
2. Настройте плагин как Вам угодно. Вот за что отвечает каждый виджет:
    | Раздел | Виджет(ы) | Значение |
    | --- | --- | --- |
    | Usage | Флажки | Включают/отключают компоненты плагина |
    | Timer | Time | Время (в минутах) доклада -- для обратного отсчёта таймера |
    | Timer | Blush | Время (в минутах) покраснения -- за это количество времени до конца доклада прямоугольник таймера изменит цвет  |
    | Voting | Поле COM | Номер COM-порта, к которому подключен сервер |
    | Voting | Check | Кнопка проверки (валидации) соединения по COM-порту. Если контакт плохой или к этому порту подключен не сервер, вылетит ошибка |
    | Voting | Поля "Diagram size" и поле gap | Размер диаграммы (соответственно ширина и высота) в пикселях и расстояние между столбиками (горизонтальными синими линиями) |
    | Logging | Кнопки выбора и просмотра файла отчёта | Выбор файла (сам плагин создать его не может) и просмотр в блокноте |
    | Special actions | Кнопка "Device manager" | Запускает devmgmt.msc |
    | Special actions | Кнопка "Remove settings" | Удаляет настройки PPVoting с компьютера |
    
3. Установите `bin/driver/CH341SER.EXE`.

4. Подключите сервер. Проверьте подключение с помощью кнопки "Check" во вкладке PPVoting. Если не отобразилось сообщение "Success", попробуйте сменить шнур/порт/модуль ESP, чтобы получить "Success".

5. Включите голосовалки (можно без них, используя телефон. Об этом в [WiFi-сеть сервера](#wifi-сеть-сервера))

6. Откройте презентацию PowerPoint (любую), запустите показ слайдов (слайд-шоу). Вот функционал плагина во время слайд-шоу:
    1. Таймер - в левом верхнем углу экрана.
    2. Диаграмма - нажмите B или . (кирилличная Ю) для показа/скрытия статистики ответов. Когда диаграмма видна, нажатие любой кнопки переключения слайдов вперёд обновляет статистику, а кнопки назад - удаляет статистику и закрывает диаграмму.
    3. Логгер - подробный отсчёт о работе плагина. Записывает в файл ошибки, информацию о соединении с сервером, общую статистику ответов и подробный отсчёт о каждом голосовавшем: IP - номер голоса.

# Компиляция прошивок, загрузка их в сервер и в голосовалки

О тонкостях в разделе [Примечания: компиляция и сборка](#примечания-компиляция-и-сборка).

1. Установите Python3, с помощью pip установите библиотеки PySerial и EspTool:
    ```sh
    py -m pip install pyserial
    py -m pip install esptool
    ```
    Возможны и другие варианты запуска pip: `pip ...`, `python -m pip ...`, `py3 -m pip ...`, `python3 -m pip ...`.
    
2. Установите Arduino IDE, с помощью менеджера плат установите пакет esp8266:
    1. Файл > Найтройки > в поле "Дополнительные ссылки для менеджера плат" вставьте http://arduino.esp8266.com/stable/package_esp8266com_index.json
    2. Инструменты > Плата > Менеджер плат... > Наберите "ESP", выберите "esp8266 by ESP8266 Community" > нажмите кнопку "Установка"
    
3. По очереди откройте скетчи .ino в `src/server` и `src/voter`. Для каждого выполните "Скетч > Экспорт бинарного файла". Полученные бинарники перетащите в `bin/server_voter` и переименуйте на "server.bin" и "voter.bin".

4. Отредактируйте загрузчики upload_server.py и upload_voter.py. Главное в них: COM-порт прошивки и настройки WiFi. Также можно отредактировать common.py (в нём все функции по загрузке прошивок и настроек), например подкорректировать таймауты. 

5. Запустите загрузчики upload_server.py и upload_voter.py. Чтобы удостовериться в том, что настройки были записаны, запустите скрипт check_settings.py

**Голосовалке для записи настроек нужно нажать на среднюю кнопку.** Если при включении (или после перезагрузки) нажата средняя кнопка, голосовалка переходит в режим настройки. О том, как подключить и прошить схему из `src/voter/scheme` в разделе [Печатная плата](#печатная-плата).
        
# Примечания: проверка и настройка

## Драйвер и соединение по USB-порту

1. Установите плагин из bin/plugin.
2. Установите драйвер из bin/driver.
2. Зайдите в PowerPoint, вкладку PPVoting, запустите диспетчер устройств.
3. Подключите сервер.
4. Диспетчер устройств во вкладке "Порты COM и LPT" должен показать новое устройство (опознать его).

![DevMgr](https://github.com/MarkLagodych/assets/blob/main/PPVoting/dev_mg.png?raw=true)

5. Если сервер всё ещё не опознан (появляется по вкладке "Другие устройства") или не появляется вообще нигде, смените USB-шнур сервера. Если всё ещё не опознан, смените модуль ESP.
6. Во вкладке "PPVoting" введите номер COM-порта, на котором висит сервер, и нажмите на кнопку "Check", после чего плагин попытается связаться с сервером и вы увидите сообщение о результате. 

## WiFi-сеть сервера

Всё общение голосовалок с сервером происходит по WiFi. Все настройки WiFi устанавливаются при прошитии, по умолчанию SSID: ArbUZ361, пароль: gumANOId1. Адрес веб-сервера, на который посылаются запросы в этой сети - `192.168.1.1`, полный запрос, где X - номер нажатой кнопки, - `http://192.168.1.1/X`. Стандартный ответ сервера - `<h1 align="center">Success :D</h1>`.

# Примечания: компиляция и сборка

## Плагин

### Ссылки

Этот плагин PowerPoint успешно собрался с такими ссылками (окно "Visual Basic" > меню "Tools" > опция "References..."):

- Visual Basic for Applications
- Microsoft PowerPoint 16.0 Object Library
- Microsoft Office 16.0 Object Library
- Microsoft Forms 2.0 Object Library
- Microsoft Scripting Runtime
    
### Библиотеки

Все исходники сторонних библиотек лежат в `lib/`

- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) (парсинг JSON)
- [CommIO](http://www.thescarms.com/VBasic/commio.aspx) (общение по COM-порту)
    
### Настройка сообщений об ошибках

1. Перейдите в Tools > VBAProject Properties.
2. В поле Conditional Compilation Arguments введите "SHOW_ERRORS = 1" без кавычек чтобы отображать или "SHOW_ERRORS = 0" чтобы скрывать сообщения об ошибках (постоянные всплывающие во время выполнения).

### Как добавлять custom UI к проекту (в данном случае, к .pptm файлу)

[Custom UI - Microsoft Docs](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file)

1. Открыть .pptm как zip-файл
2. Добить папку (напр. "PPVotingUI"), в неё .xml файл интерфейса (например, `/src/plugin/code/ui.xml`)
3. Отредактировать файл _rels/.rels -- перед тэгом `</Relationships>` вставить:
    ```xml
    <Relationship
        Id="PPVotingRelID" 
        Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" 
        Target="PPVotingUI/ui.xml"/>
    ```
        
Если требуется отредактировать существующий интерфейс, шаг 3 можно не выполнять
        
## Сервер и голосовалки

### Библиотека для хранения настроек прямо в внутри ESP8266

Исходники библиотеки `settings.h/.cpp` лежат в папке `src/settings`. Чтобы исходникам `voter.ino` и `server.ino` в соответственно `src/voter` и `src/server` подключить `settings`, им нужно указать `#include "../settings/settings.h"`. Но так как Arduino IDE не может собирать проекты, выходящие за пределы одной папки (скетчбука), `#include "../settings/settings.h"` вызовет ошибку. Поэтому исходники `settings` следует копировать в `src/voter` и `src/server` после каждого изменения.

### Печатная плата

Все схемы были сделаны с помощью сайта easyeda.com

Конечный экспортированный файл печатной платы - `src/voter/scheme/pcb/Gerber.zip`. Можно открыть (почти?) любым Gerber-viewer'ом, например https://www.altium.com/viewer/.

Плата построена таким образом, что для питания установлена просто перемычка, а для прошития - целая гребёнка, на которую вешается перемычка и переходник UART-USB.

![Voter connection](https://github.com/MarkLagodych/assets/blob/main/PPVoting/voter_expl.png?raw=true)

### GPL license notice

    PPVoting
    Copyright © 2021  Mark Lagodych

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
