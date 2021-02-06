# PPVoting
Обратная связь типа "вопрос"-"статистика ответов" во время презентаций в PowerPoint, а также небольшой таймер в углу экрана.


## Структура

Проект состоит из трёх частей:

1. **Голосовалки** (пульты голосования) с 7-ю кнопками. На них нажимают слушатели. Одна кнопка - один вариант ответа.

    ![Voter](https://github.com/MarkLagodych/assets/blob/main/PPVoting/voter.jpg?raw=true)
    
2. **Сервер** - аппарат, втыкающийся в компьютер докладчика. К нему по Wi-Fi подключаются голосовалки и передают нажатия кнопок.

    ![Server](https://github.com/MarkLagodych/assets/blob/main/PPVoting/server.jpg?raw=true)
    
    Для корректной работы нужен драйвер CH341SER. Его можно взять из bin/driver или с оф. сайта разработчика: http://www.wch.cn/downloads/CH341SER_EXE.html

3. **Плагин** для Microsoft PowerPoint - состоит из нескольких модулей:

    1. **Вкладка меню** - нужна для настройки всех частей плагина. Все настройки сохраняются на компьютере (если конкретнее, в реестре). В любой момент их можно удалить кнопкой "Remove settings".
    
        ![MenuTab](https://github.com/MarkLagodych/assets/blob/main/PPVoting/plugin.PNG?raw=true)
        
        Любой другой компонент плагина можно включить/выключить установив соответствующий флажок в группе "Usage".

    2. **Диаграмма** - отображает статистику ответов.

        ![Diagram](https://github.com/MarkLagodych/assets/blob/main/PPVoting/diagram.png?raw=true)
    
    3. **Таймер** - небольшой прямоугольник в углу слайда с текущим временем и временем до конца доклада. За заданное кол-во минут до конца он должен сменить цвет на красноватый.

        ![Normal](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer1.PNG?raw=true)
        
        ![Red](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer2.PNG?raw=true)
        
    4. **Логгер** (модуль протоколирования) - записывает всю информацию, ошибки выполнения и результаты голосования в отдельный текстовый файл. Файл (log, журнал, протокол) нужно создать самостоятельно и указать путь к нему в меню PPVoting.
        
        ![Log](https://github.com/MarkLagodych/assets/blob/main/PPVoting/log.png?raw=true)
        
# Примечания: проверка и настройка

## Драйвер и соединение по USB-порту

1. Установите плагин из bin/plugin.
2. Установите драйвер из bin/driver.
2. Зайдите в PowerPoint, вкладку PPVoting, запустите диспетчер устройств.
3. Подключите сервер.
4. Диспетчер устройств во вкладке "Порты COM и LPT" должен показать новое устройство (опознать его).

![DevMgr](https://github.com/MarkLagodych/assets/blob/main/PPVoting/dev_mg.png?raw=true)

5. Если сервер всё ещё не опознан (появляется по вкладке "Другие устройства") или не появляется вообще нигде, смените USB-шнур сервера. Это не шутка. Часто работает.
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

Все исходники сторонних библиотек лежат в `src/`

- VBA-JSON: https://github.com/VBA-tools/VBA-JSON
- CommIO (общение по COM-порту): http://www.thescarms.com/VBasic/commio.aspx
    
### Настройка сообщений об ошибках

1. Перейдите в Tools > VBAProject Properties.
2. В поле Conditional Compilation Arguments введите "SHOW_ERRORS = 1" без кавычек чтобы отображать или "SHOW_ERRORS = 0" чтобы скрывать сообщения об ошибках (постоянные всплывающие во время выполнения).

### Как добавлять custom UI к проекту (в данном случае, к .pptm файлу)

https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file

1. Открыть .pptm как zip-файл
2. Добить папку (напр. "PPVotingUI"), в неё .xml файл интерфейса
3. Отредактировать файл _rels/.rels -- перед тэгом </Relationships> вставить:
    <Relationship
        Id="PPVotingRelID" 
        Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" 
        Target="PPVotingUI/ui.xml"/>
        
## Сервер и голосовалки

### Библиотека для хранения настроек прямо в внутри ESP8266

Исходники библиотеки `settings.h/.cpp` лежат в папке `src/settings`. Чтобы исходникам `voter.ino` и `server.ino` в соответственно `src/voter` и `src/server` подключить `settings`, им нужно указать `#include "../settings/settings.h"`. Но так как Arduino IDE не может собирать проекты, выходящие за пределы одной папки (скетчбука), `#include "../settings/settings.h"` вызовет ошибку. Поэтому исходники `settings` следует копировать в `src/voter` и `src/server` после каждого изменения.

### Печатная плата

Все схемы были сделаны с помощью сайта easyeda.com

Конечный экспортированный файл печатной платы - `src/voter/scheme/pcb/Gerber.zip`. Можно открыть (почти?) любым Gerber-viewer'ом, например https://www.altium.com/viewer/.

Плата построена таким образом, что для питания установлена просто перемычка, а для прошития - целая гребёнка, на которую вешается перемычка и переходник UART-USB.

![Voter connection](https://github.com/MarkLagodych/assets/blob/main/PPVoting/voter_expl.png?raw=true)