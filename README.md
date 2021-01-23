# PPVoting
Обратная связь типа "вопрос"-"статистика ответов" во время презентаций в PowerPoint, а также небольшой таймер в углу экрана.


## Структура

Проект состоит из трёх частей:

1. **Голосовалки** (пульты голосования) с 7-ю кнопками. На них нажимают слушатели. Одна кнопка - один вариант ответа.

    ![Voter](https://github.com/MarkLagodych/assets/blob/main/PPVoting/voter.jpg?raw=true)
    
2. **Сервер** - аппарат, втыкающийся в компьютер докладчика. К нему по Wi-Fi подключаются голосовалки и передают нажатия кнопок.

    ![Server](https://github.com/MarkLagodych/assets/blob/main/PPVoting/server.jpg?raw=true)

3. **Плагин** для Microsoft PowerPoint - состоит из нескольких компонентов:

    1. **Загрузчик настроек** - находит файл настроек по содержанию первой надписи на первом слайде в презентации.
    
        Пример файла (*кодировка системная!*):
    
        ```json
            {
                "timer": {
                    "use": true,
                    "totalTime_min": 3, 
                    "blushTime_min": 2
                },  
                
                "voting": {
                    "use": true,
                    "port": 4
                },
                
                "logging": {
                    "use": true,
                    "file": "C:\\path\\to\\log\\file.txt"
                }
            }
        ```

    2. **Диаграмма** - отображает статистику ответов.

        ![Diagram](https://github.com/MarkLagodych/assets/blob/main/PPVoting/diagram.PNG?raw=true)
    
    3. **Таймер** - небольшой прямоугольник в углу слайда с текущим временем и временем до конца доклада. За заданное кол-во минут до конца он должен сменить цвет на красноватый.

        ![Normal](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer1.PNG?raw=true)
        
        ![Red](https://github.com/MarkLagodych/assets/blob/main/PPVoting/timer2.PNG?raw=true)
        
# Примечания

## Кодировка символов

Файл настроек должен быть в Юникоде в кодировке как в ОС по умолчанию потому, что PowerPoint плохо справляется с UTF-8 

