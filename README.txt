Как пользоваться скриптом: 
1. Копируем всю папку "docx-modifier" куда-нибудь к себе
2. Для работы необходим установленный python 3.9 (на других 3.x тоже должно работать, но не проверялось), также необходим pip (как правило устанавливается вместе с python)
3. Необходимо наличие пакетов lxml==4.9.3, python-docx==0.8.11, win32-setfiletime==1.0.0.
    Установка из командной строки:
    - открываем командную строку
    - переходим в папку "docx-modifier"
    - выполняем команду: pip install -r requirements.txt
    - готово
4. Открываем любым текстовым редактором файл properties.ini, заполняем данные, которые хотим изменить в документах, следим за форматом времени.
    ВАЖНО! По-умолчанию скрипт меняет ВСЕ параметры, перечисленные в настройка. Если НЕ ХОТИМ изменять какой-то из параметров - удаляем строку
5. Кладем все нужные файлы docx в подпапку "../docx-modifier/res" (структура подпапок внутри не важна, скрипт найдет все файлы docx в любых подпапках папки res.
6. В командной строке, находясь в папке "docx-modifier" запускаем скрипт: python docx-modifier.py
7. Следим за выполнением программы, до появления сообщения FINISHED. В случае вылетов и ошибок - вы что-то делаете не так, а мне было лень писать исключения.
