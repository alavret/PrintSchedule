# PrintSchedule

Скрипт для печати расписания из календаря пользователя в формате Word.

## Описание

Этот скрипт подключается к CalDAV календарю, получает все встречи на текущий день и создает документ Word с таблицей, содержащей информацию о встречах. При наличии CardDAV адресной книги, email адреса участников автоматически преобразуются в полные имена (Фамилия Имя Отчество).

## Установка

1. Установите зависимости:

```bash
pip install -r requirements.txt
```

2. Создайте файл `.env` на основе `.env.example`:

```bash
cp .env.example .env
```

3. (Опционально) Создайте файл `meeting_room_emails.txt` со списком email переговорных комнат для их исключения из списка участников:

```bash
cp meeting_room_emails.txt.example meeting_room_emails.txt
```

Отредактируйте файл и добавьте email адреса переговорных комнат (один на строку):
```
conference-room-1@company.com
meeting-room-a@company.com
переговорная-401@company.ru
```

4. Отредактируйте файл `.env` и укажите ваши данные для подключения:
   - `CALDAV_URL` - URL вашего CalDAV сервера
   - `CALDAV_USERNAME` - имя пользователя для CalDAV
   - `CALDAV_PASSWORD` - пароль для CalDAV
   - `CARDDAV_URL` - URL вашего CardDAV сервера (опционально, для разрешения имен участников)
   - `CARDDAV_USERNAME` - имя пользователя для CardDAV (опционально, по умолчанию используется CALDAV_USERNAME)
   - `CARDDAV_PASSWORD` - пароль для CardDAV (опционально, по умолчанию используется CALDAV_PASSWORD)
   - `TIMEZONE_OFFSET` - смещение часового пояса в часах от UTC (опционально, по умолчанию: +3 для Москвы)
     * Примеры: `+3` (Москва), `0` (Лондон/UTC), `-5` (Нью-Йорк EST), `+5.5` (Индия), `+8` (Пекин)
   - `DOCUMENT_TITLE` - название документа (опционально, по умолчанию: Расписание)
   - `OUTPUT_PATH` - путь для сохранения документов (опционально, по умолчанию: текущая директория)
   - `FILENAME_PREFIX` - префикс имени файла перед датой (опционально, по умолчанию: schedule_)
   
   Также можно создать файл `meeting_room_emails.txt` со списком email переговорных комнат для их исключения из списка участников.

## Примеры настройки для популярных сервисов

### Yandex 360
```
CALDAV_URL=https://caldav.yandex.ru
CALDAV_USERNAME=username@domain.ru
CALDAV_PASSWORD=ваш_пароль
CARDDAV_URL=https://carddav.yandex.ru/addressbook/username@domain.ru/
TIMEZONE_OFFSET=+3
# Для Yandex 360 обычно используются те же учетные данные
# CARDDAV_USERNAME и CARDDAV_PASSWORD можно не указывать
```

**Важно:** Для Yandex 360 необходимо создать специальный пароль приложения:
1. Перейдите на страницу https://id.yandex.ru
2. Откройте раздел **Безопасность**
3. Выберите пункт **Пароли приложений**
4. Создайте новый пароль для приложения CalDAV/CardDAV
5. Используйте этот пароль в параметре `CALDAV_PASSWORD`

### Google Calendar
```
CALDAV_URL=https://apidata.googleusercontent.com/caldav/v2/[ваш_email]/events
CALDAV_USERNAME=ваш_email@gmail.com
CALDAV_PASSWORD=пароль_приложения
CARDDAV_URL=https://www.googleapis.com/.well-known/carddav
CARDDAV_USERNAME=ваш_email@gmail.com
CARDDAV_PASSWORD=пароль_приложения
```

Примечание: для Google Calendar необходимо создать пароль приложения в настройках безопасности Google аккаунта.

### Nextcloud
```
CALDAV_URL=https://ваш-nextcloud.com/remote.php/dav/calendars/username/calendar-name
CALDAV_USERNAME=ваше_имя_пользователя
CALDAV_PASSWORD=ваш_пароль
CARDDAV_URL=https://ваш-nextcloud.com/remote.php/dav/addressbooks/username/
# Для Nextcloud обычно используются те же учетные данные
# CARDDAV_USERNAME и CARDDAV_PASSWORD можно не указывать
```

### iCloud
```
CALDAV_URL=https://caldav.icloud.com
CALDAV_USERNAME=ваш_apple_id
CALDAV_PASSWORD=пароль_приложения
CARDDAV_URL=https://contacts.icloud.com/[ваш-id]/carddavhome/card/
# Для iCloud обычно используются те же учетные данные
# CARDDAV_USERNAME и CARDDAV_PASSWORD можно не указывать
```

Примечание: для iCloud необходимо создать пароль для приложения в настройках Apple ID.

## Использование

### Базовое использование

Запустите скрипт без параметров для получения расписания на сегодня:

```bash
python print_schedule.py
```

### Использование с параметром даты

Вы можете указать дату с помощью параметра `-d` или `--date`:

**Справка по всем параметрам:**
```bash
python print_schedule.py --help
```

**Сегодня (по умолчанию):**
```bash
python print_schedule.py -d 0
```

**Вчера:**
```bash
python print_schedule.py -d -1
```

**Завтра:**
```bash
python print_schedule.py -d +1
```

**Конкретная дата (DD.MM.YYYY):**
```bash
python print_schedule.py -d 15.11.2025
```

**Конкретная дата (DD.MM.YY):**
```bash
python print_schedule.py -d 15.11.25
```

**День и месяц текущего года (DD.MM):**
```bash
python print_schedule.py -d 15.11
```

**Автоматическая печать после создания документа (только Windows):**
```bash
python print_schedule.py -p
python print_schedule.py --print
python print_schedule.py -d +1 -p  # Завтра и сразу на печать
```

Скрипт создаст файл с именем в формате `schedule_DD.MM.YY.docx` в указанной директории (или текущей).

## Формат документа

Документ содержит таблицу со следующими колонками:
- **Время** - время начала и окончания встречи с указанием продолжительности
- **Тема** - название встречи
- **Место** - место проведения встречи
- **Участники** - список участников встречи с дополнительной информацией:
  * Организатор события автоматически исключается из списка участников
  * Email переговорных комнат (из файла `meeting_room_emails.txt`) также исключаются
  * Участники разделяются на обязательных и необязательных (опциональных)
  * Если настроен CardDAV, отображаются полные имена вместо email адресов
  * Перед каждым участником отображается индикатор статуса участия:
    - ✓ - приглашение принято (ACCEPTED)
    - ✗ - приглашение отклонено (DECLINED)
    - ? - участие под вопросом (TENTATIVE)
    - → - делегировано другому участнику (DELEGATED)
    - ○ - ожидается ответ (NEEDS-ACTION или нет ответа)

## Требования

- Python 3.7+
- Доступ к CalDAV серверу
- Интернет-соединение

## Лицензия

MIT

