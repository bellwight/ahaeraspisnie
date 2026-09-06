# Расписание группы 01.26.Р.ИИ.ГД.ОФ.9-НСК

Автоматически скачивает xlsx с расписанием НГОК из Google Таблицы,
вытаскивает пары для одной группы и публикует:

- **веб-страницу** `docs/index.html` — открывается в браузере;
- **календарь** `docs/schedule.ics` — подписка для Google/Apple/Outlook
  Calendar, обновляется сама по мере изменения расписания.

Всё бесплатно: обновление раз в 6 часов делает GitHub Actions,
раздачу файлов — GitHub Pages.

## Установка (один раз)

1. Создайте новый **публичный** репозиторий на GitHub (например,
   `ngok-schedule`) и загрузите в него все файлы из этой папки
   (сохранив структуру: `scripts/`, `docs/`, `.github/workflows/`).

   Проще всего через веб-интерфейс: New repository → Upload files →
   перетащить всё содержимое этой папки.

   Либо через git:
   ```bash
   cd ngok-schedule
   git init
   git add .
   git commit -m "init"
   git branch -M main
   git remote add origin https://github.com/<ваш_логин>/ngok-schedule.git
   git push -u origin main
   ```

2. В репозитории откройте **Settings → Secrets and variables → Actions
   → Variables → New repository variable** и добавьте:
   - Name: `SCHEDULE_XLSX_URL`
   - Value:
     ```
     https://docs.google.com/spreadsheets/d/1FiMov0r4UUDKT6A56NWMImpoUakDC2YDevgaOpJQ7Qc/export?format=xlsx
     ```

3. В **Settings → Pages** выберите источник **Deploy from a branch**,
   ветка `main`, папка `/docs`. Сохраните — GitHub даст вам ссылку вида
   `https://<ваш_логин>.github.io/ngok-schedule/`.

4. В **Settings → Actions → General → Workflow permissions** включите
   **Read and write permissions** (нужно, чтобы workflow мог
   закоммитить обновлённые файлы).

5. Запустите workflow вручную первый раз: вкладка **Actions** →
   "Update schedule" → **Run workflow**. Через минуту в `docs/`
   появятся `schedule.json` и `schedule.ics`, а страница оживёт.

## Что получите

- Страница: `https://<ваш_логин>.github.io/ngok-schedule/`
- Календарь для подписки: `https://<ваш_логин>.github.io/ngok-schedule/schedule.ics`
  - Google Calendar: «Другие календари» → «+» → «По URL» → вставить ссылку на .ics
  - Apple Calendar: Файл → Новая подписка на календарь → вставить ссылку
  - Внутри «Schedule»-подобных приложений обычно есть пункт «Добавить
    календарь по ссылке (iCal/webcal)» — работает так же.

Дальше ничего делать не нужно: раз в 6 часов Action сам скачивает
свежий xlsx, перестраивает JSON/ICS и коммитит изменения — Pages и
календарь подхватывают их автоматически.

## Если понадобится другая группа или расписание

- Название группы задаётся в `scripts/parse_schedule.py` константой
  `GROUP_NAME` — можно поменять на любую другую из того же файла.
- Частота обновления — в `.github/workflows/update.yml`, строка
  `cron: "0 */6 * * *"` (сейчас раз в 6 часов).

## Как это устроено (коротко)

Файл расписания — гигантская Excel-книга с объединёнными ячейками,
где каждую неделю просто дописывается новый блок под тем же или новым
листом. `parse_schedule.py`:

1. ищет во всех листах ячейку с точным названием группы;
2. читает даты и время с учётом объединения ячеек;
3. сопоставляет предметы с их порядковым номером пары (в файле номера
   местами "съезжают" от недели к неделе, поэтому сопоставление идёт
   по порядку появления, а не по абсолютному значению номера);
4. собирает все найденные недели, убирает дубли и отдаёт JSON.

Логика проверена вручную на неделе 7–12.09.2026 — все 15 пар совпали
день‑в‑день с оригиналом таблицы.
