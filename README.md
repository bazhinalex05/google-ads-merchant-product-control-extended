# Unified Merchant Funnel Product Control Extended V2

Google Ads Script для великих Merchant Center асортиментів. Скрипт розбиває важку обробку на етапи через `RunState`, збирає Merchant snapshot, статистику Google Ads, правила `ProductTypes`, карантин і публікує фінальний лист `Products`.

## Код

Актуальний файл скрипта: [`script.js`](script.js)

## Як встановити

1. Відкрийте `script.js` у GitHub.
2. Натисніть кнопку копіювання коду.
3. У Google Ads відкрийте `Tools -> Bulk actions -> Scripts`.
4. Створіть новий скрипт.
5. Вставте код.
6. У верхній частині скрипта замініть `SPREADSHEET_URL` на URL Google Sheets клієнта.
7. Увімкніть потрібні Advanced APIs для Merchant API у Google Ads Scripts.
8. Запустіть скрипт, щоб він створив службові листи.
9. У листі `Settings` заповніть `merchant_id`.
10. Перевірте перемикачі модулів і ліміти запуску.
11. Поставте погодинний розклад запуску.
12. Після завершення пайплайна перевірте лист `Products`.

## Основні листи

- `Settings` - налаштування клієнта й модулів.
- `Products` - фінальний додатковий фід для Merchant Center.
- `ProductTypes` - дерево категорій і правила включення.
- `ProductDiagnostics` - діагностика товарів.
- `Dashboard` і `DashboardData` - зведення.
- `RunState` і `RunLog` - технічний стан запусків.

## Результат

У `Products` мають бути `id`, дві колонки `excluded_destination`, custom label для funnel stage і, якщо увімкнено, ще один custom label для benchmark / priority.
