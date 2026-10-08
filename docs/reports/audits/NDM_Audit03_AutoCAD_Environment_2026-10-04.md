# Audit03: Среда AutoCAD

## Актуальная Проверка Autodesk AutoCAD

После уточнения пользователя запущен именно Autodesk AutoCAD 2023:
`C:\Program Files\Autodesk\AutoCAD 2023\acad.exe`, версия COM
`24.2s (LMS Tech)`, HWND `25888782`, процесс `25896`.
SOFiPLUS для последующих проверок не запускается. Реестровые регистрации
COM не менялись; общий AutoCAD.Application недоступен, но ROT-запрос
AutoCAD.Application.24.2 возвращает уже запущенный Autodesk acad.exe.

Реальный export/import/save/reopen завершен: **409 passed, 0 failed**.
Журнал: `regression/Audit03/autocad_real_roundtrip_v271_2026-10-04.log`.
Созданы только собственные DWG в каталоге regression/Audit03; тест
проверяет HWND и отсутствие чужих открытых документов перед созданием.
Пользовательская Excel-сессия 23476 не изменялась и не закрывалась.

Подтверждены свойства настоящих Region/Layer/Text, миллиметровые координаты
и площади после изменения текущих единиц и знаков Config, отключение
оформления, сохранность Region при очистке и повторное открытие DWG.
Подробные границы и хеши: `NDM_Audit03_AutoCAD_Config_Review.md`.
Это не окончательная приемка всего Audit03.

Захват пикселей окна AutoCAD возвращает FrameArrived/window-capture timeout.
Текстовое дерево UI и COM-свойства доступны, но визуальная pixel-проверка
чертежа остается **NotRun**, а не PASS.

## Историческая Проба До Уточнения Пользователя

Для реального export/import необходим доступный COM-сервер CAD. Перед пробой
процессов acad/sofp/sofc не было. Проверены зарегистрированные ProgID:

- AutoCAD.Application.23/23.1: SOFiSTiK 2020 SOFiCAD,
  `C:\Program Files\SOFiSTiK\2020\soficad_oem_enu_2020\sofc.exe /Automation`.
- AutoCAD.Application.24/24.1/24.2/24.3: SOFiSTiK 2024 SOFiPLUS,
  `C:\Program Files\SOFiSTiK\2024\sofiplus_x_enu_2024\sofp.exe /Automation`.

Оба executable присутствуют. Отдельный контролируемый вызов
`New-Object -ComObject AutoCAD.Application.24.3` завершился ошибкой:

```text
CLSID: {AF18D91C-A699-4578-ADC6-972F3BA007F0}
HRESULT: 0x80080005
CO_E_SERVER_EXEC_FAILURE
```

Объект Application не получен; чертежи не создавались и не изменялись.
После terminal failure процессов acad/sofp/sofc также нет. Это не ожидание
работающего CAD-job и не доказательство дефекта инженерного ядра.

## Граница Исторической Пробы

На момент этой ранней пробы реальный DWG export/import был **NotRun**.
Это историческая диагностика OEM-регистрации, не актуальный результат.
Регистрация COM-класса и наличие
executable не подтверждают возможность использования ActiveDocument/ModelSpace.
Проверки существующих Region fixtures, snapshot, единиц и preflight могут
выполняться независимо, но не заменяют настоящую проверку DWG.
Актуальная проверка на Autodesk AutoCAD приведена выше. Другие проверки
Audit03 продолжаются; итоговый выпуск не объявляется завершенным.
