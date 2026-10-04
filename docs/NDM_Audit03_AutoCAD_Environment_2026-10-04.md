# Audit03: Среда AutoCAD

## Фактическая Проба

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

## Граница Приемки

Реальный DWG export/import: **NotRun**. Регистрация COM-класса и наличие
executable не подтверждают возможность использования ActiveDocument/ModelSpace.
Проверки существующих Region fixtures, snapshot, единиц и preflight могут
выполняться независимо, но не заменяют настоящую проверку DWG.
Другие проверки Audit03 продолжаются; итоговый выпуск не объявляется завершенным.
