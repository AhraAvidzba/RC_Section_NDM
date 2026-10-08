# RC Section NDM

Excel VBA-программа для расчета железобетонных сечений по НДМ: `N + Mx + My`, несущая способность, трещины по СП 35 и СП 63, устойчивость. Поддерживаются круг, составные прямоугольные и скругленные сечения, импорт и экспорт AutoCAD.

Готовая книга: [RC_Section_NDM.xlsm](workbook/output/RC_Section_NDM.xlsm). Инструкция находится на листе `Справка`.

## Сборка

Нужны Windows и установленный Microsoft Excel. Для сборки должен быть разрешен доступ к объектной модели проекта VBA; для работы книги разрешены макросы. AutoCAD нужен только для импорта и экспорта.

Из корня проекта:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
```

Сборка создает книгу с настройками шаблона. Не запускайте ее поверх открытой рабочей книги.

## Разработка

- `src/` - VBA; `tests/` - тесты; `tools/` - сборка и проверки.
- [AGENTS.md](AGENTS.md) - правила изменений.
- [Расчетная модель](docs/MathematicalModel.md), [единицы и знаки](docs/CoordinateSystem.md), [проверки](docs/ValidationPlan.md).
- [Нормативная трассировка](docs/NormativeTraceability.md), [открытые вопросы](docs/OpenNormativeQuestions.md), [отчеты](docs/reports/README.md).
- `norms/`, `reference/`, `control_examples/` - источники и контрольные материалы.

Поддерживается ненапрягаемая арматура. Автоматические тесты не заменяют независимую нормативную проверку расчета.
