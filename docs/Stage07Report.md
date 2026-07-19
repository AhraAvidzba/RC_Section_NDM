# Отчет этапа 7: ширина раскрытия уже образовавшихся нормальных трещин

## Статус

Этап 7 реализован как расчет ширины раскрытия уже образовавшихся нормальных трещин для состояния `N + Mx + My`.

Не реализовывались:

- момент образования трещин;
- условие появления трещины;
- `Mcrc`;
- пакетный расчет сочетаний;
- новые геометрии.

## Принятая расчетная постановка

Расчет трещин выполняется для уже трещиноватого сечения.

Принято:

- растянутый бетон в режиме по умолчанию не создает внутренних усилий;
- сжатый бетон учитывается;
- вся продольная обычная ненапрягаемая арматура участвует по фактическим деформациям;
- эксплуатационное сочетание второй группы задается отдельно и обычно меньше предельного сочетания первой группы.

`CCrackWidthCalculator` не решает равновесие. Он получает уже найденное НДС из `CSectionSolver`, рассчитанное на фактическое SLS-сочетание `N + Mx + My`.

## Concrete.TensionMode

На лист `System` добавлена видимая и редактируемая настройка:

```text
Concrete.TensionMode = Ignore / UseDiagram
```

Режим `Ignore`:

- напряжение растянутого бетона равно нулю;
- касательный модуль растянутого бетона равен нулю.

Режим `UseDiagram`:

- растянутая ветвь определяется параметрами диаграммы материала;
- на текущем этапе используется временная параметризованная растянутая ветвь по `Concrete.Eb` и `Concrete.Rbt.SLS`.

Значение по умолчанию для расчета трещин:

```text
Ignore
```

## Параметры System

Добавлены или актуализированы ключи:

- `Concrete.TensionMode`;
- `CrackWidth.Enabled`;
- `CrackWidth.Allowable`;
- `CrackWidth.CrackSpacing`;
- `CrackWidth.StrainFactor`;
- `CrackWidth.DurationFactor`.

Параметры трещин имеют статус `PROVISIONAL_FOR_SOLVER_TESTING` до нормативной трассировки по СП 35 и СП 63.

## Код

Создано:

- `src/Crack/CCrackWidthCalculator.cls`
- `tests/modTestCrackWidth.bas`
- `tools/build_workbook/Run-CrackTests.ps1`
- `docs/Stage07Report.md`

Изменено:

- `src/Materials/CConcreteBilinearMaterial.cls`
- `src/Materials/CConcreteTrilinearMaterial.cls`
- `src/Excel/CCapacityResultWriter.cls`
- `tools/build_workbook/Build-Workbook.ps1`
- `tools/build_workbook/Refresh-Workbook.ps1`

## Результаты расчета

`CCrackWidthCalculator` определяет:

- количество растянутых стержней;
- максимальную деформацию растянутой арматуры;
- напряжение в соответствующей арматуре;
- расчетную ширину раскрытия трещин;
- допустимую ширину;
- коэффициент использования.

Вывод в книгу выполняется через `CCapacityResultWriter`:

- `WriteCrackMxResult`;
- `WriteCrackMyResult`;
- `WriteCrackMxyResult`.

Результаты записываются в соответствующие блоки `rngResultMx`, `rngResultMy`, `rngResultMxy`.

## Тесты

Добавлены тесты:

- режим `Concrete.TensionMode = Ignore`;
- режим `Concrete.TensionMode = UseDiagram`;
- расчет трещин для `Mx`;
- расчет трещин для `My`;
- расчет трещин для `Mxy`;
- запись crack-результатов в блок `Mxy`;
- случай без растянутой арматуры.

## Проверка

Команда:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-CrackTests.ps1
```

Фактический результат:

```text
TOTAL_CRACK: passed=28; failed=0
```

Полный регрессионный набор также пройден:

```text
Validate-Workbook.ps1: все проверки Passed=True
TOTAL: passed=29; failed=0
TOTAL_LINEAR: passed=49; failed=0
TOTAL_MATERIAL: passed=20; failed=0
TOTAL_SECTION_SOLVER: passed=34; failed=0
TOTAL_CAPACITY: passed=92; failed=0
TOTAL_CRACK: passed=28; failed=0
```

Для обновления книги использован узкий скрипт этапа 7:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Refresh-Stage07.ps1
```

## Ограничения

- Формулы раскрытия трещин пока используют временную параметризованную схему.
- Нормативные коэффициенты СП 35 и СП 63 должны быть уточнены и трассированы до инженерного выпуска.
- Расчет выполняется только для уже образовавшихся нормальных трещин.
- Пакетный расчет сочетаний не реализован.
