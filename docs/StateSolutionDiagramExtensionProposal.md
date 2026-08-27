# StateSolution и numerical extension: финальная архитектура

Статус документа: целевая архитектура для текущей реализации. Документ фиксирует финальные решения перед внесением изменений в VBA-код.

## 1. Назначение

`StateSolution` - отдельная расчетная цель для прямого поиска НДС заданного сочетания `N + Mx + My`. Она нужна для получения фактических `Stress/Strain`, записи snapshot на `Results`, построения схемы и экспорта в AutoCAD.

`StateSolution` не является расчетом несущей способности, не является расчетом момента образования трещин и не должен менять физические диаграммы материалов.

## 2. Настройка extension

На `Config` вводится настройка:

```text
Solver.DirectState.DiagramExtension = Yes / No
```

Значение по умолчанию: `Yes`.

Если настройка включена, только для прямого `StateSolution` строится численно расширенная диаграмма:

- `Eext = 0.01E`;
- бетон расширяется только в сжатии;
- арматура расширяется в растяжении и сжатии;
- реальные физические `epsilon_ult` не изменяются;
- технический диапазон extension ориентировочно `-10 ... +10`;
- extension является только численным механизмом поиска равновесия.

`Capacity`, `Mcrc`, `CrackedNDS` и другие физические расчеты extension не получают.

## 3. Purpose

Свободные строковые `purpose` заменяются на типизированный enum:

```vb
Public Enum ECalculationPurpose
    cpStrength = 1
    cpMcrc = 2
    cpCrackedNDS = 3
    cpStateSolution = 4
End Enum
```

В отдельном небольшом module/helper хранятся только преобразования и проверки:

```text
PurposeFromText(text)          -> ECalculationPurpose
PurposeToText(purpose)         -> String
IsPhysicalPurpose(purpose)     -> Boolean
IsStateBasePurpose(purpose)    -> Boolean
StateBasePurposeForGroup(type) -> ECalculationPurpose
```

Для `StateSolution` отдельно передается физическая база:

```text
Group1 -> cpStrength
Group2 -> cpCrackedNDS
```

Не создаются отдельные `StateSolutionStrength` или `StateSolutionCrackedNDS`.

`CSectionSolver` не знает о `purpose`, I/II ГПС, `Rb/Rb,ser`, `Rs/Rs,ser` и extension. Он получает уже готовые объекты диаграмм от `CMaterialModelProvider`.

## 4. Material pipeline

Поток остается таким:

```text
расчетный класс
  -> CMaterialModelProvider(purpose/basePurpose)
  -> CMaterialDiagram
  -> CSectionSolver
```

`CMaterialModelProvider`:

- строит физические диаграммы `cpStrength`, `cpMcrc`, `cpCrackedNDS`;
- для `cpStateSolution` берет физическую базу `cpStrength` или `cpCrackedNDS`;
- добавляет numerical extension только к state-диаграммам и только при включенной настройке;
- сохраняет физические пределы отдельно от добавленных технических точек.

`CMaterialDiagram` остается универсальной кусочно-линейной диаграммой и не знает о СП, бетоне, арматуре или расчетной цели. В нее добавляются только универсальные признаки:

```text
PhysicalCompressionStrain
PhysicalTensionStrain
HasCompressionExtension
HasTensionExtension
IsInPhysicalRange(epsilon)
IsInExtensionRange(epsilon)
```

Для диаграмм с отключенным растянутым бетоном положительные деформации бетона допустимы как зона с нулевым сопротивлением и не считаются extension.

## 5. Статусы Summary

В Summary используется общий короткий набор пользовательских статусов:

```text
OK
FAIL
NumFail
InputErr
N/A
```

Нужны отдельные статусы:

- `OverallStatus` - первый столбец Summary, итог по сочетанию;
- `DirectStateStatus` - статус прямого НДС;
- `CapacityStatus` - статус расчета несущей способности;
- `CrackStatus` - статус расчета трещин.

Подробные `StopReason` и диагностические сообщения solver-ов остаются внутренними и не смешиваются с короткими пользовательскими статусами.

### DirectStateStatus

```text
OK       solver сошелся, extension не использован, физические пределы соблюдены
FAIL     solver сошелся, но использован extension или физические пределы превышены
NumFail  solver не нашел равновесие
InputErr ошибка исходных данных
```

Пользовательский `StrainSafetyFactor` для DirectState больше не выводится. Количественный запас по прочности должен определяться расчетом `Capacity`.

### CapacityStatus

```text
OK       расчет выполнен, несущая способность достаточна
FAIL     расчет выполнен, несущая способность недостаточна
NumFail  численный поиск capacity не удался
InputErr ошибка исходных данных
N/A      расчет capacity не выполнялся в текущем режиме
```

Даже при `FullCapacity` фактическое состояние LC сначала решается как `StateSolution` с optional extension. Сам расчет несущей способности выполняется только по физической диаграмме `cpStrength`.

### CrackStatus

```text
OK       расчет трещин выполнен, проверка проходит или трещина не образована
FAIL     расчет трещин выполнен, раскрытие больше допустимого
NumFail  расчет трещин численно не выполнен
InputErr ошибка исходных данных
N/A      расчет трещин неприменим или не запускался
```

Для `Group2` расчет трещин запускается только если:

```text
DirectStateStatus = OK
ExtensionUsed = False
```

Если прямой `StateSolution` вышел в extension:

```text
DirectStateStatus = FAIL
ExtensionUsed = True
CrackStatus = N/A
```

`a_crc` и остальные результаты трещин в этом случае не рассчитываются и не выводятся. Сам факт включенной настройки `Solver.DirectState.DiagramExtension = Yes` не запрещает расчет трещин: решение принимается только по фактическому `ExtensionUsed` найденного состояния.

### OverallStatus

`OverallStatus` агрегирует только применимые расчетные разделы LC с приоритетом:

```text
InputErr -> NumFail -> FAIL -> OK
```

`N/A` не ухудшает общий статус, если соответствующий расчет просто не относится к данному LC или режиму. Если все разделы `N/A`, итоговый статус тоже `N/A`.

В `rngBatchSummary` вся строка LC окрашивается красным фоном при:

```text
OverallStatus = FAIL / NumFail / InputErr
```

`OK` и `N/A` получают нейтральное оформление.

## 6. Results snapshot

Для каждого LC в snapshot сохраняются как минимум:

```text
OverallStatus
DirectStateStatus
CapacityStatus
CrackStatus
ExtensionUsed
DirectState.DiagramExtensionEnabled
```

`ExtensionUsed` - машинный признак. Он нужен для расчета трещин, схемы, AutoCAD export и диагностики, но не является отдельным пользовательским статусом.

Все значения `Stress/Strain` в `Results` остаются значениями последнего успешно записанного snapshot. Схема и AutoCAD не пересчитывают их по текущему `Config`.

## 7. Схема и AutoCAD

Если в сохраненном snapshot для выбранного LC:

```text
ExtensionUsed = True
```

то схема и AutoCAD export все равно показывают найденное состояние `Stress/Strain`, но обязательно добавляют предупреждение:

```text
ВНЕ ФИЗИЧЕСКОЙ ДИАГРАММЫ МАТЕРИАЛА
```

Excel-схема выводит предупреждение под сечением красным жирным текстом. AutoCAD export добавляет аналогичную заметную русскую подпись.

Предупреждение определяется только по сохраненному `ExtensionUsed` из `Results`, а не по текущей настройке `Config`.

## 8. Справочник статусов

Справа от `rngBatchSummary` на листе `Results` выводится компактная таблица:

```text
Статус | Расшифровка
```

Содержимое:

| Статус | Смысл |
| --- | --- |
| `OK` | Расчетный раздел выполнен, физические пределы соблюдены, проверка проходит. |
| `FAIL` | Равновесие найдено, но физический предел нарушен, использован extension или проверка не проходит. |
| `NumFail` | Численное решение не найдено. Это не физический отказ, а проблема сходимости. |
| `InputErr` | Расчет не выполнен из-за ошибки исходных данных. |
| `N/A` | Расчетный раздел не относится к данному сочетанию или режиму. |

Тот же словарь добавляется на лист `Справка`.

## 9. Изменения по классам

### modCalculationPurpose

Новый общий module/helper:

- объявляет `ECalculationPurpose`;
- преобразует текстовые значения из Config/Results в enum;
- определяет физическую базу `StateSolution` по группе сочетания.

### CSystemSettingsReader

Добавить чтение `Solver.DirectState.DiagramExtension` с default `Yes`.

### CMaterialModelProvider

Изменить публичный контракт с текстового `purpose As String` на `purpose As ECalculationPurpose`.

Добавить методы:

```text
ConcreteStateMaterial(basePurpose As ECalculationPurpose)
SteelStateMaterial(basePurpose As ECalculationPurpose)
DirectStateDiagramExtensionEnabled
```

### CMaterialDiagram

Добавить хранение физического диапазона и проверку, использована ли искусственная область extension.

### CBatchSectionCalculator

Изменить расчет LC:

1. прямое состояние всегда решается как `StateSolution`;
2. physical base выбирается по группе LC;
3. после solve вычисляется `ExtensionUsed`;
4. формируются `DirectStateStatus`, `CapacityStatus`, `CrackStatus`, `OverallStatus`;
5. Group2 crack запускается только при `DirectStateStatus=OK` и `ExtensionUsed=False`;
6. `StrainSafetyFactor` больше не выводится пользователю.

### CBatchResultWriter

Перестроить `rngBatchSummary`:

- первый столбец `OverallStatus`;
- внутри блока текущего НДС первый столбец `DirectStateStatus`;
- внутри блока несущей способности первый столбец `CapacityStatus`;
- внутри блока трещин остается `CrackStatus`;
- справа вывести справочник статусов;
- красить строки с `OverallStatus = FAIL / NumFail / InputErr`.

### CNDMResultsWriter

Добавить LC-свойства snapshot:

```text
OverallStatus
DirectStateStatus
CapacityStatus
CrackStatus
ExtensionUsed
DirectState.DiagramExtensionEnabled
```

### CSectionPlotDataReader / CSectionPlotter

Reader читает `ExtensionUsed` из `rngNDMSectionProperties`. Plotter добавляет warning shape только по этому сохраненному признаку.

### modAutoCADStressExport

Экспорт читает `ExtensionUsed` из `rngNDMSectionProperties` и добавляет предупреждающий текст в AutoCAD. Текущий `Config` для этого не используется.

## 10. Справка

На листе `Справка` описать `Solver.DirectState.DiagramExtension`:

- extension - только численный механизм прямого StateSolution;
- extension не повышает физическую несущую способность;
- `Stress` в extension не является физическим напряжением материала;
- `Capacity` всегда считается по физической диаграмме;
- при `ExtensionUsed=True` прямой НДС получает `DirectStateStatus=FAIL`;
- для `Group2` при `ExtensionUsed=True` расчет трещин не выполняется;
- `NumFail` означает численную несходимость, а не физический отказ.

## 11. Проверки

После реализации проверить:

1. `cpStateSolution` не попадает в физические расчеты как самостоятельная база.
2. `Strength`, `Mcrc`, `CrackedNDS` всегда строятся без extension.
3. StateSolution получает extension только при включенной настройке.
4. Бетон расширяется только в сжатии, арматура - в сжатии и растяжении.
5. Физические ultimate strain не меняются после добавления extension.
6. Физически допустимое состояние получает `DirectStateStatus=OK`, `ExtensionUsed=False`.
7. Состояние в extension получает `DirectStateStatus=FAIL`, `ExtensionUsed=True`.
8. При `ExtensionUsed=True` для Group2 трещины не считаются, `CrackStatus=N/A`.
9. При включенной настройке, но фактическом `ExtensionUsed=False`, трещины считаются штатно.
10. Summary содержит `OverallStatus`, отдельные статусы блоков и справочник статусов.
11. Results snapshot содержит новые LC-свойства.
12. Excel-схема и AutoCAD показывают warning только по сохраненному `ExtensionUsed`.
