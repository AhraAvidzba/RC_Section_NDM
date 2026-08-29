# Архитектура расчета сочетаний NDM после разгрузки batch

Документ фиксирует фактическое состояние после практичного рефакторинга
варианта 2. Цель правки - оставить `CBatchSectionCalculator` оркестратором LC,
а численные сценарии прямого состояния и построение lambda-траектории вынести в
отдельные небольшие сущности.

## Текущий Pipeline

```text
Config
  -> CSystemSettingsReader
  -> CUnitSystem
  -> CLoadCombinationReader
  -> CSectionModel
  -> CMaterialModelProvider
  -> CBatchSectionCalculator
       -> CCapacityLoadPath
       -> CStateSolutionRunner -> CSectionSolver
       -> CCapacitySolver
       -> CCrackWidthCalculator
  -> CBatchResultWriter + CNDMResultsWriter
  -> Results snapshot
  -> Excel Plot / AutoCAD
```

Старый параллельный вывод определяющего сочетания на лист `Расчет` удален.
Численный результат теперь имеет один источник истины: batch summary и snapshot
на листе `Results`. Excel-схема и AutoCAD export читают сохраненный snapshot и
не запускают повторные solve-ы.

## Что Осталось В CBatchSectionCalculator

`CBatchSectionCalculator` по-прежнему хранит массивы результатов, потому что
именно он собирает batch summary и governing. Его основные обязанности:

- пройти по LC;
- выбрать, нужно ли запускать capacity и crack;
- вызвать `CStateSolutionRunner` для прямого НДС;
- вызвать `CCapacitySolver` с уже готовым `CCapacityLoadPath`;
- вызвать `CCrackWidthCalculator`, только если DirectState физически допустим;
- агрегировать `DirectStateStatus`, `CapacityStatus`, `CrackStatus`,
  `OverallStatus`;
- выбрать governing по прочности и трещинам;
- отдать writer-ам готовые данные.

В batch больше нет управления retry прямого solver-а, построения стартовых
extension-попыток и ручной сборки Offset/Base.

## CStateSolutionRunner

`CStateSolutionRunner` отвечает за один численный сценарий: найти прямое
состояние НДС от заданного `N + Mx + My`.

Внутри runner-а находятся:

- создание и настройка `CSectionSolver`;
- стартовая плоскость для чистого изгиба через `CStateGuessBuilder`;
- повторные попытки при включенном technical extension;
- проверка `ExtensionUsed`;
- проверка нахождения всех волокон/стержней в физических пределах диаграмм;
- компактный diagnostic log по retry-сценарию.

Runner не знает о `Group1/Group2`, не выбирает material purpose и не пишет
Results. Он получает уже готовые `CMaterialDiagram` от `CMaterialModelProvider`.

## CCapacityLoadPath

`CCapacityLoadPath` переводит пользовательский `CapacityLoadPath` в единую форму:

```text
Target(lambda) = Offset + lambda * Base
```

Класс хранит:

- `NOffset`, `NBase`;
- `MxOffset`, `MxBase`;
- `MyOffset`, `MyBase`;
- нормализованный ключ и подпись для Results;
- признаки `HasPath`, `HasScaledLoad`;
- признаки `ScalesN`, `ScalesMx`, `ScalesMy`;
- `ForceOnly` для силовой осевой траектории.

Смысл старого force-only special case сохранен, но больше не живет private-методом
batch-а. Теперь это свойство выбранной lambda-траектории, а не отдельная скрытая
ветка расчета.

## CStateGuessBuilder И ApplyPureBendingProbeGuess

`CStateGuessBuilder` остается builder-ом стартовых плоскостей. Он не запускает
solver и не принимает решения о статусах.

Он используется в двух местах:

1. `CStateSolutionRunner` - для прямого чистого изгиба и extension warm-start.
2. `CCapacitySolver.ApplyPureBendingProbeGuess` - для probe-точек
   `LoadMultiplier` при чистом изгибе.

`ApplyPureBendingProbeGuess` оставлен в `CCapacitySolver`, потому что это локальная
подсказка именно для внутренней probe-точки capacity-поиска. Переносить ее в
`CStateSolutionRunner` неправильно: runner управляет прямым пользовательским
StateSolution, а не одномерными пробными точками capacity.

## Что Удалено

- старый отдельный result-writer для листа `Расчет`;
- повторный пересчет governing LC после batch;
- отдельный direct-state initial guess для старой ветки;
- дублирующее построение lambda-траектории вне batch;
- имя старого result-блока в сборке книги и проверке workbook layout.

## Оставшиеся Границы Ответственности

`CBatchSectionCalculator` все еще остается большим классом, потому что хранит
широкий snapshot результатов для writer-ов. Это приемлемо для текущего проекта:
численная логика уже вынесена, а дальнейшее дробление на scenario-runner-ы
увеличило бы число классов без немедленного выигрыша.

Следующий разумный шаг, если класс снова начнет разрастаться, - выносить не
математику, а структуры хранения/вывода результатов batch. Пока это не требуется.
