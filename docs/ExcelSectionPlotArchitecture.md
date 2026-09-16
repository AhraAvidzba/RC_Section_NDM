# Архитектура Отрисовки Сечения В Excel

Дата актуализации: 2026-08-19.

## Статус

Документ описывает целевую архитектуру будущей отрисовки расчетного ЖБ-сечения на листе Excel. На этом этапе код VBA, книга Excel, листы `Config`/`Results` и расчетная модель не изменялись.

Учтены:

- фактическая архитектура проекта с `CSectionModel`;
- требования из `Plotter_prompt.txt`;
- решения по выводу расчетного снимка на лист `Results`;
- финальное решение хранить semantic annotations на `Results`, а не только в памяти.

## Главный Принцип

Визуализация не является частью расчетного ядра.

```text
INPUT
  -> CUnitSystem / Input Adapter
  -> фиксированная INTERNAL-система
  -> Solver
  -> CUnitSystem / Output Adapter
  -> OUTPUT snapshot
  -> Results
  -> CSectionPlotDataReader
  -> CSectionPlotter
  -> Excel Chart + Shapes
```

`CSectionPlotter` только отображает уже сохраненные данные. Он не пересчитывает геометрию, напряжения, деформации, главные оси, точку приложения нагрузки или армирование.

Запрещенный поток:

```text
CSectionPlotter -> CSectionSolver
CSectionPlotter -> CCapacitySolver
CSectionPlotter -> CCrackWidthCalculator
CSectionPlotter -> BuildWorkbookSectionModel
CSectionPlotter -> CAutoCADSectionModelImporter
CSectionPlotter -> CFiberMeshBuilder
CSectionPlotter -> CRebarLayout
```

## Последний Расчетный Снимок

Лист `Results` должен хранить неизменяемый `OUTPUT`-снимок последнего успешно выполненного расчета:

```text
                    LAST CALCULATED SNAPSHOT
                             |
       +-------------+--------------+--------------+--------------+
       |             |              |              |              |
rngNDMElementResults  rngNDMSectionGeometry  rngNDMSectionProperties  rngNDMMaterialDiagrams  rngNDMSectionAnnotations
       |             |              |              |              |
       +-------------+--------------+--------------+--------------+
                             |
                    CSectionPlotDataReader
                             |
                       CSectionPlotter
                             |
                    Excel Chart + Shapes
```

При новом расчете все блоки snapshot обновляются согласованно.

При нажатии кнопки `Обновить схему` эти блоки не изменяются. Кнопка только читает сохраненный снимок.

Если пользователь после расчета изменил геометрию, армирование, материалы, `OUTPUT units` или `OUTPUT sign convention` на листе `Config`, но новый расчет не выполнил, содержимое `Results` остается прежним. Схема должна продолжать показывать старое рассчитанное сечение в тех единицах и знаках, которые были активны при последнем расчете.

Новые входные данные, единицы и знаки попадают в `Results` и на схему только после нового расчета.

Пример:

```text
Расчет выполнен:
  B = 1000
  арматура = 6Ø32

Results содержит:
  ElementResults для B = 1000
  SectionProperties для B = 1000
  SectionAnnotations с подписью 6Ø32

Пользователь изменил входные данные:
  B = 1200
  арматура = 8Ø32

Кнопка "Обновить схему" без нового расчета:
  показывает B = 1000 и 6Ø32

Только после нового расчета:
  показывает B = 1200 и 8Ø32
```

Пример с единицами:

```text
Расчет выполнен при:
  StressUnit = MPa
  ForceUnit = kN
  +N = Compression

После расчета пользователь изменил System:
  StressUnit = kPa
  ForceUnit = tf
  +N = Tension

Без нового расчета:
  Results не перезаписывается
  Обновить схему не преобразует единицы
  Обновить схему не меняет знаки
  схема продолжает использовать MPa, kN и старую sign convention из snapshot
```

## Блоки Results

Сохранить существующий `rngNDMSectionGeometry` и добавить к нему новые логически разделенные блоки:

```text
rngNDMSectionGeometry
rngNDMElementResults
rngNDMSectionProperties
rngNDMSectionAnnotations
```

Не объединять их в одну таблицу.

Итоговое разделение:

```text
rngNDMSectionGeometry
  -> фактическая геометрия и арматура, один раз на расчетный snapshot

rngNDMElementResults
  -> только LC-зависимые результаты элементов

rngNDMSectionProperties
  -> общие свойства сечения и LC-зависимые свойства уровня сечения

rngNDMSectionAnnotations
  -> размеры и смысловые подписи арматуры
```

Связь между геометрией и результатами выполнять по `ElementID`:

```text
rngNDMSectionGeometry.ElementID
  <-> rngNDMElementResults.ElementID + LoadCase
```

`CSectionPlotDataReader` при построении схемы должен объединять геометрию из `rngNDMSectionGeometry` с результатами выбранного `LoadCase` из `rngNDMElementResults`.

`rngNDMRunMetadata` больше не использовать. Новое имя: `rngNDMSectionProperties`.

`PlotMetadata` как концепция данных только в памяти больше не используется. Новое имя для сохраняемой semantic-информации: `SectionAnnotations`, диапазон `rngNDMSectionAnnotations`.

## rngNDMSectionGeometry

Назначение:

```text
неизменяемые данные расчетных элементов, записанные один раз для расчетного snapshot
```

Существующий диапазон `rngNDMSectionGeometry` не удалять и не заменять.

Текущий код уже использует этот диапазон:

- `CNDMResultsWriter.WriteSectionGeometry` записывает геометрию один раз после расчета;
- `modAutoCADStressExport.ReadSectionGeometryFromResults` читает геометрию для AutoCAD export из `rngNDMSectionGeometry`.

Архитектура должна максимально переиспользовать эту работающую структуру.

Содержит постоянные свойства элементов:

- `ElementID`;
- `Type` или `MaterialType`;
- `X`;
- `Y`;
- `Area`;
- `Diameter` для арматуры;
- `Material`;
- `GeometryInterpretationStatus`;
- `Width`;
- `Height`;
- `Rotation`;
- локальные моменты инерции, если они нужны AutoCAD export или другим существующим функциям;
- другие геометрические/постоянные свойства.

Все численные значения в `rngNDMSectionGeometry` записываются уже в `OUTPUT units`, выбранных на момент расчета. Если для какого-либо будущего геометрического поля появится пользовательская sign convention, оно также должно записываться уже в конечном `OUTPUT`-виде.

Геометрию не дублировать в `rngNDMElementResults` для каждого `LoadCase`.

## rngNDMElementResults

Назначение:

```text
как каждый расчетный элемент работает в данном сочетании
```

Содержит только LC-зависимые результаты:

- `ElementID`;
- `LoadCase`;
- `Strain`;
- `Stress`;
- `PhysicalState`;
- `N`, `Mx`, `My` при необходимости;
- другие действительно LC-зависимые результаты.

Этот блок LC-зависимый: для каждого рассчитанного сочетания в нем могут быть строки по тем же расчетным элементам с другим НДС.

Все численные значения в `rngNDMElementResults` записываются уже в `OUTPUT units` и `OUTPUT sign convention`, выбранных на момент расчета.

`Epsilon0`, `KappaX`, `KappaY` в этом блоке не хранить. Это свойства плоскости деформаций всего сечения для конкретного `LoadCase`, а не свойства отдельного элемента. Для каждого элемента достаточно хранить его фактические `Strain`/`Stress` в сохраненном OUTPUT-снимке и физическую классификацию `PhysicalState`.

Не дублировать в `rngNDMElementResults`:

- `X`;
- `Y`;
- `Area`;
- `Diameter`;
- `Material`;
- `GeometryInterpretationStatus`;
- другие постоянные свойства, которые уже однозначно доступны через `rngNDMSectionGeometry`.

Не хранить `INTERNAL`-копии `Stress/Strain` только ради plotter-а. Если какие-либо внутренние значения уже нужны другим существующим функциям `Results`, их не удалять без отдельного анализа, но `Обновить схему` не должен использовать их для повторного `OUTPUT`-преобразования.

Численные подписи результата на первом этапе строятся только по бетонным элементам. Численные значения отдельных арматурных стержней автоматически не подписывать.

## rngNDMSectionProperties

Назначение:

```text
общие численные свойства рассчитанного сечения и расчетного снимка
```

Это не поэлементные данные и не semantic-оформление.

Все численные значения в этом блоке записываются в `OUTPUT units` и `OUTPUT sign convention`, активных на момент расчета.

Рекомендуемый формат:

```text
RunID | LoadCase | Parameter | Value | Unit | Comment
```

Для свойств, не зависящих от сочетания, `LoadCase` оставлять пустым или писать `ALL`.

Общие свойства:

| Parameter | Unit | Комментарий |
|---|---|---|
| `Bounds.MinX` | output length | Минимальный X фактической расчетной модели |
| `Bounds.MaxX` | output length | Максимальный X фактической расчетной модели |
| `Bounds.MinY` | output length | Минимальный Y фактической расчетной модели |
| `Bounds.MaxY` | output length | Максимальный Y фактической расчетной модели |
| `Transformed.CentroidX` | output length | Центр тяжести приведенного расчетного сечения |
| `Transformed.CentroidY` | output length | Центр тяжести приведенного расчетного сечения |
| `Transformed.PrincipalAngle` | rad | Угол от +X к главной оси 1 приведенного сечения; положительный против часовой стрелки |
| `Transformed.PrincipalI1` | output length^4 | Главный момент инерции I1 приведенного сечения |
| `Transformed.PrincipalI2` | output length^4 | Главный момент инерции I2 приведенного сечения |
| `Output.LengthUnit` | - | Единица вывода длины |
| `Output.AreaUnit` | - | Единица вывода площади |
| `Output.ForceUnit` | - | Единица вывода силы |
| `Output.MomentUnit` | - | Единица вывода момента |
| `Output.StressUnit` | - | Единица вывода напряжений |
| `Output.CurvatureUnit` | - | Единица вывода кривизны |
| `Output.SignConvention.N` | - | Пользовательский знак вывода N на момент последнего расчета |
| `Output.SignConvention.Mx` | - | Пользовательский знак вывода Mx на момент последнего расчета |
| `Output.SignConvention.My` | - | Пользовательский знак вывода My на момент последнего расчета |

LC-зависимые свойства:

| Parameter | Unit | Комментарий |
|---|---|---|
| `Epsilon0` | - | Деформация в начале координат расчетной модели для выбранного LC |
| `KappaX` | output curvature | Кривизна по X для выбранного LC |
| `KappaY` | output curvature | Кривизна по Y для выбранного LC |
| `LoadReferenceX` | output length | X точки приложения нагрузки для выбранного LC, если она зависит от LC |
| `LoadReferenceY` | output length | Y точки приложения нагрузки для выбранного LC, если она зависит от LC |

Generated и AutoCAD geometry должны формировать одинаковый по смыслу набор `SectionProperties`, если соответствующие данные существуют.

`CSectionPlotter` должен использовать сохраненные в `rngNDMSectionProperties` единицы и sign convention snapshot для легенды, подписей, размеров, обозначения выбранной величины и других выводимых численных значений. Он не должен брать текущие `OUTPUT` settings из `Config`, потому что они могли быть изменены после расчета.

`Output.SignConvention.*` в `rngNDMSectionProperties` является только описанием snapshot: оно фиксирует, в какой пользовательской системе знаков уже записаны значения на `Results`. Не применять эту sign convention повторно к уже записанным `OUTPUT`-значениям.

Например, `Stress`, `N`, `Mx`, `My` на `Results` уже имеют конечный пользовательский знак. При построении схемы их нужно читать как готовые значения.

`Bounds` должны существовать для любого источника геометрии.

Не дублировать:

```text
B = Bounds.MaxX - Bounds.MinX
H = Bounds.MaxY - Bounds.MinY
```

как отдельные свойства без необходимости. Если plotter нужны общие габариты, он получает их из `Bounds`.

Не хранить:

```text
NeutralLine.A
NeutralLine.B
NeutralLine.C
```

если уже сохраняются:

```text
Epsilon0
KappaX
KappaY
```

Нейтральная линия однозначно получается:

```text
epsilon(x, y) = epsilon0 + kappaX * y + kappaY * x
kappaY * x + kappaX * y + epsilon0 = 0
```

То есть:

```text
A = kappaY
B = kappaX
C = epsilon0
```

Формирование `A/B/C` в `CSectionPlotDataReader` или `CSectionPlotter` не является повторным расчетом сечения. Это только другое представление уже сохраненной плоскости деформаций.

`CSectionPlotDataReader` берет `Epsilon0`, `KappaX`, `KappaY` только из `rngNDMSectionProperties` и по ним строит нейтральную линию. Не дублировать эти значения в каждой строке `rngNDMElementResults`.

`GeometrySource` хранить только если он реально нужен для диагностики или контроля целостности. `CSectionPlotter` не должен использовать `GeometrySource` для ветвления логики.

`ConcreteCount` и `RebarCount` хранить только если они действительно полезны для контроля целостности или существующих функций. Не хранить легко вычисляемые значения просто “на всякий случай”.

## rngNDMSectionAnnotations

Назначение:

```text
что и где нужно дополнительно обозначить на схеме
```

Это сохраненная semantic-информация оформления рассчитанной геометрии. Она не является расчетным результатом.

Здесь хранить только ту семантику, которую невозможно надежно восстановить из обычных расчетных результатов.

Минимальный формат может быть табличным:

```text
RunID | AnnotationType | AnnotationID | Field | Value | Unit | Comment
```

или широким по типам annotation, если это проще для VBA. Важно, чтобы блок оставался отдельным от `rngNDMElementResults` и `rngNDMSectionProperties`.

### DIMENSION

Для размерных обозначений хранить:

- `StartX`;
- `StartY`;
- `EndX`;
- `EndY`;
- `OutsideNormalX`;
- `OutsideNormalY`;
- `Offset`;
- `Text`;
- `Value`;
- `Unit`;
- другие минимально необходимые параметры.

Эта запись соответствует будущему классу или структуре:

```text
CPlotDimension
  StartX
  StartY
  EndX
  EndY
  OutsideNormalX
  OutsideNormalY
  Offset
  Text
  Value
  Unit
```

`CSectionPlotter` только рисует `CPlotDimension`. Он не знает, какой размер является стенкой, полкой, высотой, шириной или диаметром.

### REBAR ANNOTATION

Для групповых подписей арматуры хранить:

- `AnchorX`;
- `AnchorY`;
- `DirectionX`;
- `DirectionY`;
- `OutsideNormalX`;
- `OutsideNormalY`;
- `Offset`;
- готовую смысловую подпись `Text`, например `6Ø32` или `6Ø32 + 3Ø20`;
- при необходимости `EdgeName` только как трассировочное поле, не для логики plotter-а.

Класс можно назвать:

```text
CPlotRebarLabel
```

если он описывает именно визуальную подпись. Если при реализации выяснится, что объект хранит более широкую semantic-информацию, лучше имя:

```text
CRebarAnnotation
```

Широко используемые существующие сущности не переименовывать ради косметики. На текущем этапе эти классы еще не реализованы, поэтому предпочтительный термин для сохраняемых данных - `SectionAnnotations`.

## Не Смешивать Блоки

Пример для параметрического `LShape`:

`rngNDMSectionProperties`:

- `Bounds.MinX`;
- `Bounds.MaxX`;
- `Bounds.MinY`;
- `Bounds.MaxY`;
- `Transformed.CentroidX`;
- `Transformed.CentroidY`;
- `Transformed.PrincipalAngle`;
- `Transformed.PrincipalI1`;
- `Transformed.PrincipalI2`.

`rngNDMSectionAnnotations`:

- Dimension `B`;
- Dimension `H`;
- Dimension `толщина стенки`;
- Dimension `толщина полки`;
- Rebar annotation `6Ø32`;
- Rebar annotation `6Ø32 + 3Ø20`.

Пример для AutoCAD:

`rngNDMSectionProperties`:

- `Bounds.MinX`;
- `Bounds.MaxX`;
- `Bounds.MinY`;
- `Bounds.MaxY`;
- `Transformed.CentroidX`;
- `Transformed.CentroidY`;
- `Transformed.PrincipalAngle`, если рассчитан;
- `Transformed.PrincipalI1`, если рассчитан;
- `Transformed.PrincipalI2`, если рассчитан.

`rngNDMSectionAnnotations`:

- может быть пустым.

В этом случае `CSectionPlotter` при включенных размерах может автоматически показать только общие `B/H` по `Bounds`.

## Load Case

Переключение `Plot.LoadCase` не создает заново геометрию и `SectionAnnotations`.

Геометрия, арматура, размеры и semantic annotations относятся к одному расчетному запуску и остаются неизменными.

Для выбранного `LoadCase` меняется только сохраненное напряженно-деформированное состояние:

- `Stress`;
- `Strain`;
- `Epsilon0`;
- `KappaX`;
- `KappaY`;
- нейтральная линия, построенная из `Epsilon0/KappaX/KappaY`;
- точка приложения нагрузки, если она зависит от сочетания;
- другие LC-dependent результаты.

Идея:

```text
тот же Section Snapshot + другое сохраненное НДС
```

Поэтому переключение сочетания должно быть легкой операцией чтения уже сохраненных строк `Results`.

## AutoCAD

Для AutoCAD-imported geometry действует та же структура `Results`, что и для Generated geometry. Не создавать специальную структуру результатов для AutoCAD.

`rngNDMSectionGeometry`:

- `ElementID`;
- `Type`;
- `X/Y`;
- `Area`;
- `Diameter`;
- `Material`;
- `GeometryInterpretationStatus`;
- другие постоянные данные геометрии, если они реально нужны plotter-у или AutoCAD export.

`rngNDMElementResults`:

- `ElementID`;
- `LoadCase`;
- `Stress`;
- `Strain`;
- `PhysicalState`;
- другие действительно LC-dependent результаты.

`rngNDMSectionProperties`:

- `Bounds`;
- центр тяжести;
- главные оси;
- единицы вывода;
- другие общие свойства.

`rngNDMSectionAnnotations`:

- только автоматически надежно определяемые annotations.

При отсутствии semantic-информации о гранях и группах арматуры:

- не создавать специальные `RebarAnnotations`;
- не угадывать характерные размеры;
- не создавать ветку `CSectionPlotter` специально для AutoCAD.

Общие `B/H` могут отображаться по `Bounds`.

AutoCAD export после введения этих блоков также должен читать последний расчетный снимок с `Results`, а не пересчитывать свойства сечения и не импортировать геометрию повторно.

Настройки AutoCAD export должны включать:

| Параметр | По умолчанию | Варианты | Комментарий |
|---|---:|---|---|
| `AutoCAD.Export.ResultType` | `Stress` | `Stress`, `Strain` | Какая величина используется для подписей/раскраски экспортируемых результатов. |

`AutoCAD.Export.ResultType` использует тот же смысл, что и `Plot.ResultType`, но является отдельной настройкой: пользователь может смотреть в Excel деформации, а в AutoCAD экспортировать напряжения, или наоборот.

Экспортер не должен иметь отдельные ветки “stress exporter” и “strain exporter”. Reader/helper должен выбрать поле `Stress` или `Strain` из `rngNDMElementResults` и передать в общую логику экспорта универсальное значение `ResultValue`.

## Единицы И Система Знаков

В проекте поддерживается пользовательская система знаков для ввода и вывода, но расчетное ядро работает в одной фиксированной `INTERNAL` convention.

Архитектурный источник истины для пересчета единиц и пользовательских знаков уже существует:

```text
CUnitSystem
```

Он читает:

```text
rngUnitSettings
rngSignConventionSettings
```

и выполняет преобразования:

```text
INPUT -> INTERNAL
INTERNAL -> OUTPUT
```

Архитектурное решение для `Results`: принят вариант 1.

```text
Results = неизменяемый OUTPUT snapshot последнего успешно выполненного расчета
```

Преобразование `INTERNAL -> OUTPUT` выполняется при записи `Results`, то есть в writer/output-adapter слое:

```text
Solver
  -> CUnitSystem / Output Adapter
  -> CBatchResultWriter / CNDMResultsWriter / другие writer-ы
  -> Results
```

Пользователь должен видеть на `Results` именно те единицы и знаки, которые были выбраны в `Config` перед запуском расчета.

Примеры значений на `Results`:

```text
X = 800 mm
Area = 625 mm2
Stress = +18.4 MPa
N = -1250 kN
Mx = 384 kN*m
```

если именно такие `OUTPUT units` и `OUTPUT sign convention` были активны при расчете.

`Обновить схему` не выполняет повторное преобразование единиц или знаков. Оно использует численные значения и единицы, сохраненные в `Results`.

Если после расчета пользователь изменил `OUTPUT units` или `OUTPUT sign convention`, это не меняет ни `Results`, ни схему. Новые единицы и знаки начинают применяться только после следующего выполнения расчета.

Физическая классификация элемента не должна зависеть от пользовательского знака вывода. Цвет элемента не должен случайно меняться только потому, что пользователь выбрал другую `OUTPUT` convention.

`CSectionPlotter` не должен самостоятельно определять физический смысл состояния через жестко заданные проверки:

```text
Strain > 0 = tension
Strain < 0 = compression
Stress > 0 = tension
Stress < 0 = compression
```

и не должен зависеть от `INPUT` sign convention.

Единый источник физической классификации для plotter-а:

```text
CSectionPlotDataReader
```

Reader должен читать из `rngNDMElementResults` сохраненное поле:

```text
PhysicalState
```

Это поле формируется во время расчетного запуска расчетным/writer-слоем на основании фиксированной `INTERNAL` convention проекта, типа материала, активной материальной модели расчетного режима и фактических внутренних `Stress/Strain`, до потери физического смысла из-за пользовательского `OUTPUT`-знака.

Допустимые значения первого этапа:

```text
Compression
Tension
InactiveTensionConcrete
NearZero
```

Смысл:

- `Compression` - элемент физически работает в сжатии;
- `Tension` - элемент физически работает в растяжении и его напряжение учитывается расчетной моделью;
- `InactiveTensionConcrete` - бетон физически растянут, но растягивающее напряжение бетона отключено активной моделью выбранного расчетного режима, например `CrackedNDS` с фиксированным `Ignore`;
- `NearZero` - напряжение/деформация около нуля в пределах принятого допуска.

`CSectionPlotter` получает уже готовый `PhysicalState` и только выбирает цвет/маркер. Он не знает, какой численный знак в `INTERNAL` или `OUTPUT` означает растяжение или сжатие.

Так как кнопка `Обновить схему` не должна обращаться к текущим настройкам материалов, итоговый `PhysicalState` должен быть сохранен в `rngNDMElementResults` при расчетном запуске. `CSectionPlotDataReader` читает это состояние из snapshot, может проверить его целостность и передает в plotter. Повторная классификация по текущему `Config` при обновлении схемы запрещена.

Допустима ситуация:

```text
Stress = +18 MPa
PhysicalState = Compression
```

если пользовательская `OUTPUT convention` на момент расчета задавала положительный знак для сжатия.

Таким образом:

- `CUnitSystem` отвечает за пользовательские единицы и пользовательские знаки чисел при чтении input и записи output;
- writer-слой отвечает за запись `OUTPUT snapshot` и сохранение `PhysicalState`;
- `CSectionPlotDataReader` отвечает за чтение сохраненного snapshot;
- `CSectionPlotter` отвечает только за отображение.

Перед реализацией plotter-а не требуется заново выяснять, какой знак `Strain` означает растяжение: это должно быть один раз явно зафиксировано в классификаторе расчетного/writer-слоя через `INTERNAL` convention проекта и покрыто тестами.

## Размеры Сечения

Для параметрических сечений показывать только характерные геометрические размеры, которые имеют смысл на инженерном чертеже.

Для Г-образного сечения:

- общая ширина `B`;
- общая высота `H`;
- толщина вертикальной части;
- толщина горизонтальной полки.

Для круга:

- диаметр `ØD`.

Для AutoCAD-сечения без semantic annotations:

- общий `B` по `Bounds.MinX/Bounds.MaxX`;
- общий `H` по `Bounds.MinY/Bounds.MaxY`.

Размеры располагать снаружи сечения:

- выносные линии от соответствующих точек или граней;
- размерная линия параллельно измеряемому направлению;
- значение размера по центру;
- наконечники-стрелки с обеих сторон через `BeginArrowheadStyle` и `EndArrowheadStyle`;
- несколько размеров автоматически разносить по уровням, чтобы текст и линии не пересекались между собой, с сечением и с арматурными подписями.

Визуальный порядок снаружи:

```text
сечение -> подпись арматуры -> размерная линия
```

## Подписи Арматуры По Граням

`Plot.RebarLabels.Enabled` означает групповые подписи арматуры по смысловым граням. Отдельные стержни не подписывать.

Примеры:

```text
6Ø32
4Ø20
6Ø32 + 3Ø20
```

Для каждой смысловой грани параметрического сечения должна быть одна компактная подпись снаружи сечения, примерно напротив середины грани и параллельно этой грани.

Ориентация:

- горизонтальная грань - горизонтальный текст;
- вертикальная грань - текст повернут на 90 градусов;
- наклонная грань - по возможности текст параллельно грани.

`CSectionPlotter` не должен по координатам определять, какие стержни относятся к грани или ряду. Эти semantic annotations должен формировать генератор геометрии/арматуры, потому что именно он знает принадлежность стержней к граням, рядам и схему расстановки.

Архитектура должна поддерживать несколько рядов, разные диаметры и будущую расстановку второго/третьего ряда через один стержень первого ряда без изменения логики `CSectionPlotter`.

Для AutoCAD-сечения при отсутствии semantic annotations групповые подписи арматуры не выводить.

## Настройки System

Отдельный именованный диапазон `rngPlotSettings` не создавать. Блок настроек схемы находится внутри общего `rngSystemSettings`.

Добавить блок:

```text
[Excel plot]
```

Предлагаемый набор:

| Параметр | По умолчанию | Варианты | Комментарий |
|---|---:|---|---|
| `Plot.Enabled` | `Yes` | `Yes`, `No` | Разрешает обновление схемы Excel. Не влияет на расчет. |
| `Plot.LoadCase` | `Worst` | `Worst` + динамический список LC | Какое сочетание показывать на схеме. |
| `Plot.ResultType` | `Stress` | `Stress`, `Strain` | Какая величина используется для цветового градиента, легенды и пространственных численных подписей. |
| `Plot.ResultGradient` | `Yes` | `Yes`, `No` | Раскраска элементов по фактическому состоянию и выбранному `Plot.ResultType`. |
| `Plot.ResultLabelsEnabled` | `Yes` | `Yes`, `No` | Показывать пространственно распределенные подписи выбранного `Plot.ResultType` по бетонным элементам. |
| `Plot.ResultLabelSpacing` | `100` | длина в output units | Целевой пространственный шаг между подписями выбранного результата, а не “каждый N-й ElementID”. |
| `Plot.NeutralLineEnabled` | `Yes` | `Yes`, `No` | Показывать нейтральную линию выбранного LC. |
| `Plot.PrincipalAxesMode` | `Transformed` | `Transformed`, `Concrete`, `None` | Какие главные центральные оси показывать: приведенного сечения, бетонного сечения или не показывать. |
| `Plot.LoadApplicationPointEnabled` | `Yes` | `Yes`, `No` | Показывать точку приложения нагрузки. |
| `Plot.Dimensions.Enabled` | `Yes` | `Yes`, `No` | Показывать характерные размеры из `rngNDMSectionAnnotations`; для AutoCAD без annotations - только `B/H` по `Bounds`. |
| `Plot.RebarLabels.Enabled` | `Yes` | `Yes`, `No` | Показывать групповые подписи арматуры по граням, например `6Ø32 + 3Ø20`. |
| `Plot.Dimensions.Placement` | `Outside` | `Outside`, `Inside` | Сторона текста размера относительно размерной линии; сама размерная линия остается на своей semantic-стороне. |
| `Plot.RebarLabels.Placement` | `Outside` | `Outside`, `Inside` | Сторона подписи арматуры относительно линии осей стержней. |
| `Plot.Dimensions.Offset`, `Plot.RebarLabels.Offset` | `100` / `60` | мм сечения | Геометрический отступ задается в реальных миллиметрах сечения: для размеров от грани до размерной линии, для арматуры от линии осей стержней. |
| `Plot.Dimensions.TextUnits`, `Plot.RebarLabels.TextUnits` | `pt` | `mm`, `pt` | Единицы для `TextHeight` и `TextGap`. `pt` задает обычные Excel points и стабилизирует визуальный размер текста для сечений разных габаритов, `mm` сохраняет модельное масштабирование. |
| `Plot.Dimensions.TextHeight`, `Plot.RebarLabels.TextHeight` | `13` / `13` | по `TextUnits` | Высота текста: модельные миллиметры сечения или фиксированный размер шрифта Excel в points. |
| `Plot.Dimensions.TextGap`, `Plot.RebarLabels.TextGap` | `9` / `9` | по `TextUnits` | Зазор между линией аннотации и текстом: модельные миллиметры сечения или фиксированный экранный зазор в points. |
| `Plot.RebarLabels.LineEnabled` | `Yes` | `Yes`, `No` | Показывать короткую линию обозначения арматуры; при `No` остается только текст. |
| `Plot.Dimensions.ArrowType` | `Triangle` | `Triangle`, `Stealth`, `Diamond`, `Oval`, `Open` | Тип наконечников размерной линии. |
| `Plot.Dimensions.ArrowSize` | `Wide` | `Small`, `Medium`, `Wide` | Размер наконечников размерной линии. |
| `Plot.LegendEnabled` | `Yes` | `Yes`, `No` | Показывать легенду выбранной величины справа от схемы. |

Не использовать:

```text
Plot.LabelMode
Plot.CombinationID
Plot.LoadPointEnabled
Plot.LoadApplicationPoint
Plot.StressGradient
Plot.StressLabelsEnabled
Plot.StressLabelSpacing
AutoCAD.LabelMode
```

Использовать единый понятный нейминг:

```text
Plot.LoadCase
Plot.ResultType
Plot.ResultGradient
Plot.ResultLabelsEnabled
Plot.ResultLabelSpacing
Plot.LoadApplicationPointEnabled
AutoCAD.Export.ResultType
AutoCAD.Export.LabelMode
```

Excel plot не должен копировать `AutoCAD.Export.LabelMode` ради симметрии: в AutoCAD нужны подписи значений/имен элементов, а в Excel-схеме нужны spatial result labels и групповые подписи арматуры.

## Динамический Список LoadCase

`Plot.LoadCase` должен иметь динамический список:

```text
Worst
LC1
LC2
...
```

Источник - сохраненные сочетания последнего расчета или `rngLoadCombinations`, если список нужен до расчета. При отрисовке приоритет имеет наличие LC в последнем расчетном снимке `Results`: нельзя показывать LC, которого нет в сохраненном снимке.

## Лист Расчет

Схему выводить на существующем листе `Расчет`.

Добавить именованный диапазон-якорь:

```text
chtNDMSectionPlot
```

Кнопка:

```text
Обновить схему
```

Макрос кнопки:

```text
UpdateSectionPlot
```

Макрос:

- читает настройки `Plot.*` из `Config` только как настройки представления уже сохраненного snapshot;
- читает последний расчетный снимок из `Results`;
- не обращается к текущим входным данным геометрии/армирования;
- не запускает расчет;
- не импортирует AutoCAD;
- не преобразует `Results` в текущие `OUTPUT units`;
- не меняет sign convention сохраненных результатов;
- не изменяет `rngNDMElementResults`, `rngNDMSectionProperties`, `rngNDMSectionAnnotations`.

Без нового расчета через настройки `Plot.*` можно менять только способ представления уже сохраненного snapshot:

- `Plot.LoadCase`;
- `Plot.ResultType`;
- `Plot.ResultGradient`;
- `Plot.ResultLabelsEnabled`;
- `Plot.ResultLabelSpacing`;
- `Plot.NeutralLineEnabled`;
- `Plot.LoadApplicationPointEnabled`;
- `Plot.PrincipalAxesMode`;
- `Plot.Dimensions.Enabled`;
- `Plot.RebarLabels.Enabled`.

Эти настройки не изменяют `Results`.

## Выбор Отображаемой Величины

Сразу поддержать универсальный выбор результата:

```text
Plot.ResultType = Stress
Plot.ResultType = Strain
```

При `Plot.ResultType = Stress`:

- цветовой градиент строится по `Stress`;
- пространственные численные подписи строятся по `Stress`;
- легенда и формат чисел используют сохраненную `Output.StressUnit` из `rngNDMSectionProperties`.

При `Plot.ResultType = Strain`:

- цветовой градиент строится по `Strain`;
- пространственные численные подписи строятся по `Strain`;
- легенда и формат чисел используют безразмерную деформацию из сохраненного snapshot.

Вся остальная схема не меняется:

- геометрия;
- арматура;
- `PhysicalState`;
- нейтральная линия;
- размеры;
- групповые подписи арматуры;
- центр тяжести;
- главные оси;
- точка приложения нагрузки.

Не создавать две реализации plotter-а для `Stress` и `Strain`. `CSectionPlotDataReader` должен выбрать числовое поле результата по `Plot.ResultType` и передать в `CSectionPlotter` единый набор:

```text
ResultValue
ResultUnit
ResultCaption
PhysicalState
```

`CSectionPlotter` строит gradient, labels и legend по `ResultValue`, не зная, является ли это напряжением или деформацией.

Аналогичный выбор добавить для AutoCAD export:

```text
AutoCAD.Export.ResultType = Stress
AutoCAD.Export.ResultType = Strain
```

Экспортер должен использовать выбранную величину для тех элементов визуализации и подписей, где сейчас используется `Stress`. Общую логику выбора `ResultType` желательно вынести в небольшой helper/reader, чтобы Excel plot и AutoCAD export не дублировали код без необходимости.

## Тип Визуализации

Использовать связку:

```text
XY Scatter Chart + Shapes внутри Chart
```

Почему:

- `XY Scatter` хорошо держит численную систему координат;
- `Shapes` удобны для размерных линий, подписей, стрелок, маркеров центра, точки нагрузки и нейтральной линии;
- можно добиться инженерного вида без создания тяжелой фигуры на каждый бетонный элемент.

Все автоматически создаваемые объекты схемы NDMProgramm должны создаваться внутри соответствующего `Chart`:

```text
Chart.Shapes
```

а не на листе Excel.

`ChartObject` должен быть единым переносимым графическим объектом. При его перемещении вся схема вместе с размерными линиями, стрелками, подписями и аннотациями перемещается целиком.

Координаты `Shapes` рассчитывать относительно внутренней координатной области `Chart` через единый преобразователь:

```text
model X/Y -> chart plot area X/Y
```

Этот преобразователь должен учитывать:

- сохраненные `Bounds`;
- фактические размеры `Chart.PlotArea`;
- margins;
- физически равный масштаб X/Y.

Не привязывать generated shapes к абсолютным координатам листа.

## Жизненный Цикл Shapes

При обновлении или перерисовке конкретной схемы удалять только те `Shapes`, которые ранее автоматически создал NDMProgramm для этого конкретного `Chart`.

Не удалять:

- пользовательские `Shapes`;
- объекты других схем;
- объекты на листе вне данного `Chart`;
- встроенные серии диаграммы, если они переиспользуются.

Каждому экземпляру схемы нужен устойчивый идентификатор. Рекомендуемое простое решение:

```text
Shape.Name = "NDMPlot_<PlotInstanceID>_<Role>_<Index>"
```

где:

- `PlotInstanceID` связан с конкретным `ChartObject`;
- `Role` - `DimLine`, `DimText`, `RebarLabel`, `NeutralAxis`, `Centroid`, `LoadPoint`, `Legend` и т.п.;
- `Index` - порядковый номер внутри роли.

При перерисовке `CSectionPlotter` удаляет только `Chart.Shapes`, имя которых начинается с префикса текущего `PlotInstanceID`, затем создает актуальный набор заново.

Это решение предпочтительнее хранения списка объектов в памяти, потому что схема должна восстанавливаться между запусками Excel/VBA и не зависеть от состояния объектов в памяти.

## Размерные Shapes

Размерные линии выполнять `Shape`-линиями внутри `Chart.Shapes`.

Требования:

- наконечники-стрелки с обеих сторон через `BeginArrowheadStyle` и `EndArrowheadStyle`;
- для горизонтальной размерной линии текст размера горизонтальный и расположен над линией;
- для вертикальной размерной линии текст размера вертикальный и расположен слева от линии;
- положение текста пересчитывается при каждом обновлении схемы через тот же преобразователь координат `model -> chart`, поэтому сохраняется при масштабировании и перестроении.

Нормальная реализация - именно стрелки. Короткие засечки допускаются только как технический fallback после проверки, если стрелочные наконечники невозможно надежно реализовать внутри `Chart.Shapes`. При fallback сохраняется поведение: размерная линия остается внутри `Chart`, перемещается вместе со схемой и однозначно принадлежит текущему экземпляру plot-а.

Бетон на первом этапе показывать маркерами:

- квадрат;
- ромб;
- другой компактный marker, если выглядит аккуратнее.

Арматуру показывать маркерами с размером, зависящим от `Diameter`.

Не создавать отдельный `Shape` на каждое бетонное волокно, если достаточно точек диаграммы. Это важно для производительности.

## Цвета И Состояния

Раскраска должна основываться на поле `PhysicalState`, сохраненном в расчетном snapshot и прочитанном `CSectionPlotDataReader`, а не на знаке координаты относительно нейтральной оси и не на пользовательской `OUTPUT` sign convention.

`CSectionPlotter` не анализирует знак `Strain` или `Stress`. Он только отображает уже классифицированное состояние:

```text
Compression
Tension
InactiveTensionConcrete
NearZero
```

### Неработающий Растянутый Бетон

Отдельное состояние:

```text
бетон находится в растяжении, но активная расчетная модель материала не учитывает его растягивающее напряжение
```

Пример: для режима `CrackedNDS` растянутый бетон фиксированно отключен: растягивающая деформация есть, но `Stress = 0`.

Такие бетонные элементы показывать светло-серым цветом. Они не должны попадать в обычный gradient сжатого бетона.

Если активная модель материала действительно учитывает растягивающее напряжение бетона, цвет определяется его фактическим `Stress`.

В легенде явно показать:

- сжатый бетон;
- растянутый бетон, если он работает по диаграмме;
- неработающий растянутый бетон, светло-серый;
- арматура в растяжении;
- арматура в сжатии;
- нейтральное или почти нулевое состояние, если оно требуется визуально.

## Spatial Result Labels

Не использовать подход:

```text
подписывать каждый N-й ElementID
```

Порядок `ElementID` не гарантирует пространственно равномерного распределения подписей, особенно для `LShape` и произвольной AutoCAD-геометрии.

Настройки:

```text
Plot.ResultType
Plot.ResultLabelsEnabled
Plot.ResultLabelSpacing
```

`Plot.ResultLabelSpacing` задает целевой пространственный шаг между подписями в выбранных output-единицах длины.

Подписи строятся по выбранному `Plot.ResultType`:

```text
Stress -> подписывать Stress
Strain -> подписывать Strain
```

Рекомендуемый алгоритм:

1. Взять `Bounds` расчетной модели.
2. Сформировать регулярную сетку точек-кандидатов с шагом `Plot.ResultLabelSpacing`.
3. Для каждой точки-кандидата найти ближайшее бетонное волокно.
4. Проверить, что кандидат действительно попадает в область фактического бетона:
   - для прямоугольных эквивалентных элементов - по `X/Y/Width/Height/Rotation`;
   - для круглых эквивалентных элементов - по `X/Y/Diameter`;
   - если формы нет, использовать допустимый радиус поиска от центра элемента и не выводить подпись при сомнении.
5. Не создавать подписи вне фактического бетона.
6. Соблюдать минимальное расстояние между уже размещенными labels.
7. Учитывать примерный bounding box текста, чтобы уменьшить наложения.
8. Если несколько кандидатов попали в один элемент, оставить одну подпись.
9. Подписывать только бетонные элементы.

Это пространственный алгоритм. Он не зависит от порядка строк в `Results` и работает одинаково для Generated и AutoCAD-like данных.

## Нейтральная Линия

Нейтральная линия строится для выбранного `Plot.LoadCase` из сохраненных:

```text
Epsilon0
KappaX
KappaY
```

Уравнение:

```text
kappaY * x + kappaX * y + epsilon0 = 0
```

Если линия находится далеко за пределами сечения, она все равно может быть показана как линия пересечения с областью видимого окна или как вынесенная линия в пределах plot area. Но она не должна расширять масштаб схемы.

## Физически Равный Масштаб X/Y

Одного равенства диапазонов X/Y недостаточно.

Нужно обеспечить одинаковый физический масштаб координат на экране с учетом реальных размеров `Chart.PlotArea`.

Должно выполняться:

```text
(Xmax - Xmin) / PlotAreaWidth = (Ymax - Ymin) / PlotAreaHeight
```

с учетом margins.

Алгоритм:

1. Получить исходные `Bounds` расчетной модели.
2. Добавить расчетный margin вокруг сечения.
3. Получить фактические размеры `Chart.PlotArea.InsideWidth` и `Chart.PlotArea.InsideHeight`.
4. Посчитать:

```text
scaleX = (xMax - xMin) / plotAreaWidth
scaleY = (yMax - yMin) / plotAreaHeight
scale = Max(scaleX, scaleY)
```

5. Расширить меньший диапазон вокруг центра `Bounds`, чтобы:

```text
xRange = scale * plotAreaWidth
yRange = scale * plotAreaHeight
```

6. Установить `MinimumScale`/`MaximumScale` осей диаграммы.

Удаленная нейтральная линия не участвует в расчете `Bounds` и не расширяет масштаб.

Это нужно, чтобы:

- круг отображался кругом;
- геометрические пропорции не искажались;
- расположение арматуры визуально соответствовало расчетной геометрии.

## Ответственность Классов

Минимизировать количество новых классов, но не смешивать ответственности.

Предпочтительный минимум:

```text
CSectionPlotDataReader
CSectionPlotter
```

`CSectionPlotDataReader`:

- читает `rngNDMSectionGeometry`;
- читает `rngNDMElementResults`;
- читает `rngNDMSectionProperties`;
- читает `rngNDMSectionAnnotations`;
- выбирает нужный `LoadCase`;
- объединяет геометрию и LC-зависимые результаты по `ElementID`;
- выбирает числовую величину по `Plot.ResultType`;
- восстанавливает нейтральную линию из `Epsilon0/KappaX/KappaY`;
- передает `ResultValue`, `ResultUnit`, `ResultCaption` и `PhysicalState`;
- готовит структуры для plotter-а.

`CSectionPlotter`:

- принимает уже подготовленные данные;
- рисует chart series;
- рисует shapes;
- не знает источник геометрии;
- не знает, выбран `Stress` или `Strain`, а работает с универсальным `ResultValue`;
- не обращается к расчетному ядру;
- не читает листы напрямую, если данные уже переданы reader-ом.

Если при реализации понадобится объект для annotations:

```text
CSectionAnnotations
```

Это предпочтительнее старого имени `CSectionPlotMetadata`, потому что данные сохраняются на `Results` как часть расчетного снимка оформления.

`CPlotDimension` можно оставить как имя, потому что оно точно описывает визуальный размер.

`CPlotRebarLabel` можно оставить, если класс действительно описывает готовую визуальную подпись. Если он будет хранить более широкую семантику арматуры, лучше `CRebarAnnotation`.

## Взаимодействие С Существующими Writer-Ами

После расчета writer-слой должен блочно записывать:

```text
CBatchResultWriter / CNDMResultsWriter
  -> rngNDMSectionGeometry
  -> rngNDMElementResults
  -> rngNDMSectionProperties
  -> rngNDMSectionAnnotations
```

Построчная запись запрещена.

`rngNDMSectionAnnotations` формируется во время расчетного запуска из фактически созданной геометрии/арматуры. После этого он становится частью последнего расчетного снимка.

Кнопка `Обновить схему` не пересоздает `rngNDMSectionAnnotations`.

## Проверка Текущего Output Adapter

Фактическая архитектура кода уже в целом соответствует месту преобразования `INTERNAL -> OUTPUT`:

- `CUnitSystem` читает `rngUnitSettings` и `rngSignConventionSettings`;
- `CUnitSystem` содержит методы `InternalLengthToOutput`, `InternalAreaToOutput`, `InternalForceToOutput`, `InternalMomentMxToOutput`, `InternalMomentMyToOutput`, `InternalStressToOutput`, `InternalCurvatureToOutput`;
- `CNDMResultsWriter` и `CBatchResultWriter` принимают `CUnitSystem` и записывают пользовательские значения через `Internal...ToOutput`;
- значит правильное место `OUTPUT`-преобразования - writer/output-adapter слой при записи `Results`, а не `CSectionPlotter`.

Что нужно учесть при реализации этой архитектуры:

- `CNDMResultsWriter` уже пишет `rngNDMSectionGeometry` один раз; этот диапазон сохранить и переиспользовать как источник геометрии для plotter-а и AutoCAD export;
- `modAutoCADStressExport` уже читает геометрию из `rngNDMSectionGeometry`; при реализации не ломать этот путь;
- `CNDMResultsWriter` сейчас записывает `X/Y/Area/Diameter/Material` и `Epsilon0/KappaX/KappaY` в каждую строку `rngNDMElementResults`; по новой архитектуре постоянную геометрию надо оставить в `rngNDMSectionGeometry`, а `Epsilon0/KappaX/KappaY` перенести в `rngNDMSectionProperties` и хранить один раз на `LoadCase`;
- `CUnitSystem` сейчас имеет публичные getters для output units, но не имеет публичных getters для выбранных пользовательских sign convention; для записи `Output.SignConvention.N/Mx/My` в snapshot нужно добавить read-only getters или другой единый метод экспорта этих настроек;
- AutoCAD export сейчас читает `Results` и местами переводит output-значения обратно через `Output...ToInternal`; после ввода `rngNDMSectionProperties` он должен использовать единицы snapshot, сохраненные на `Results`, а не текущие настройки `Config`.

## Допустимые Технические Решения

При реализации можно использовать более простое, надежное или архитектурно аккуратное техническое решение, если оно полностью сохраняет утвержденное поведение и инварианты этой архитектуры.

Если такое решение существенно меняет ответственность классов, состав сохраняемых диапазонов `Results` или поток данных `snapshot -> reader -> plotter`, сначала нужно отдельно описать предлагаемое изменение и причину. Молчаливо менять утвержденную архитектуру нельзя.

## Проверка На Дублирование

После перехода на разделенные блоки ненужное дублирование устраняется так:

- `rngNDMSectionGeometry` хранит геометрию и постоянные свойства элементов один раз на snapshot.
- `rngNDMElementResults` хранит только поэлементные LC-зависимые результаты.
- `rngNDMSectionProperties` хранит общие численные свойства сечения и LC-зависимые свойства плоскости деформаций.
- `rngNDMSectionAnnotations` хранит только semantic-оформление.

Не хранить в `SectionProperties` то, что надежно вычисляется из `Bounds`, например `B/H`.

Не хранить `NeutralLine.A/B/C`, потому что они вычисляются из `Epsilon0/KappaX/KappaY`.

Не хранить в `SectionAnnotations` расчетные напряжения, деформации, координаты и площади элементов.

Не хранить в `ElementResults`:

- semantic-подписи вида `6Ø32 + 3Ø20`;
- геометрию `X/Y/Area/Diameter/Material`, если она уже есть в `rngNDMSectionGeometry`.

Допустимое частичное повторение координат возможно только если это разные сущности:

- координаты расчетных элементов в `rngNDMSectionGeometry`;
- координаты размерных линий и подписей в `SectionAnnotations`.

Это не дублирование расчетных данных, а сохранение оформления, которое нельзя надежно восстановить после изменения входных данных.

## Тесты Для Будущей Реализации

Добавить тесты:

| Тест | Что проверяет |
|---|---|
| `plot.snapshot.usesResultsOnly` | `Обновить схему` не читает текущие geometry/rebar settings |
| `plot.snapshot.stableAfterInputChange` | После изменения входных данных без нового расчета схема показывает старый snapshot |
| `plot.ranges.exist` | Есть `rngNDMSectionGeometry`, `rngNDMElementResults`, `rngNDMSectionProperties`, `rngNDMSectionAnnotations` |
| `plot.sectionProperties.noRunMetadata` | Нет нового `rngNDMRunMetadata`; используется `rngNDMSectionProperties` |
| `plot.sectionProperties.noNeutralABC` | `NeutralLine.A/B/C` не записываются при наличии `Epsilon0/KappaX/KappaY` |
| `plot.sectionProperties.boundsUniversal` | `Bounds` записываются для Generated и AutoCAD-like данных |
| `plot.annotations.persisted` | Размеры и групповые подписи сохраняются в `rngNDMSectionAnnotations` |
| `plot.annotations.emptyAutoCADOk` | Пустые annotations для AutoCAD не вызывают ошибку |
| `plot.loadCase.switchNoRebuild` | Переключение LC не пересоздает геометрию и annotations |
| `plot.resultType.stress` | `Plot.ResultType = Stress` строит gradient, labels и legend по `Stress` |
| `plot.resultType.strain` | `Plot.ResultType = Strain` строит gradient, labels и legend по `Strain` |
| `autocad.resultType.stressStrain` | `AutoCAD.Export.ResultType` выбирает `Stress` или `Strain` без отдельной реализации экспортера |
| `plot.reader.joinsGeometryAndResults` | Reader объединяет `rngNDMSectionGeometry` и `rngNDMElementResults` по `ElementID` |
| `plot.resultLabels.spatial` | Подписи выбранного `Plot.ResultType` распределяются пространственно, а не по номеру элемента |
| `plot.inactiveTensionConcreteColor` | Неработающий растянутый бетон показывается светло-серым |
| `plot.physicalScaleEqual` | Физический масштаб X/Y одинаков с учетом `Chart.PlotArea` |
| `plot.neutralLine.noScaleImpact` | Далекая нейтральная линия не расширяет масштаб |
| `plot.rebarLabels.grouped` | Подписи арматуры групповые, например `6Ø32 + 3Ø20` |
| `plot.elementResults.noGeometryDuplication` | `X/Y/Area/Diameter/Material` не дублируются в `rngNDMElementResults` по каждому LC |
| `plot.elementResults.noPlaneDuplication` | `Epsilon0/KappaX/KappaY` хранятся только в `rngNDMSectionProperties`, не в каждой строке `rngNDMElementResults` |
| `plot.physicalState.snapshotOwned` | `PhysicalState` сохранен в `rngNDMElementResults`; `CSectionPlotter` не проверяет знаки `Stress/Strain` |
| `plot.outputSettingsAfterRunIgnored` | Изменение текущих `OUTPUT units/sign convention` после расчета не меняет `Results`, подписи схемы, физическую классификацию и цвет до нового расчета |
| `plot.shapes.insideChart` | Автоматические shapes создаются в `Chart.Shapes`, а не на листе |
| `plot.shapes.lifecycleScoped` | При обновлении удаляются только shapes текущего экземпляра NDMPlot по префиксу имени |
| `plot.dimensionArrows` | Размерные линии имеют стрелки с обеих сторон; текст ориентирован по типу размера |

## Открытые Решения Перед Реализацией

Перед кодированием нужно уточнить:

1. Точные позиции на листе `Results` зафиксированы под верхней сводкой, подробным блоком прочности, подробным блоком трещин и подробным блоком устойчивости: `rngStrengthSummaryAnchor = Results!A38`, `rngCrackSummaryAnchor = Results!A64`, `rngStabilitySummaryAnchor = Results!A91`, `rngNDMElementResults = Results!A115`, `rngNDMSectionGeometry = Results!L115`, `rngNDMSectionProperties = Results!AC115`, `rngNDMMaterialDiagrams = Results!AK115`, `rngNDMSectionAnnotations = Results!AX115`. Каждый следующий якорь расположен через две пустые строки по вертикали или два пустых столбца по горизонтали, поэтому блоки могут расти независимо друг от друга.
2. Итоговый формат `rngNDMSectionAnnotations` принят как широкая таблица строк `DIMENSION` и `REBAR_ANNOTATION`.
3. Нужен ли `GeometrySource` как диагностическое свойство или его не записывать.
4. Нужны ли `ConcreteCount/RebarCount` для контроля целостности или их вычислять при чтении.
5. Цвет светло-серого состояния и допуск “почти нулевого” напряжения.
6. Какой fallback использовать, если `Plot.LoadCase` ссылается на LC, которого нет в последнем snapshot.

## Рекомендуемый Первый Этап Реализации

1. Добавить `rngNDMSectionProperties` на `Results`.
2. Добавить `rngNDMSectionAnnotations` на `Results`.
3. Перестать использовать архитектурное имя `rngNDMRunMetadata`.
4. Сохранить существующий `rngNDMSectionGeometry` как источник геометрии и не ломать текущий AutoCAD export.
5. Записывать `Bounds`, центр тяжести, главные оси, output units и LC-зависимые `Epsilon0/KappaX/KappaY` в `rngNDMSectionProperties`.
6. Не записывать `NeutralLine.A/B/C`.
7. Не записывать `Epsilon0/KappaX/KappaY` в `rngNDMElementResults`.
8. Не дублировать `X/Y/Area/Diameter/Material` в `rngNDMElementResults`, если они есть в `rngNDMSectionGeometry`.
9. Сохранять `PhysicalState` в `rngNDMElementResults` при расчетном запуске; `CSectionPlotDataReader` читает его из snapshot и передает в `CSectionPlotter`.
10. Сохранять `CPlotDimension`/`CPlotRebarLabel`-эквивалентные данные в `rngNDMSectionAnnotations`.
11. Добавить настройки `[Excel plot]` внутри `rngSystemSettings`, включая `Plot.ResultType`.
12. Добавить `AutoCAD.Export.ResultType`.
13. Создавать ChartObject `chtNDMSectionPlot` около столбца `AP`; именованный диапазон `rngSectionPlot` не использовать.
14. Реализовать `CSectionPlotDataReader`.
15. Реализовать `CSectionPlotter` на `XY Scatter + Chart.Shapes`.
16. Реализовать жизненный цикл generated shapes по префиксу конкретного `Chart`.
17. Реализовать кнопку `Обновить схему`.
18. Перевести AutoCAD export на чтение нужных данных из сохраненного snapshot, где это применимо.

## Итоговое Разделение

```text
rngNDMSectionGeometry
  = фактическая геометрия и арматура, записанные один раз для snapshot.
    Источник X/Y/Area/Diameter/Material/GeometryInterpretationStatus для plotter-а и AutoCAD export.

rngNDMElementResults
  = LC-зависимые результаты элементов.
    Не хранит Epsilon0/KappaX/KappaY и не дублирует геометрию из rngNDMSectionGeometry.

rngNDMSectionProperties
  = общие численные свойства рассчитанного сечения и LC-зависимые численные свойства расчетного состояния.
    Хранит Epsilon0/KappaX/KappaY один раз на LoadCase.

rngNDMSectionAnnotations
  = сохраненная семантика оформления рассчитанной геометрии:
    какие характерные размеры показать,
    где их показать,
    какие группы арматуры подписать,
    как они относятся к смысловым граням.

CSectionPlotDataReader
  = читает последний расчетный OUTPUT-снимок с Results,
    использует сохраненные в rngNDMSectionProperties единицы/знаки snapshot для подписей,
    читает/передает сохраненное физическое состояние элемента и готовит данные.

CSectionPlotter
  = универсальный отрисовщик.
    Не знает источник геометрии, не знает решатели, не знает генераторы,
    не интерпретирует знаки Stress/Strain.
```
