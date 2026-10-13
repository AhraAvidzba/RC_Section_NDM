Attribute VB_Name = "modSectionRecipeApiGuide"
'
'
' ====================================================================================================
' СОЗДАНИЕ СЕЧЕНИЯ: ГЕОМЕТРИЯ, АРМАТУРА И АННОТАЦИИ
' ====================================================================================================
'
'   Здесь собраны готовые методы для автора нового типа сечения. Описание задаёт, что нужно получить;
'   программа по нему строит бетонную сетку, размещает стержни и готовит схему.
'
'   Модуль состоит только из комментариев. Точные объявления методов приведены вместе с параметрами; код
'   открывается в соответствующем классе через окно проекта Ctrl+R. Поиск по имени — Ctrl+F.
'
' СОДЕРЖАНИЕ
'     01  Начало и завершение описания
'     02  Входные размеры и единицы
'     03  Границы бетона и отверстий
'     04  Линии центров стержней и преобразования
'     05  Размещение арматуры
'     06  Условия размещения стержней
'     07  Размеры, пояснения и подписи арматуры
'     08  Составные линии и отдельные участки
'     09  Адреса ввода и сообщения об ошибках
'
'
' ====================================================================================================
' КАК ЧИТАТЬ ОПИСАНИЕ
' ====================================================================================================
'
'   Все координаты, длины и диаметры внутри рецепта задаются в мм, площади — в мм^2, углы — в радианах.
'   Значения из таблиц переводит context.Units; повторно переводить уже полученные мм не нужно.
'
'   Path — геометрическая линия из отрезков и дуг. Замкнутая линия может быть границей бетона или
'   отверстия. Линия центров стержней получается, например, смещением контура на as; вычислять
'   координаты каждого стержня вручную не требуется.
'
'   Segment — один отрезок или одна дуга. PathGroup объединяет несколько независимых линий. Station в
'   имени метода означает расстояние от начала линии в мм.
'
'   ID — уникальное текстовое имя линии, ряда, привязки или подписи. Anchor — именованная точка для
'   размеров и пояснений. Rule — описание одного способа размещения стержней.
'
'   Optional означает необязательный параметр; его значение по умолчанию указано после знака =. ByVal
'   передаёт значение, ByRef позволяет методу вернуть результат через параметр. Nothing вместо options
'   создаёт обычные условия размещения.
'
'   Initialize открывает одно описание, Finish завершает его. CirclePath, LinePath, RectanglePath и
'   другие создающие команды сами сохраняют линию. Возвращённый path уже зарегистрирован; Outer, Opening
'   и армирование получают его ID. AddPath нужен только для отдельно созданной линии.
'
'   Таблицы и единицы context — независимый снимок ввода. context.Units используется для преобразования
'   чисел; перенастраивать его нельзя. Проверка и размещение стержней выполняются общим Build до
'   создания сетки.
'
'   Команды подписей задают группу, геометрическую привязку и направление. Расстояние от линии арматуры
'   до края бетона программа определяет по фактическому контуру; дополнительный отступ задаётся в
'   настройках плоттера.
'
'   МИНИМАЛЬНЫЙ ПРИМЕР
'     Тело ISectionRecipe_CreateDefinition в классе с Implements ISectionRecipe. Пример задаёт
'     прямоугольник и пять верхних стержней диаметром 20 мм:
'
'       Dim facade As CSectionDefinitionBuilder, path As CGeometryPath  ' Сборщик описания и геометрическая линия.
'       Dim options As CRebarRuleOptions  ' Переменная условий размещения арматуры.
'       Set facade = New CSectionDefinitionBuilder  ' Создаём сборщик описания сечения.
'       facade.Initialize "MySection"  ' "MySection" — имя типа сечения.
'       Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)  ' "Boundary" — имя прямоугольника; 0 — X левого нижнего угла, мм; 0 — Y левого нижнего угла, мм; 600 — ширина, мм; 1000 — высота, мм.
'       facade.Outer "OUTER", path.ID  ' "OUTER" — имя наружной границы бетона; path.ID — замкнутый контур.
'       Set path = facade.LinePath("TopAxis", 50#, 950#, 550#, 950#)  ' "TopAxis" — имя линии; 50 — X начала, мм; 950 — Y начала, мм; 550 — X конца, мм; 950 — Y конца, мм.
'       Set options = New CRebarRuleOptions  ' Создаём условия размещения арматуры.
'       options.AnnotationGroup = "Top"  ' AnnotationGroup — имя общей подписи для всех рядов этой группы.
'       facade.RebarByCount "Top.Main", path.ID, 5, 20#, options  ' "Top.Main" — имя нового ряда; path.ID — линия центров стержней; 5 — заданное число стержней; 20 — диаметр стержня, мм; options — отступы, нумерация и условия размещения.
'       ' "Top" — группа общей подписи всех её рядов; "Top.Main" — ряд, к концам которого привязана
'       ' подпись; 0 — X-компонента направления подписи; 1 — Y-компонента направления подписи.
'       facade.RebarRowLabel "Top", "Top.Main", 0#, 1#
'       facade.HorizontalDistance "DIM_B", "B", 0#, 600#, 1000#  ' "DIM_B" — имя размера; "B" — обозначение на схеме; 0 — X первой точки, мм; 600 — X второй точки, мм; 1000 — уровень исходной размерной линии Y, мм.
'       Set ISectionRecipe_CreateDefinition = facade.Finish  ' Проверяем и возвращаем описание сечения.
'
'     В CSectionTypeRegistry.Class_Initialize добавьте RegisterRecipe "MySection", New
'     CSectionRecipeMySection, Array(). Затем выполните UpdateSectionRecipeConfig. Пустой Array()
'     подходит этому примеру с постоянными размерами; для размеров из Config укажите имена входных
'     таблиц.
'
'
' ====================================================================================================
' 01  НАЧАЛО И ЗАВЕРШЕНИЕ ОПИСАНИЯ
' ====================================================================================================
'
'
'   Класс / модуль: ISectionRecipe
'
'     ----------------------------------------------------------------------------------------------
'     ISectionRecipe.CreateDefinition
'
'       Обязательный метод класса нового типа. Возвращает завершённое описание сечения. Сетку, стержни и
'       рисунок программа строит по этому описанию автоматически.
'
'       Объявление:
'         Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
'
'       Параметры:
'         ByVal context As CSectionRecipeContext
'           Объект с копиями входных таблиц и системой единиц; ячейки Excel из рецепта не читаются.
'
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Initialize
'
'       Открывает один новый фасад с TypeID и версией схемы. Повторная инициализация запрещена.
'
'       Объявление:
'         Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
'
'       Параметры:
'         ByVal typeID As String
'           Смысловое имя зарегистрированного типа; непустой текст.
'
'         Optional ByVal schemaVersion As Long = 1
'           Положительная целая версия схемы определения.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Finish
'
'       Проверяет контуры, ссылки, зависимости и декларации аннотаций; публикует неизменяемое
'       CSectionDefinition и закрывает фасад. Допустимость параметров и фактическое размещение арматуры
'       проверяет CSectionBuilder до построения сетки.
'
'       Объявление:
'         Public Function Finish() As CSectionDefinition
'
'
'   Класс / модуль: CSectionTypeRegistry
'
'     ----------------------------------------------------------------------------------------------
'     CSectionTypeRegistry.RegisterRecipe
'
'       Регистрирует один ISectionRecipe и массив имён его входных блоков; ID типа уникальный.
'
'       Объявление:
'         Public Sub RegisterRecipe(ByVal typeID As String, ByVal recipe As ISectionRecipe, ByVal
'         tableNames As Variant)
'
'       Параметры:
'         ByVal typeID As String
'           Смысловое имя зарегистрированного типа; непустой текст.
'
'         ByVal recipe As ISectionRecipe
'           Один объект с реализацией ISectionRecipe.
'
'         ByVal tableNames As Variant
'           Массив имён блоков выбранного типа; Empty-array допустим для программного рецепта.
'
'
'
'
' ====================================================================================================
' 02  ВХОДНЫЕ РАЗМЕРЫ И ЕДИНИЦЫ
' ====================================================================================================
'
'
'   Класс / модуль: CSectionRecipeContext
'
'     ----------------------------------------------------------------------------------------------
'     CSectionRecipeContext.Table
'
'       Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.
'
'       Объявление:
'         Public Function Table(ByVal tableName As String) As CSectionInputTable
'
'       Параметры:
'         ByVal tableName As String
'           Имя единственного корневого входного блока этого снимка.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionRecipeContext.Units
'
'       Закрытая копия системы единиц этого снимка ввода. Внешняя перенастройка units не меняет
'       контекст; методы преобразования доступны, InitializeDefaults/LoadFromSettings запрещены.
'
'       Объявление:
'         Public Property Get Units() As CUnitSystem
'
'
'   Класс / модуль: CSectionInputTable
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Name
'
'       Имя входной таблицы.
'
'       Объявление:
'         Public Property Get Name() As String
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowCount
'
'       Число строк в копии входной таблицы.
'
'       Объявление:
'         Public Property Get RowCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.ColumnCount
'
'       Число столбцов в копии входной таблицы.
'
'       Объявление:
'         Public Property Get ColumnCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.SourceReference
'
'       Возвращает исходный адрес ячейки или программный tableName[row,col].
'
'       Объявление:
'         Public Function SourceReference(ByVal row As Long, ByVal column As Long) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.CellValue
'
'       Возвращает исходное значение из копии таблицы. При ошибке Excel сообщает адрес ячейки, которую
'       нужно исправить.
'
'       Объявление:
'         Public Function CellValue(ByVal row As Long, ByVal column As Long) As Variant
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.IsBlank
'
'       Проверяет пустое значение снимка; ошибка Excel не подменяется пустотой.
'
'       Объявление:
'         Public Function IsBlank(ByVal row As Long, ByVal column As Long) As Boolean
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Text
'
'       Читает непустой текст либо явно переданный fallback, обрезая крайние пробелы.
'
'       Объявление:
'         Public Function Text(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Variant) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Number
'
'       Читает число, исключая Boolean/ошибки/объекты; обязательное пустое поле не заменяется нулём.
'
'       Объявление:
'         Public Function Number(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Variant) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Variant
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Count
'
'       Читает целое неотрицательное 0..2147483647; отдельный executor проверяет суммарное число
'       позиций.
'
'       Объявление:
'         Public Function Count(ByVal row As Long, ByVal column As Long, Optional ByVal defaultValue As
'         Long = 0) As Long
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         Optional ByVal defaultValue As Long = 0
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Choice
'
'       Проверяет вариант по allowed без учёта регистра и возвращает каноническое значение allowed.
'
'       Объявление:
'         Public Function Choice(ByVal row As Long, ByVal column As Long, ByVal allowed As Variant,
'         Optional ByVal defaultValue As String = vbNullString) As String
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal allowed As Variant
'           Массив разрешённых строк; Choice возвращает каноническую строку из этого массива.
'
'         Optional ByVal defaultValue As String = vbNullString
'           Значение для пустой ячейки. В Text и Number без этого параметра пустая ячейка считается
'           ошибкой; остальные методы используют указанное в объявлении значение по умолчанию.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.Length
'
'       Читает обязательное число и один раз переводит input length в мм через CUnitSystem.
'
'       Объявление:
'         Public Function Length(ByVal row As Long, ByVal column As Long, ByVal units As CUnitSystem) As
'         Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.OptionalDiameter
'
'       Пустая ячейка или 0 отключают ряд; положительный диаметр переводится в мм. Текст, включая
'       одиночный дефис и TODO, и отрицательный диаметр вызывают адресную ошибку ввода.
'
'       Объявление:
'         Public Function OptionalDiameter(ByVal row As Long, ByVal column As Long, ByVal units As
'         CUnitSystem) As Double
'
'       Параметры:
'         ByVal row As Long
'           Номер строки снимка 1..RowCount.
'
'         ByVal column As Long
'           Номер столбца снимка 1..ColumnCount.
'
'         ByVal units As CUnitSystem
'           CUnitSystem, уже подготовленная вызывающим workflow; преобразование только через её методы.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.ColumnByCaption
'
'       Находит уникальный столбец заголовка, без фиксированного адреса.
'
'       Объявление:
'         Public Function ColumnByCaption(ByVal headerRow As Long, ByVal caption As String) As Long
'
'       Параметры:
'         ByVal headerRow As Long
'           Строка заголовков, в пределах снимка.
'
'         ByVal caption As String
'           Смысловой заголовок столбца; повторные заголовки — ошибка.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowByKey
'
'       Находит уникальную строку смыслового ключа в указанном подблоке.
'
'       Объявление:
'         Public Function RowByKey(ByVal key As String, Optional ByVal keyColumn As Long = 1, Optional
'         ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
'
'       Параметры:
'         ByVal key As String
'           Смысловой ключ строки, не Excel-адрес.
'
'         Optional ByVal keyColumn As Long = 1
'           Столбец смысловых ключей 1..ColumnCount.
'
'         Optional ByVal firstRow As Long = 1
'           Начало диапазона поиска строки, включительно.
'
'         Optional ByVal lastRow As Long = 0
'           Конец поиска, включительно; 0 означает RowCount.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionInputTable.RowByAnyKey
'
'       Находит уникальную строку по массиву допустимых альтернативных ключей; неоднозначность
'       запрещена.
'
'       Объявление:
'         Public Function RowByAnyKey(ByVal keys As Variant, Optional ByVal keyColumn As Long = 1,
'         Optional ByVal firstRow As Long = 1, Optional ByVal lastRow As Long = 0) As Long
'
'       Параметры:
'         ByVal keys As Variant
'           Массив допустимых альтернативных ключей; неоднозначное совпадение отклоняется.
'
'         Optional ByVal keyColumn As Long = 1
'           Столбец смысловых ключей 1..ColumnCount.
'
'         Optional ByVal firstRow As Long = 1
'           Начало диапазона поиска строки, включительно.
'
'         Optional ByVal lastRow As Long = 0
'           Конец поиска, включительно; 0 означает RowCount.
'
'
'
'   Класс / модуль: CUnitSystem
'
'     ----------------------------------------------------------------------------------------------
'     CUnitSystem.InputLengthToInternal
'
'       Переводит InputLengthToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
'       принятой конвенцией знаков для N/Mx/My.
'
'       Объявление:
'         Public Function InputLengthToInternal(ByVal value As Double) As Double
'
'       Параметры:
'         ByVal value As Double
'           Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный
'           перевод.
'
'
'     ----------------------------------------------------------------------------------------------
'     CUnitSystem.InputAreaToInternal
'
'       Переводит InputAreaToInternal во внутренние мм/Н/Н*мм/МПа/1/мм соответствующей размерности, с
'       принятой конвенцией знаков для N/Mx/My.
'
'       Объявление:
'         Public Function InputAreaToInternal(ByVal value As Double) As Double
'
'       Параметры:
'         ByVal value As Double
'           Входное/выходное скалярное значение в размерности операции; CUnitSystem задаёт единственный
'           перевод.
'
'
'
'
' ====================================================================================================
' 03  ГРАНИЦЫ БЕТОНА И ОТВЕРСТИЙ
' ====================================================================================================
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.CirclePath
'
'       Создаёт замкнутую окружность CCW от точки +X; один сегмент с тем же ID, что путь.
'
'       Объявление:
'         Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As
'         Double, ByVal radius As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.LinePath
'
'       Создаёт непрерывный открытый путь из одной прямой от первого конца ко второму.
'
'       Объявление:
'         Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal
'         x2 As Double, ByVal y2 As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal x1 As Double
'           Модельная координата x1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y1 As Double
'           Модельная координата y1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal x2 As Double
'           Модельная координата x2, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y2 As Double
'           Модельная координата y2, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.ArcPath
'
'       Создаёт точную дугу по центру/радиусу/startAngle/signed sweep; полный оборот замкнут.
'
'       Объявление:
'         Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As
'         Double, ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'         ByVal startAngle As Double
'           Начальный угол, радианы; от +X.
'
'         ByVal sweep As Double
'           Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, |sweep|<=2*pi.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RectanglePath
'
'       Создаёт замкнутую наружную границу одного прямоугольника от (x,y) с положительными width/height.
'
'       Объявление:
'         Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double,
'         ByVal width As Double, ByVal height As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'         ByVal width As Double
'           Координата/размер width, мм; размеры width/height положительны.
'
'         ByVal height As Double
'           Координата/размер height, мм; размеры width/height положительны.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RoundedPolygon
'
'       Создаёт согласованную границу по вершинам/радиусам/ID. Edge.ID начинается у указанной вершины,
'       Corner.ID существует только при R>0.
'
'       Объявление:
'         Public Function RoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal radii
'         As Variant, ByVal vertexIDs As Variant) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal vertices As Variant
'           Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
'
'         ByVal radii As Variant
'           Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса
'           отсутствует.
'
'         ByVal vertexIDs As Variant
'           Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.AdjacentRectangles
'
'       Создаёт одну границу смежных неперекрывающихся прямоугольников. Произвольное пересекающееся CAD
'       Union не выполняется.
'
'       Объявление:
'         Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal rectangles As Variant
'           Array(Array(x,y,width,height),...); мм, размеры положительные, прямоугольники смежные без
'           перекрытий.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.AddPath
'
'       Регистрирует непрерывную линию, созданную вручную или отдельным общим инструментом. Создающие
'       команды фасада уже регистрируют свой результат; повторный AddPath для него является ошибкой.
'       Назначение Outer/Opening выполняется отдельно.
'
'       Объявление:
'         Public Sub AddPath(ByVal path As CGeometryPath)
'
'       Параметры:
'         ByVal path As CGeometryPath
'           Существующий непрерывный опубликованный CGeometryPath.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Outer
'
'       Назначает замкнутому пути роль наружной границы материала. Не требует наличия арматуры.
'
'       Объявление:
'         Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As
'         String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "")
'
'       Параметры:
'         ByVal loopID As String
'           Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal contourIDs As Variant
'           Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути.
'
'         Optional ByVal arcComment As String = ""
'           Отдельный комментарий аналитических дуг; пустой использует обычный comment.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Opening
'
'       Назначает замкнутому пути роль отверстия с конкретным parentOuterID. Действительное вложение
'       проверяется при Finish.
'
'       Объявление:
'         Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As
'         String, Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional
'         ByVal arcComment As String = "")
'
'       Параметры:
'         ByVal loopID As String
'           Непустой устойчивый ID контура; у CConcreteRegion численный индекс loop 1..LoopCount.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal parentOuterID As String
'           ID ранее объявленной наружной границы для Opening; проверяется реальное вложение.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal contourIDs As Variant
'           Необязательный массив ID сегментов сохранённого контура, по одному на каждый сегмент пути.
'
'         Optional ByVal arcComment As String = ""
'           Отдельный комментарий аналитических дуг; пустой использует обычный comment.
'
'
'
'
' ====================================================================================================
' 04  ЛИНИИ ЦЕНТРОВ СТЕРЖНЕЙ И ПРЕОБРАЗОВАНИЯ
' ====================================================================================================
'
'   Сначала задайте контур, затем объявите смещение на нужное as. Для второго и третьего независимого
'   ряда задайте свои линии и свои as; связь с первым рядом не обязательна.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Offset
'
'       Смещает линию/аналитическую цепочку слева от обхода при положительном signedLeftDistance.
'       Line/line стыки miter; arc стыки должны быть гладкими.
'
'       Объявление:
'         Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         signedLeftDistance As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal signedLeftDistance As Double
'           Поперечное смещение, мм; положительно слева от направления пути, отрицательно справа.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.OffsetRoundedPolygon
'
'       Смещает весь скруглённый многоугольник с согласованными стыками. Положительное расстояние внутрь
'       площади; отрицательное наружу, из Opening в бетон.
'
'       Объявление:
'         Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal vertices As Variant, ByVal
'         radii As Variant, ByVal vertexIDs As Variant, ByVal distance As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal vertices As Variant
'           Двумерные вершины: Array(Array(x,y),...); координаты во внутренних мм.
'
'         ByVal radii As Variant
'           Массив неотрицательных радиусов вершин в мм, длина равна vertices; clamp бетонного радиуса
'           отсутствует.
'
'         ByVal vertexIDs As Variant
'           Массив уникальных смысловых ID вершин; определяет Edge.ID/Corner.ID.
'
'         ByVal distance As Double
'           Signed offset, мм: положительно внутрь площади, отрицательно наружу; это не left-offset
'           обхода.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.ConstrainLine
'
'       Ограничивает или продлевает одну прямую двумя поперечными плоскостями nX*x+nY*y=bound;
'       направление исходного пути сохраняется.
'
'       Объявление:
'         Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         startNX As Double, ByVal startNY As Double, ByVal startBound As Double, ByVal endNX As Double,
'         ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal startNX As Double
'           X-компонента начальной поперечной нормали; линия не должна быть параллельна плоскости.
'
'         ByVal startNY As Double
'           Y-компонента начальной поперечной нормали.
'
'         ByVal startBound As Double
'           Значение начальной плоскости startNX*x+startNY*y=startBound; мм при единичной нормали.
'
'         ByVal endNX As Double
'           X-компонента конечной поперечной нормали.
'
'         ByVal endNY As Double
'           Y-компонента конечной поперечной нормали.
'
'         ByVal endBound As Double
'           Значение конечной плоскости endNX*x+endNY*y=endBound; мм при единичной нормали.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Chain
'
'       Выбирает непрерывный участок между двумя именованными сегментными привязками. Направление обхода
'       явно; fullLoopIfSame задаёт полный оборот при совпавших концах.
'
'       Объявление:
'         Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, ByVal
'         startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal
'         endParameter As Double, Optional ByVal forward As Boolean = True, Optional ByVal
'         fullLoopIfSame As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal startSegmentID As String
'           Сегмент начальной привязки цепочки.
'
'         ByVal startParameter As Double
'           Параметр начальной точки выбранного сегмента [0;1].
'
'         ByVal endSegmentID As String
'           Сегмент конечной привязки цепочки.
'
'         ByVal endParameter As Double
'           Параметр конечной точки выбранного сегмента [0;1].
'
'         Optional ByVal forward As Boolean = True
'           Явное направление следования исходному обходу; False — противоположное.
'
'         Optional ByVal fullLoopIfSame As Boolean = False
'           Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Transform
'
'       Создаёт новый путь из существующего изометрией; сохраняет сегментные ID и происхождение,
'       отражение меняет знак sweep.
'
'       Объявление:
'         Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX
'         As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional ByVal
'         reflectX As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal sourcePathID As String
'           ID ранее объявленного пути, который читается без изменения.
'
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathGroup
'
'       Объявляет массив ID независимых путей; порядок значим, фиктивные соединения отсутствуют.
'
'       Объявление:
'         Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
'
'       Параметры:
'         ByVal groupID As String
'           Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
'
'         ByVal pathIDs As Variant
'           Массив ID существующих независимых путей; без повторов.
'
'
'
'
' ====================================================================================================
' 05  РАЗМЕЩЕНИЕ АРМАТУРЫ
' ====================================================================================================
'
'   Нулевой диаметр или нулевое число стержней отключают ряд. Пустой диаметр из таблицы читает
'   OptionalDiameter как 0; одиночный дефис недопустим. Для каждого независимого ряда используется свой
'   ID.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByCount
'
'       Задаёт места для указанного числа стержней на рабочей части линии размещения. EndMode определяет
'       включение концов. В группе это число применяется отдельно к каждой линии; нулевое число стержней
'       или нулевой диаметр отключают ряд.
'
'       Объявление:
'         Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long,
'         ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing, Optional
'         ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal count As Long
'           Число стержней на одной линии размещения; целое неотрицательное. В группе применяется
'           отдельно к каждой линии. Ноль отключает ряд.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByMaxStep
'
'       Задаёт места стержней так, чтобы расстояние между соседними стержнями вдоль линии не превышало
'       step. Исполнитель подбирает число стержней; фактическое расстояние может быть меньше заданного.
'
'       Объявление:
'         Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As
'         Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal step As Double
'           Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarByFixedStep
'
'       Задаёт точное расстояние step между соседними стержнями вдоль линии. RemainderPlacement
'       определяет, где оставить неиспользованный остаток длины. При IncludeBoth или ClosedUniform
'       рабочая длина должна соответствовать заданному расстоянию.
'
'       Объявление:
'         Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As
'         Double, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal step As Double
'           Расстояние между соседними стержнями вдоль линии размещения, мм; строго положительное.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarAtStations
'
'       Задаёт отдельные места стержней расстояниями от начала линии размещения, в мм. Порядок этих
'       расстояний в массиве определяет исходную нумерацию стержней.
'
'       Объявление:
'         Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As
'         Variant, ByVal diameter As Double, Optional ByVal options As CRebarRuleOptions = Nothing,
'         Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal stations As Variant
'           Массив расстояний от начала линии размещения до каждого стержня, мм; порядок элементов
'           задаёт исходную нумерацию.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         Optional ByVal options As CRebarRuleOptions = Nothing
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarDependent
'
'       Задаёт дополнительный ряд смещением относительно центров стержней исходного ряда. TangentOffset
'       задаёт смещение вдоль касательной, NormalOffset — поперёк линии. Число и места стержней связаны
'       с исходным рядом, новое равномерное распределение не выполняется.
'
'       Объявление:
'         Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter
'         As Double, ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal sourceRuleID As String
'           ID исходного правила зависимости; циклы запрещены.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         ByVal options As CRebarRuleOptions
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarProjected
'
'       Задаёт новый ряд проекцией центров стержней исходного ряда на целевую линию target. Направление
'       — нормаль исходной линии или общий вектор DirectionX/Y. StartSetback и EndSetback ограничивают
'       рабочий участок целевой линии; отсутствие пересечения отклоняется либо пропускается по
'       SkipMissingProjection.
'
'       Объявление:
'         Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID
'         As String, ByVal diameter As Double, ByVal options As CRebarRuleOptions, Optional ByVal
'         steelClass As String = "Rebar")
'
'       Параметры:
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal sourceRuleID As String
'           ID исходного правила зависимости; циклы запрещены.
'
'         ByVal diameter As Double
'           Диаметр стержня, мм; отрицательный запрещён, ноль выключает ряд.
'
'         ByVal options As CRebarRuleOptions
'           CRebarRuleOptions; поля копируются при объявлении. Nothing создаёт условия размещения по
'           умолчанию.
'
'         Optional ByVal steelClass As String = "Rebar"
'           Строка принадлежности материала арматуры; совпадения разных материалов требуют явной
'           политики.
'
'
'
'
' ====================================================================================================
' 06  УСЛОВИЯ РАЗМЕЩЕНИЯ СТЕРЖНЕЙ
' ====================================================================================================
'
'   Создайте New CRebarRuleOptions и измените только нужные свойства. Метод размещения копирует их при
'   объявлении ряда, поэтому позднейшая правка options не меняет уже объявленные ряды.
'
'
'   Класс / модуль: CRebarRuleOptions
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.StartSetback
'
'       Начальное отсечение вдоль каждого target-пути, мм>=0; у полного ClosedUniform 0.
'
'       Объявление:
'         Public StartSetback As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EndSetback
'
'       Конечное отсечение вдоль каждого target-пути, мм>=0; ProjectFrom фильтрует hit без
'       перераспределения.
'
'       Объявление:
'         Public EndSetback As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EndMode
'
'       IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
'       замкнутости.
'
'       Объявление:
'         Public EndMode As ERebarEndMode
'       По умолчанию: IncludeBoth.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DistributionScope
'
'       Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются
'       отдельно к каждой независимой линии группы.
'
'       Объявление:
'         Public DistributionScope As ERebarDistributionScope
'       По умолчанию: PerPath.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.RemainderPlacement
'
'       RemainderStart/End/Center: положение остатка для точного шага.
'
'       Объявление:
'         Public RemainderPlacement As ERemainderPlacement
'       По умолчанию: RemainderCenter.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.TangentOffset
'
'       Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.
'
'       Объявление:
'         Public TangentOffset As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.NormalOffset
'
'       Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.
'
'       Объявление:
'         Public NormalOffset As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.EverySecond
'
'       Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal
'       задаёт отдельный отсчёт успешных проекций.
'
'       Объявление:
'         Public EverySecond As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.IncomingTangentAtJoint
'
'       В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся
'       касательная следующего сегмента.
'
'       Объявление:
'         Public IncomingTangentAtJoint As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DirectionX
'
'       Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.
'
'       Объявление:
'         Public DirectionX As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.DirectionY
'
'       Глобальная Y направления ProjectFrom; направление не нулевое.
'
'       Объявление:
'         Public DirectionY As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceNormalSign
'
'       -1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.
'
'       Объявление:
'         Public SourceNormalSign As Double
'       По умолчанию: 0.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SuccessfulProjectionOrdinal
'
'       True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд
'       по всей группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на
'       каждой линии.
'
'       Объявление:
'         Public SuccessfulProjectionOrdinal As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.NewPublicationRoot
'
'       У ProjectFrom начинает новый корень публикации по успешным позициям.
'
'       Объявление:
'         Public NewPublicationRoot As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SkipMissingProjection
'
'       Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.
'
'       Объявление:
'         Public SkipMissingProjection As Boolean
'       По умолчанию: False.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Publication
'
'       PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.
'
'       Объявление:
'         Public Publication As ERebarPublication
'       По умолчанию: PublishByRule.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.PositionOrder
'
'       Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места
'       стержней и их исходные номера сохраняются.
'
'       Объявление:
'         Public PositionOrder As ERebarPositionOrder
'       По умолчанию: AlongPath.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Coincidence
'
'       RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного
'       опубликованного владельца.
'
'       Объявление:
'         Public Coincidence As ERebarCoincidence
'       По умолчанию: RejectCoincidence.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.PlacementCheck
'
'       CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
'
'       Объявление:
'         Public PlacementCheck As ERebarPlacementCheck
'       По умолчанию: CheckCenter.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.OverlapPolicy
'
'       ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
'
'       Объявление:
'         Public OverlapPolicy As ERebarOverlapPolicy
'       По умолчанию: ReportOverlap.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourcePrefix
'
'       Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.
'
'       Объявление:
'         Public SourcePrefix As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceSuffix
'
'       Суффикс SourcePositionId, в том числе для связанных рядов.
'
'       Объявление:
'         Public SourceSuffix As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.AnnotationGroup
'
'       Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.
'
'       Объявление:
'         Public AnnotationGroup As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.Comment
'
'       Инженерное пояснение правила; не selector или ResultCode.
'
'       Объявление:
'         Public Comment As String
'       По умолчанию: пустое значение VBA.
'
'     ----------------------------------------------------------------------------------------------
'     CRebarRuleOptions.SourceReference
'
'       Реальные адреса исходных параметров правила или программная ссылка.
'
'       Объявление:
'         Public SourceReference As String
'       По умолчанию: пустое значение VBA.
'
'
'   Класс / модуль: modSectionRecipeTypes
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarEndMode
'
'       IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей
'       замкнутости.
'
'       Объявление:
'         Public Enum ERebarEndMode
'
'       Допустимые значения / состав:
'         IncludeBoth = 1
'         ExcludeBoth = 2
'         IncludeStart = 3
'         IncludeEnd = 4
'         ClosedUniform = 5
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarDistributionScope
'
'       Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются
'       отдельно к каждой независимой линии группы.
'
'       Объявление:
'         Public Enum ERebarDistributionScope
'
'       Допустимые значения / состав:
'         PerPath = 1
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERemainderPlacement
'
'       RemainderStart/End/Center: положение остатка для точного шага.
'
'       Объявление:
'         Public Enum ERemainderPlacement
'
'       Допустимые значения / состав:
'         RemainderStart = 1
'         RemainderEnd = 2
'         RemainderCenter = 3
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPublication
'
'       PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.
'
'       Объявление:
'         Public Enum ERebarPublication
'
'       Допустимые значения / состав:
'         PublishByRule = 1
'         PublishPerSourcePosition = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPositionOrder
'
'       Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. Места
'       стержней и их исходные номера сохраняются.
'
'       Объявление:
'         Public Enum ERebarPositionOrder
'
'       Допустимые значения / состав:
'         AlongPath = 1
'         FromPathMiddle = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarCoincidence
'
'       RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного
'       опубликованного владельца.
'
'       Объявление:
'         Public Enum ERebarCoincidence
'
'       Допустимые значения / состав:
'         RejectCoincidence = 1
'         MergeIdentical = 2
'         KeepPublishedOwner = 3
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarPlacementCheck
'
'       CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
'
'       Объявление:
'         Public Enum ERebarPlacementCheck
'
'       Допустимые значения / состав:
'         CheckCenter = 1
'         CheckFullDisk = 2
'
'     ----------------------------------------------------------------------------------------------
'     modSectionRecipeTypes.ERebarOverlapPolicy
'
'       ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
'
'       Объявление:
'         Public Enum ERebarOverlapPolicy
'
'       Допустимые значения / состав:
'         ReportOverlap = 1
'         RejectOverlap = 2
'
'
'
' ====================================================================================================
' 07  РАЗМЕРЫ, ПОЯСНЕНИЯ И ПОДПИСИ АРМАТУРЫ
' ====================================================================================================
'
'   Для общей подписи стороны задайте одинаковую AnnotationGroup всем её рядам. Подпись считает
'   фактически созданные стержни всей группы и исчезает, если их нет. Смещение подписи в настройках
'   плоттера отсчитывается от границы бетона.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Anchor
'
'       Объявляет именованную модельную точку. Данные копируются как два скаляра.
'
'       Объявление:
'         Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathAnchor
'
'       Получает точку конкретного сегмента по ID и параметру [0;1], объявляет именованную привязку.
'
'       Объявление:
'         Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0#)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0#
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.TransformAnchor
'
'       Создаёт новую привязку из существующей: отражение по локальной X, поворот вокруг (0,0), перенос.
'
'       Объявление:
'         Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal
'         shiftX As Double, ByVal shiftY As Double, Optional ByVal rotation As Double = 0#, Optional
'         ByVal reflectX As Boolean = False)
'
'       Параметры:
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal sourceAnchorID As String
'           ID существующей исходной привязки.
'
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.VerticalExtent
'
'       Измеряет полный вертикальный габарит выбранного пути/группы с точными экстремумами дуг.
'       placementX и normalX задают сторону отображения.
'
'       Объявление:
'         Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As
'         String, ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal placementX As Double
'           Модельная X размерной линии, мм.
'
'         Optional ByVal normalX As Double = -1#
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.HorizontalExtent
'
'       Измеряет полный горизонтальный габарит выбранного пути/группы. Это не длина пути и не исходная
'       ширина до скругления.
'
'       Объявление:
'         Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name
'         As String, ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal placementY As Double
'           Модельная Y размерной линии, мм.
'
'         Optional ByVal normalY As Double = 1#
'           Y-компонента стороны размерной/подписной линии.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.HorizontalDistance
'
'       Объявляет размер между двумя заданными X на уровне placementY. DisplayIncrement округляет только
'       подпись, не геометрию.
'
'       Объявление:
'         Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, ByVal startX
'         As Double, ByVal endX As Double, ByVal placementY As Double, Optional ByVal normalY As Double
'         = 1#, Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal startX As Double
'           Модельная координата startX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal endX As Double
'           Модельная координата endX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal placementY As Double
'           Модельная Y размерной линии, мм.
'
'         Optional ByVal normalY As Double = 1#
'           Y-компонента стороны размерной/подписной линии.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'         Optional ByVal displayIncrement As Double = 0#
'           Неотрицательный шаг округления только видимого справочного значения, мм; 0 означает
'           отсутствие такого округления.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.VerticalDistance
'
'       Объявляет вертикальный размер между двумя Y на линии placementX; значение — модуль разности.
'
'       Объявление:
'         Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, ByVal startY
'         As Double, ByVal endY As Double, ByVal placementX As Double, Optional ByVal normalX As Double
'         = -1#, Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         ByVal startY As Double
'           Модельная координата startY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal endY As Double
'           Модельная координата endY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal placementX As Double
'           Модельная X размерной линии, мм.
'
'         Optional ByVal normalX As Double = -1#
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Dimension
'
'       Измеряет между двумя именованными точками: Horizontal/Vertical проекция или Aligned расстояние.
'       Проекция размещается через первую точку.
'
'       Объявление:
'         Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal
'         endAnchorID As String, ByVal name As String, Optional ByVal measurement As String = "Aligned",
'         Optional ByVal normalX As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal
'         comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal startAnchorID As String
'           ID начальной модельной точки измерения.
'
'         ByVal endAnchorID As String
'           ID конечной модельной точки измерения.
'
'         ByVal name As String
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal measurement As String = "Aligned"
'           Ровно Horizontal, Vertical или Aligned; значение чувствительно к указанному контракту.
'
'         Optional ByVal normalX As Double = 0#
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         Optional ByVal normalY As Double = 0#
'           Y-компонента стороны размерной/подписной линии.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.PathLength
'
'       Объявляет точную суммарную длину target-путей и текстовую привязку; не создаёт фиктивную
'       соединяющую линию группы.
'
'       Объявление:
'         Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal
'         placementAnchorID As String, Optional ByVal name As String = "L", Optional ByVal comment As
'         String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal placementAnchorID As String
'           ID точки размещения текстового измерения.
'
'         Optional ByVal name As String = "L"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Radius
'
'       Измеряет радиус конкретной Arc/Circle; параметр выбирает конечную точку радиуса. Прямая —
'       ошибка.
'
'       Объявление:
'         Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R",
'         Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0.5
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'         Optional ByVal name As String = "R"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Diameter
'
'       Измеряет диаметр конкретной Arc/Circle с противоположным концом на той же окружности. Прямая —
'       ошибка.
'
'       Объявление:
'         Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As
'         String, Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D",
'         Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0.5
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'         Optional ByVal name As String = "D"
'           Обозначение измерения или имя контейнера.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Text
'
'       Объявляет простой текст на именованной модельной привязке; пиксели/шрифт задаёт общий layout.
'
'       Объявление:
'         Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As
'         String, Optional ByVal comment As String = "")
'
'       Параметры:
'         ByVal annotationID As String
'           Непустой уникальный ID аннотации.
'
'         ByVal anchorID As String
'           Непустой уникальный ID модельной привязки.
'
'         ByVal textValue As String
'           Текст подписи; Unicode сохраняется.
'
'         Optional ByVal comment As String = ""
'           Пояснение происхождения/инженерного смысла; не управляющий статус.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarLabel
'
'       Привязывает фактическую подпись группы к габариту target-пути, со стороны normalX/Y.
'
'       Объявление:
'         Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As
'         Double, ByVal normalY As Double)
'
'       Параметры:
'         ByVal groupID As String
'           Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
'
'         ByVal target As String
'           ID пути/PathGroup, разрешаемый в этом определении.
'
'         ByVal normalX As Double
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         ByVal normalY As Double
'           Y-компонента стороны размерной/подписной линии.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowLabel
'
'       Привязывает фактическую подпись группы к концам исполненного ruleID.
'
'       Объявление:
'         Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As
'         Double, ByVal normalY As Double)
'
'       Параметры:
'         ByVal groupID As String
'           Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal normalX As Double
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         ByVal normalY As Double
'           Y-компонента стороны размерной/подписной линии.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowExtentLabel
'
'       Привязывает подпись к фактическому габариту позиций правила; reverseDirection меняет ориентацию
'       anchor-линии.
'
'       Объявление:
'         Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX
'         As Double, ByVal normalY As Double, Optional ByVal reverseDirection As Boolean = False)
'
'       Параметры:
'         ByVal groupID As String
'           Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal normalX As Double
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         ByVal normalY As Double
'           Y-компонента стороны размерной/подписной линии.
'
'         Optional ByVal reverseDirection As Boolean = False
'           Обращает направление линии привязки подписи; физический ряд не меняется.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.RebarRowExtremeLabel
'
'       Выбирает крайние созданные позиции в directionX/Y; сторона текста normalX/Y независима от выбора
'       ряда.
'
'       Объявление:
'         Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal
'         directionX As Double, ByVal directionY As Double, ByVal normalX As Double, ByVal normalY As
'         Double)
'
'       Параметры:
'         ByVal groupID As String
'           Смысловой ID группы; путь и арматурная annotation-group — разные контракты.
'
'         ByVal ruleID As String
'           Непустой уникальный ID правила; при диагностике — ID ранее исполненного правила.
'
'         ByVal directionX As Double
'           X-компонента направления проекции/выбора экстремума; вектор не нулевой, у
'           RebarRowExtremeLabel единичный.
'
'         ByVal directionY As Double
'           Y-компонента направления проекции/выбора экстремума.
'
'         ByVal normalX As Double
'           X-компонента стороны размерной/подписной линии; layout нормирует вектор.
'
'         ByVal normalY As Double
'           Y-компонента стороны размерной/подписной линии.
'
'
'
'
' ====================================================================================================
' 08  СОСТАВНЫЕ ЛИНИИ И ОТДЕЛЬНЫЕ УЧАСТКИ
' ====================================================================================================
'
'   Обычно достаточно готовых методов разделов 03–04. Для собственной последовательности отрезков и дуг:
'   Initialize, AddLine/AddArc, Seal, затем facade.AddPath. После Seal менять состав линии нельзя.
'
'
'   Класс / модуль: CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Initialize
'
'       Создаёт mutable путь с непустым ID один раз.
'
'       Объявление:
'         Public Sub Initialize(ByVal pathID As String)
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AddLine
'
'       Добавляет прямой сегмент с устойчивым ID; начало совпадает с предыдущим концом.
'
'       Объявление:
'         Public Sub AddLine(ByVal segmentID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2
'         As Double, ByVal y2 As Double)
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         ByVal x1 As Double
'           Модельная координата x1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y1 As Double
'           Модельная координата y1, мм; знаки соответствуют section XY/WCS.
'
'         ByVal x2 As Double
'           Модельная координата x2, мм; знаки соответствуют section XY/WCS.
'
'         ByVal y2 As Double
'           Модельная координата y2, мм; знаки соответствуют section XY/WCS.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AddArc
'
'       Добавляет аналитическую дугу по центру/R/startAngle/sweep, без хордовой аппроксимации.
'
'       Объявление:
'         Public Sub AddArc(ByVal segmentID As String, ByVal centerX As Double, ByVal centerY As Double,
'         ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double)
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         ByVal centerX As Double
'           Модельная координата centerX, мм; знаки соответствуют section XY/WCS.
'
'         ByVal centerY As Double
'           Модельная координата centerY, мм; знаки соответствуют section XY/WCS.
'
'         ByVal radius As Double
'           Положительный радиус, мм.
'
'         ByVal startAngle As Double
'           Начальный угол, радианы; от +X.
'
'         ByVal sweep As Double
'           Signed sweep, радианы: CCW положителен, CW отрицателен, 0 запрещён, |sweep|<=2*pi.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Seal
'
'       Публикует путь. closed=True проверяет совпадение концов, не достраивает соединение.
'
'       Объявление:
'         Public Sub Seal(Optional ByVal closed As Boolean = False)
'
'       Параметры:
'         Optional ByVal closed As Boolean = False
'           Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.ID
'
'       Смысловой ID этого пути.
'
'       Объявление:
'         Public Property Get ID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentCount
'
'       Число аналитических line/arc сегментов.
'
'       Объявление:
'         Public Property Get SegmentCount() As Long
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.IsClosed
'
'       Замкнут ли фактический путь, без фиктивного соединения.
'
'       Объявление:
'         Public Property Get IsClosed() As Boolean
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Length
'
'       Точная суммарная длина пути, мм.
'
'       Объявление:
'         Public Property Get Length() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentID
'
'       Возвращает смысловой ID сегмента по индексу.
'
'       Объявление:
'         Public Property Get SegmentID(ByVal index As Long) As String
'
'       Параметры:
'         ByVal index As Long
'           1-based индекс в пределах соответствующего готового контейнера.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.SegmentIndex
'
'       Разрешает существующий смысловой ID в текущий индекс; исчезнувшая дуга даёт MISSING_SEGMENT.
'
'       Объявление:
'         Public Function SegmentIndex(ByVal segmentID As String) As Long
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AnchorStation
'
'       Находит расстояние от начала всей линии до точки заданного сегмента, мм. Точка задаётся именем
'       сегмента и параметром [0;1]. Если при построении участок исчез, обе его границы сохраняют
'       привязку к оставшемуся стыку. Неизвестное имя вызывает ошибку.
'
'       Объявление:
'         Public Function AnchorStation(ByVal segmentID As String, Optional ByVal parameter As Double =
'         0#) As Double
'
'       Параметры:
'         ByVal segmentID As String
'           Устойчивый ID существующего сегмента; исчезнувший Corner не подменяется индексом.
'
'         Optional ByVal parameter As Double = 0#
'           Нормированный параметр точки сегмента [0;1], не расстояние по всему пути.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.BetweenAnchors
'
'       Возвращает точную цепочку между привязками с явным направлением и правилом полного обхода.
'
'       Объявление:
'         Public Function BetweenAnchors(ByVal pathID As String, ByVal startSegmentID As String, ByVal
'         startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, Optional
'         ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As
'         CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'         ByVal startSegmentID As String
'           Сегмент начальной привязки цепочки.
'
'         ByVal startParameter As Double
'           Параметр начальной точки выбранного сегмента [0;1].
'
'         ByVal endSegmentID As String
'           Сегмент конечной привязки цепочки.
'
'         ByVal endParameter As Double
'           Параметр конечной точки выбранного сегмента [0;1].
'
'         Optional ByVal forward As Boolean = True
'           Явное направление следования исходному обходу; False — противоположное.
'
'         Optional ByVal fullLoopIfSame As Boolean = False
'           Полный оборот при совпавших привязках замкнутого пути; иначе нулевая цепочка недопустима.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.AtDistance
'
'       Возвращает координаты точки на заданном расстоянии от начала линии и единичный вектор
'       касательной в этой точке. На стыке можно выбрать касательную предыдущего сегмента.
'
'       Объявление:
'         Public Sub AtDistance(ByVal distance As Double, ByRef x As Double, ByRef y As Double, ByRef
'         tangentX As Double, ByRef tangentY As Double, Optional ByVal incomingTangentAtJoint As Boolean
'         = False)
'
'       Параметры:
'         ByVal distance As Double
'           Расстояние или граница интервала distance, внутренние мм; направление и допустимые пределы
'           определяет вызываемый метод.
'
'         ByRef x As Double
'           Модельная координата x, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
'           скаляры/буфер можно изменять, исходная геометрия не меняется.
'
'         ByRef y As Double
'           Модельная координата y, мм; знаки соответствуют section XY/WCS. ByRef: возвращаемые
'           скаляры/буфер можно изменять, исходная геометрия не меняется.
'
'         ByRef tangentX As Double
'           Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно
'           изменять, исходная геометрия не меняется.
'
'         ByRef tangentY As Double
'           Компонента единичной касательной, без размерности. ByRef: возвращаемые скаляры/буфер можно
'           изменять, исходная геометрия не меняется.
'
'         Optional ByVal incomingTangentAtJoint As Boolean = False
'           Явный Boolean режима замыкания/касательной; не исправляет геометрию автоматически.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Slice
'
'       Обрезает полный непрерывный путь по длине; точные дуги и source parameters сохраняются.
'
'       Объявление:
'         Public Function Slice(ByVal startDistance As Double, ByVal endDistance As Double) As
'         CGeometryPath
'
'       Параметры:
'         ByVal startDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'         ByVal endDistance As Double
'           Расстояние от начала всей линии, мм, в [0;Length]; endDistance строго больше startDistance.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Reversed
'
'       Копирует путь с обратным порядком сегментов/обходом; signed sweep и source parameters
'       обращаются.
'
'       Объявление:
'         Public Function Reversed() As CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.Transformed
'
'       Копирует путь изометрией: отражение по локальной X, поворот, перенос; длина сохраняется.
'
'       Объявление:
'         Public Function Transformed(ByVal shiftX As Double, ByVal shiftY As Double, Optional ByVal
'         rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
'
'       Параметры:
'         ByVal shiftX As Double
'           Перенос по X, мм.
'
'         ByVal shiftY As Double
'           Перенос по Y, мм.
'
'         Optional ByVal rotation As Double = 0#
'           Угол поворота вокруг (0,0), радианы, CCW положителен.
'
'         Optional ByVal reflectX As Boolean = False
'           Отражение относительно локальной оси X перед поворотом и переносом.
'
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.WithID
'
'       Копирует опубликованный путь с новым ID, сохраняя сегментные ссылки и исходное происхождение.
'
'       Объявление:
'         Public Function WithID(ByVal pathID As String) As CGeometryPath
'
'       Параметры:
'         ByVal pathID As String
'           Непустой уникальный ID пути в определении.
'
'
'
'   Класс / модуль: modGeometryTypes
'
'     ----------------------------------------------------------------------------------------------
'     modGeometryTypes.GEOM_PI
'
'       Число pi как сумма сохраняемых VBE литералов, без изменения Double при компиляции.
'
'       Объявление:
'         Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
'
'     ----------------------------------------------------------------------------------------------
'     modGeometryTypes.TPoint2D
'
'       Два координатных скаляра X/Y, мм, без Excel или объекта на каждую точку.
'
'       Объявление:
'         Public Type TPoint2D
'
'       Допустимые значения / состав:
'         X As Double
'         Y As Double
'
'
'
' ====================================================================================================
' 09  АДРЕСА ВВОДА И СООБЩЕНИЯ ОБ ОШИБКАХ
' ====================================================================================================
'
'   Source связывает имя геометрического объекта с адресом исходных данных. Ошибка должна указывать,
'   какой размер или ряд нужно исправить. LastGeometryDiagnostic может вернуть Nothing, если сохранённой
'   ошибки нет.
'
'
'   Класс / модуль: CSectionDefinitionBuilder
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.Source
'
'       Связывает смысловой ID со строкой источника/адреса без сохранения объекта Excel.
'
'       Объявление:
'         Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
'
'       Параметры:
'         ByVal entityID As String
'           Имя объекта, которому назначается ссылка на исходные данные.
'
'         ByVal sourceReference As String
'           Строка исходного адреса/программного параметра; без объекта Excel.
'
'
'     ----------------------------------------------------------------------------------------------
'     CSectionDefinitionBuilder.LastGeometryDiagnostic
'
'       Возвращает последнюю геометрическую ошибку с адресом исходных данных. Если сохранённой ошибки
'       нет, возвращает Nothing.
'
'       Объявление:
'         Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
'
'
'   Класс / модуль: CGeometryPath
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryPath.LastDiagnostic
'
'       Последняя машинная диагностика; Nothing при отсутствии записанного события.
'
'       Объявление:
'         Public Property Get LastDiagnostic() As CGeometryDiagnostic
'
'
'   Класс / модуль: CGeometryDiagnostic
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Code
'
'       Машинный код причины; назначается напрямую, не извлекается из русского текста.
'
'       Объявление:
'         Public Property Get Code() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.EntityID
'
'       Смысловой ID сущности отказа, если он известен этому сервису.
'
'       Объявление:
'         Public Property Get EntityID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.RelatedEntityID
'
'       ID связанного сегмента/позиции/стержня для локализации причины.
'
'       Объявление:
'         Public Property Get RelatedEntityID() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Reason
'
'       Человекочитаемая причина; не selector/status.
'
'       Объявление:
'         Public Property Get Reason() As String
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Detail1
'
'       Первое численное поле конкретного Code; единицы определяет Code.
'
'       Объявление:
'         Public Property Get Detail1() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.Detail2
'
'       Второе численное поле конкретного Code.
'
'       Объявление:
'         Public Property Get Detail2() As Double
'
'     ----------------------------------------------------------------------------------------------
'     CGeometryDiagnostic.SourceReference
'
'       Адрес исходной ячейки либо пояснение, откуда взят параметр геометрии или арматуры.
'
'       Объявление:
'         Public Property Get SourceReference() As String
'
