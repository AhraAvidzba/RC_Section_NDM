VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNDMResultsWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Запись расчетного снимка на Results
' ==========================================================================
' Класс получает уже посчитанную модель сечения и результаты batch-расчета,
' переводит их в выбранные OUTPUT-единицы и записывает на лист Results одним
' согласованным снимком последнего успешного расчета.
'
' Writer не запускает solver, не генерирует геометрию и не читает текущие
' Inputs. Он только сериализует то состояние, которое уже рассчитано.

Private Const ELEMENT_RESULT_COLUMNS As Long = 6 ' LC-зависимые результаты: ElementID, LoadCase, Strain, Stress, PhysicalState.
Private Const GEOMETRY_COLUMNS As Long = 15 ' Постоянная геометрия элементов, записанная один раз на расчетный снимок.
Private Const SECTION_PROPERTY_COLUMNS As Long = 6 ' Общие свойства сечения и LC-свойства уровня всего сечения.
Private Const ANNOTATION_COLUMNS As Long = 13 ' Semantic-аннотации схемы: размеры, контур, групповые подписи арматуры.
Private Const MATERIAL_DIAGRAM_COLUMNS As Long = 8 ' Диагностический вывод фактических точек диаграмм, переданных solver-ам.
Private Const PHYSICAL_TOLERANCE As Double = 0.000000001 ' Порог, ниже которого напряжение считается NearZero.

' ------------------------------
' Полная запись Results
' ------------------------------
' Очищает текущий снимок и заново пишет четыре согласованных блока:
' geometry, element results, section properties и section annotations.
Public Sub WriteResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 4452, "CNDMResultsWriter", "Material provider не передан."
    If batch Is Nothing Then Err.Raise vbObjectError + 4454, "CNDMResultsWriter", "Пакетный расчетчик не передан."

    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSectionGeometry workbook, section, runID, units
    WriteElementResults workbook, section, batch, runID, units
    WriteSectionProperties workbook, section, materialProvider, batch, runID, units
    WriteMaterialDiagrams workbook, materialProvider, batch, runID, units
    EnsureAutoCADBoundsAnnotations section
    WriteSectionAnnotations workbook, section, runID, units
End Sub

' Записывает на Results только импортированную геометрию без расчетных LC.
' Такой снимок нужен после кнопки AutoCAD import: пользователь сразу видит
' элементы на схеме, а последующий расчет берет эту же геометрию из Results.
Public Sub WriteGeometryPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."

    ClearResults workbook

    Dim runID As Date
    runID = Now

    WriteSectionGeometry workbook, section, runID, units
    WriteEmptyElementResults workbook, runID, units
    WriteSectionPropertiesPreview workbook, section, runID, units
    EnsureAutoCADBoundsAnnotations section
    WriteSectionAnnotations workbook, section, runID, units
End Sub

' Очищает только расчетные блоки Results перед новым снимком.
Public Sub ClearResults(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4455, "CNDMResultsWriter", "Книга Excel не передана."

    ClearNamedRegion workbook, "rngNDMElementResults"
    ClearNamedRegion workbook, "rngNDMSectionGeometry"
    ClearNamedRegion workbook, "rngNDMSectionProperties"
    ClearNamedRegion workbook, "rngNDMSectionAnnotations"
    ClearNamedRegion workbook, "rngNDMMaterialDiagrams"
End Sub

' Безопасно чистит именованный диапазон, если он существует в книге.
Private Sub ClearNamedRegion(ByVal workbook As Object, ByVal rangeName As String)
    On Error GoTo MissingRange
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange
    anchor.CurrentRegion.ClearContents
    If anchor.Row > 1 Then anchor.Offset(-1, 0).Resize(1, 20).Clear
MissingRange:
End Sub

' ------------------------------
' LC-зависимые результаты элементов
' ------------------------------
' В этот блок не пишутся координаты и площади: они уже есть в
' rngNDMSectionGeometry и связываются по ElementID.
Private Sub WriteElementResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim elementCount As Long
    elementCount = section.ConcreteCount + section.RebarCount

    Dim activeCombinationCount As Long
    Dim i As Long
    ' Считаем только те сочетания, для которых solver действительно сохранил
    ' состояние. Ошибочные или пустые LC не должны создавать бессмысленные rows.
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then activeCombinationCount = activeCombinationCount + 1
    Next i

    Dim rowCount As Long
    rowCount = 1 + activeCombinationCount * elementCount
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ELEMENT_RESULT_COLUMNS)
    PutElementResultHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2
    ' Все строки собираются в массив VBA и одним присваиванием выгружаются
    ' на лист. Это принципиально быстрее и устойчивее построчной записи Cells.
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then
            PutCombinationElementResults data, rowIndex, section, batch, i, runID, units
        End If
    Next i

    target.Resize(rowCount, ELEMENT_RESULT_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, ELEMENT_RESULT_COLUMNS, "Результаты элементов по сочетаниям"
End Sub

' ------------------------------
' Геометрия расчетных элементов
' ------------------------------
' Геометрия не зависит от LoadCase, поэтому записывается один раз. Схема и
' AutoCAD export затем объединяют эту таблицу с нужным LC из ElementResults.
Private Sub WriteSectionGeometry(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim rowCount As Long
    rowCount = 1 + section.ConcreteCount + section.RebarCount

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To GEOMETRY_COLUMNS)
    PutGeometryHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    ' Геометрия записывается один раз на расчетный снимок. Она не зависит от LoadCase,
    ' поэтому координаты и площади не дублируются в rngNDMElementResults для каждого сочетания.
    For i = 1 To section.ConcreteCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.ConcreteID(i)
        data(rowIndex, 3) = "Concrete"
        data(rowIndex, 4) = OutputLength(section.ConcreteX(i), units)
        data(rowIndex, 5) = OutputLength(section.ConcreteY(i), units)
        data(rowIndex, 6) = OutputArea(section.ConcreteArea(i), units)
        data(rowIndex, 7) = section.ConcreteShapeType(i)
        data(rowIndex, 8) = OutputLength(section.ConcreteWidth(i), units)
        data(rowIndex, 9) = OutputLength(section.ConcreteHeight(i), units)
        data(rowIndex, 10) = vbNullString
        data(rowIndex, 11) = section.ConcreteRotation(i)
        data(rowIndex, 12) = OutputFourthPowerLength(section.ConcreteLocalIx(i), units)
        data(rowIndex, 13) = OutputFourthPowerLength(section.ConcreteLocalIy(i), units)
        data(rowIndex, 14) = OutputFourthPowerLength(section.ConcreteLocalIxy(i), units)
        data(rowIndex, 15) = section.ConcreteComment(i)
        rowIndex = rowIndex + 1
    Next i

    ' Арматура хранится в той же таблице геометрии, что и бетонные волокна, но с диаметром
    ' вместо width/height: downstream-слои различают тип элемента по MaterialType.
    For i = 1 To section.RebarCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.RebarID(i)
        data(rowIndex, 3) = "Rebar"
        data(rowIndex, 4) = OutputLength(section.RebarX(i), units)
        data(rowIndex, 5) = OutputLength(section.RebarY(i), units)
        data(rowIndex, 6) = OutputArea(section.RebarArea(i), units)
        data(rowIndex, 7) = "Circle"
        data(rowIndex, 8) = vbNullString
        data(rowIndex, 9) = vbNullString
        data(rowIndex, 10) = OutputLength(section.RebarDiameter(i), units)
        data(rowIndex, 11) = 0#
        data(rowIndex, 12) = 0#
        data(rowIndex, 13) = 0#
        data(rowIndex, 14) = 0#
        data(rowIndex, 15) = section.RebarComment(i)
        rowIndex = rowIndex + 1
    Next i

    target.Resize(rowCount, GEOMETRY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, GEOMETRY_COLUMNS, "Расчетная геометрия элементов"
End Sub

' ------------------------------
' Свойства сечения и LC-состояния
' ------------------------------
' Здесь лежат свойства уровня всей модели: Bounds, бетонные и приведенные
' геометрические характеристики, сохраненные единицы вывода и плоскость
' деформаций каждого LC. Нагрузки пользователя задаются от бетонного центра,
' а главные оси схемы пока используют приведенные свойства для совместимости.
Private Sub WriteSectionProperties(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim rowCount As Long
    rowCount = 1 + 92 + batch.Count * 26

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    ' Bounds нужны любой геометрии, Generated или AutoCAD. По ним схема может
    ' восстановить общий габарит без обращения к текущим настройкам Config.
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    ' Бетонные свойства нужны как база пользовательской точки нагрузки.
    ' Приведенные свойства сохраняются отдельно: они используются для справки,
    ' главных осей и совместимости существующего plot/export layer.
    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section

    Dim transformedProps As CSectionPropertiesCalculator
    Set transformedProps = New CSectionPropertiesCalculator
    transformedProps.CalculateTransformed section, materialProvider.ConcreteMaterial(cpStrength), _
        materialProvider.SteelMaterial(cpStrength)

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceX", OutputLength(batch.LoadReferenceX, units), OutputLengthUnit(units), "Общая X точки приложения нагрузки относительно центра тяжести бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "LoadReferenceY", OutputLength(batch.LoadReferenceY, units), OutputLengthUnit(units), "Общая Y точки приложения нагрузки относительно центра тяжести бетонного сечения."
    PutConcreteSectionProperties data, rowIndex, runID, concreteProps, units
    PutTransformedSectionProperties data, rowIndex, runID, transformedProps, units

    ' Совместимые имена без префикса пока означают приведенное сечение. Их читают
    ' существующие plot/export-компоненты, поэтому aliases оставлены явно.
    PutProperty data, rowIndex, runID, "ALL", "CentroidX", OutputLength(transformedProps.CentroidX, units), OutputLengthUnit(units), "Центр тяжести приведенного сечения; совместимое имя для схемы и AutoCAD."
    PutProperty data, rowIndex, runID, "ALL", "CentroidY", OutputLength(transformedProps.CentroidY, units), OutputLengthUnit(units), "Центр тяжести приведенного сечения; совместимое имя для схемы и AutoCAD."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalAngle", transformedProps.PrincipalAngleRad, "rad", "Угол главных центральных осей приведенного сечения; совместимое имя."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalI1", OutputFourthPowerLength(transformedProps.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный момент инерции I1 приведенного сечения; совместимое имя."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalI2", OutputFourthPowerLength(transformedProps.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный момент инерции I2 приведенного сечения; совместимое имя."
    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов арматуры."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.ForceUnit", OutputForceUnit(units), "-", "Единица вывода силы на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.MomentUnit", OutputMomentUnit(units), "-", "Единица вывода момента на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.N", OutputSignN(units), "-", "Пользовательский знак N на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.Mx", OutputSignMx(units), "-", "Пользовательский знак Mx на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.My", OutputSignMy(units), "-", "Пользовательский знак My на момент расчета."
    PutMaterialModelProperties data, rowIndex, runID, materialProvider, cpStrength
    PutMaterialModelProperties data, rowIndex, runID, materialProvider, cpMcrc
    PutMaterialModelProperties data, rowIndex, runID, materialProvider, cpCrackedNDS

    Dim i As Long
    ' LC-зависимые свойства уровня сечения хранятся отдельно от поэлементных
    ' строк: epsilon0/kappaX/kappaY не дублируются на каждое волокно.
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Epsilon0", batch.Epsilon0(i), "-", "Деформация в начале координат расчетной модели."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaX", OutputCurvature(batch.KappaX(i), units), OutputCurvatureUnit(units), "Кривизна KappaX."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaY", OutputCurvature(batch.KappaY(i), units), OutputCurvatureUnit(units), "Кривизна KappaY."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", OutputForce(batch.Nint(i), units), OutputForceUnit(units), "Внутренняя продольная сила."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Epsilon0", vbNullString, "-", "Деформация в начале координат расчетной модели."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaX", vbNullString, OutputCurvatureUnit(units), "Кривизна KappaX."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaY", vbNullString, OutputCurvatureUnit(units), "Кривизна KappaY."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", vbNullString, OutputForceUnit(units), "Внутренняя продольная сила."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "N", OutputForce(batch.N(i), units), OutputForceUnit(units), "Заданная продольная сила."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUser", OutputMx(batch.UserMx(i), units), OutputMomentUnit(units), "Пользовательский Mx."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUser", OutputMy(batch.UserMy(i), units), OutputMomentUnit(units), "Пользовательский My."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxInternal", OutputMx(batch.Mx(i), units), OutputMomentUnit(units), "Mx после переноса к координатам элементов."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyInternal", OutputMy(batch.My(i), units), OutputMomentUnit(units), "My после переноса к координатам элементов."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CalculationMode", batch.CalculationMode, "-", "Режим Calculation.Mode, по которому был выполнен последний расчет этого LC."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "OverallStatus", batch.OverallStatus(i), "-", "Итоговый пользовательский статус сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "DirectStateStatus", batch.DirectStateStatus(i), "-", "Статус прямого решения НДС фактического сочетания."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacityStatus", batch.CapacityStatus(i), "-", "Статус расчета несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacityLoadPath", batch.CapacityLoadPath(i), "-", "Какие компоненты нагрузки масштабировались при поиске несущей способности."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CapacitySolutionMethod", batch.CapacitySolutionMethod(i), "-", "Фактический численный метод, которым найдена предельная несущая способность."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LambdaCapacity", batch.LambdaCapacity(i), "-", "Предельный множитель для выбранной lambda-траектории."
        If Abs(batch.NUltimate(i)) > 0# Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", OutputForce(batch.NUltimate(i), units), OutputForceUnit(units), "Предельная продольная сила, если N входила в выбранную lambda-траекторию."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nultimate", vbNullString, OutputForceUnit(units), "Заполняется только для траекторий lambda*N и lambda*NMxy."
        End If
        If Abs(batch.MxUltimate(i)) > 0# Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", OutputMx(batch.MxUltimate(i), units), OutputMomentUnit(units), "Предельный пользовательский Mx, если Mx входил в выбранную lambda-траекторию."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUltimate", vbNullString, OutputMomentUnit(units), "Заполняется только для траекторий, где масштабируется Mx."
        End If
        If Abs(batch.MyUltimate(i)) > 0# Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", OutputMy(batch.MyUltimate(i), units), OutputMomentUnit(units), "Предельный пользовательский My, если My входил в выбранную lambda-траекторию."
        Else
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUltimate", vbNullString, OutputMomentUnit(units), "Заполняется только для траекторий, где масштабируется My."
        End If
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "CrackStatus", batch.CrackStatus(i), "-", "Статус расчета ширины раскрытия трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrackStatus", batch.LongitudinalCrackStatus(i), "-", "Статус SLS-проверки продольных трещин по сжатому бетону; для Group1 не применяется."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.MaxConcreteCompressionStress", OutputStress(batch.MaxConcreteCompressionStress(i), units), OutputStressUnit(units), "Максимальное по модулю сжимающее напряжение бетона в прямом НДС."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.RbMc2", OutputStress(batch.LongitudinalCrackRbMc2(i), units), OutputStressUnit(units), "Пользовательское допустимое Rb,mc2 для проверки продольных трещин."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "LongitudinalCrack.Utilization", batch.LongitudinalCrackUtilization(i), "-", "Использование по продольным трещинам: sigma_c,max / Rb,mc2."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "ExtensionUsed", BooleanText(batch.ExtensionUsed(i)), "-", "True, если прямой StateSolution вышел за физические пределы диаграммы в numerical extension."
        PutProperty data, rowIndex, runID, batch.CombinationID(i), "DirectState.DiagramExtensionEnabled", BooleanText(batch.DirectStateDiagramExtensionEnabled(i)), "-", "Сохраненное значение настройки Solver.DirectState.DiagramExtension на момент расчета."
    Next i

    target.Resize(rowCount, SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' Возвращает английскую запись логического флага для Results snapshot.
' Так файл не зависит от локали Excel, где Boolean может отображаться как
' ИСТИНА/ЛОЖЬ, TRUE/FALSE или в другом переводе интерфейса.
Private Function BooleanText(ByVal value As Boolean) As String
    If value Then
        BooleanText = "True"
    Else
        BooleanText = "False"
    End If
End Function

' ------------------------------
' Диагностический вывод фактических диаграмм
' ------------------------------
' Пишет точки CMaterialDiagram, которые solver/provider использовали для
' прямого НДС, capacity и crack-веток каждого LC. Это временный контрольный
' блок Results: расчетное ядро и plotter его не читают.
Private Sub WriteMaterialDiagrams(ByVal workbook As Object, ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMMaterialDiagrams").RefersToRange

    Dim rowCapacity As Long
    rowCapacity = 1 + batch.Count * 128

    Dim data() As Variant
    ReDim data(1 To rowCapacity, 1 To MATERIAL_DIAGRAM_COLUMNS)
    PutMaterialDiagramHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    For i = 1 To batch.Count
        PutCombinationMaterialDiagrams data, rowIndex, runID, materialProvider, batch, i, units
    Next i

    Dim rowCount As Long
    rowCount = rowIndex - 1
    If rowCount < 1 Then rowCount = 1

    target.Resize(rowCount, MATERIAL_DIAGRAM_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, MATERIAL_DIAGRAM_COLUMNS, "Фактические диаграммы материалов"
End Sub

' Собирает материалные диаграммы одного сочетания в том же порядке, в каком
' пользователь обычно читает Summary: прямое состояние, capacity, cracks.
Private Sub PutCombinationMaterialDiagrams(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal materialProvider As CMaterialModelProvider, _
        ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, ByVal units As CUnitSystem)
    Dim policy As CBatchStatusPolicy
    Set policy = New CBatchStatusPolicy

    Dim basePurpose As ECalculationPurpose
    basePurpose = StateBasePurposeForCalculationType(batch.CalculationType(combinationIndex))

    If batch.StateAvailable(combinationIndex) Then
        Dim stateConcrete As CMaterialDiagram
        Dim stateSteel As CMaterialDiagram
        Set stateConcrete = batch.ConcreteMaterialForCombination(combinationIndex)
        Set stateSteel = batch.SteelMaterialForCombination(combinationIndex)
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "DirectState", "StateSolution/" & PurposeToText(basePurpose), "Concrete", stateConcrete, units
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "DirectState", "StateSolution/" & PurposeToText(basePurpose), "Rebar", stateSteel, units
    End If

    If policy.ToUserStatus(batch.CapacityStatus(combinationIndex)) <> policy.NotApplicable And _
            Len(Trim$(batch.CapacityStatus(combinationIndex))) > 0 Then
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "Capacity", PurposeToText(cpStrength), "Concrete", materialProvider.ConcreteMaterial(cpStrength), units
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "Capacity", PurposeToText(cpStrength), "Rebar", materialProvider.SteelMaterial(cpStrength), units
    End If

    If policy.ToUserStatus(batch.CrackStatus(combinationIndex)) <> policy.NotApplicable And _
            Len(Trim$(batch.CrackStatus(combinationIndex))) > 0 Then
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "Crack.CrackedNDS", PurposeToText(cpCrackedNDS), "Concrete", materialProvider.ConcreteMaterial(cpCrackedNDS), units
        AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
            "Crack.CrackedNDS", PurposeToText(cpCrackedNDS), "Rebar", materialProvider.SteelMaterial(cpCrackedNDS), units
        If batch.CrackLambdaCrc(combinationIndex) > 0# Then
            AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
                "Crack.Mcrc", PurposeToText(cpMcrc), "Concrete", materialProvider.ConcreteMaterial(cpMcrc), units
            AppendMaterialDiagramPoints data, rowIndex, runID, batch.CombinationID(combinationIndex), _
                "Crack.Mcrc", PurposeToText(cpMcrc), "Rebar", materialProvider.SteelMaterial(cpMcrc), units
        End If
    End If
End Sub

' Добавляет все точки одной диаграммы. Напряжения пишутся в OUTPUT stress unit,
' деформации остаются безразмерными, как их получает solver.
Private Sub AppendMaterialDiagramPoints(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal combinationID As String, ByVal calculationStage As String, _
        ByVal purposeText As String, ByVal materialKind As String, ByVal diagram As CMaterialDiagram, _
        ByVal units As CUnitSystem)
    If diagram Is Nothing Then Exit Sub

    Dim pointIndex As Long
    For pointIndex = 1 To diagram.PointCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = combinationID
        data(rowIndex, 3) = calculationStage
        data(rowIndex, 4) = purposeText
        data(rowIndex, 5) = materialKind
        data(rowIndex, 6) = pointIndex
        data(rowIndex, 7) = diagram.PointStrain(pointIndex)
        data(rowIndex, 8) = OutputStress(diagram.PointStress(pointIndex), units)
        rowIndex = rowIndex + 1
    Next pointIndex
End Sub

' Шапка диагностической таблицы диаграмм материалов.
Private Sub PutMaterialDiagramHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "CalculationStage"
    data(1, 4) = "Purpose"
    data(1, 5) = "Material"
    data(1, 6) = "PointIndex"
    data(1, 7) = "Strain"
    data(1, 8) = "Stress, " & OutputStressUnit(units)
End Sub

' Записывает справочные свойства бетонной части сечения.
' Эти строки нужны, чтобы на Results было видно, от какой точки реально
' задается Load.ReferenceOffset и какие геометрические характеристики имеет
' бетонное сечение без учета продольной арматуры.
Private Sub PutConcreteSectionProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal props As CSectionPropertiesCalculator, ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Area", OutputArea(props.Area, units), OutputAreaUnit(units), "Площадь бетонной части сечения по расчетной сетке."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести бетонного сечения; от этой точки задается Load.ReferenceOffsetX."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести бетонного сечения; от этой точки задается Load.ReferenceOffsetY."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно исходной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции бетонного сечения относительно исходной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент бетонного сечения относительно исходных осей."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции бетонного сечения относительно оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции бетонного сечения относительно оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центральный центробежный момент бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол главных центральных осей бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Больший главный момент инерции бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Меньший главный момент инерции бетонного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения относительно центральной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения по главной оси I1."
    PutProperty data, rowIndex, runID, "ALL", "Concrete.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции бетонного сечения по главной оси I2."
End Sub

' Записывает свойства приведенного сечения бетон + арматура.
' Они остаются отдельным набором справочных данных и совместимы с уже
' существующей логикой главных осей на схеме и в AutoCAD export.
Private Sub PutTransformedSectionProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal props As CSectionPropertiesCalculator, ByVal units As CUnitSystem)
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Area", OutputArea(props.Area, units), OutputAreaUnit(units), "Площадь приведенного сечения в бетонной базе."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "X центра тяжести приведенного сечения бетон + арматура."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Y центра тяжести приведенного сечения бетон + арматура."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ix", OutputFourthPowerLength(props.Ix, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно исходной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iy", OutputFourthPowerLength(props.Iy, units), OutputLengthUnit(units) & "4", "Момент инерции приведенного сечения относительно исходной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixy", OutputFourthPowerLength(props.Ixy, units), OutputLengthUnit(units) & "4", "Центробежный момент приведенного сечения относительно исходных осей."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixc", OutputFourthPowerLength(props.Ixc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции приведенного сечения относительно оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Iyc", OutputFourthPowerLength(props.Iyc, units), OutputLengthUnit(units) & "4", "Центральный момент инерции приведенного сечения относительно оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.Ixyc", OutputFourthPowerLength(props.Ixyc, units), OutputLengthUnit(units) & "4", "Центральный центробежный момент приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол главных центральных осей приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный момент инерции I1 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный момент инерции I2 приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusX", OutputLength(props.RadiusX, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси X."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.RadiusY", OutputLength(props.RadiusY, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения относительно центральной оси Y."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius1", OutputLength(props.PrincipalRadius1, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения по главной оси I1."
    PutProperty data, rowIndex, runID, "ALL", "Transformed.PrincipalRadius2", OutputLength(props.PrincipalRadius2, units), OutputLengthUnit(units), "Радиус инерции приведенного сечения по главной оси I2."
End Sub

' Записывает в Results краткое описание материалной модели расчетного snapshot.
' Эти строки нужны для диагностики и для внешних выводов, которые не должны
' перечитывать текущий Config после расчета.
Private Sub PutMaterialModelProperties(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal runID As Variant, ByVal materialProvider As CMaterialModelProvider, _
        ByVal purpose As ECalculationPurpose)
    Dim prefix As String
    prefix = "Material." & PurposeToText(purpose) & "."
    PutProperty data, rowIndex, runID, "ALL", prefix & "LimitStateGroup", materialProvider.LimitStateGroup(purpose), "-", "Группа предельных состояний material model."
    PutProperty data, rowIndex, runID, "ALL", prefix & "ConcreteDiagram", materialProvider.ConcreteDiagramType(purpose), "-", "Тип бетонной диаграммы snapshot."
    PutProperty data, rowIndex, runID, "ALL", prefix & "ConcreteTension", materialProvider.ConcreteTensionBehavior(purpose), "-", "Учет растянутого бетона в бетонной диаграмме snapshot."
    PutProperty data, rowIndex, runID, "ALL", prefix & "SteelDiagram", materialProvider.SteelDiagramType(purpose), "-", "Тип диаграммы арматуры snapshot."
    PutProperty data, rowIndex, runID, "ALL", prefix & "ConcreteCompressionLimit", materialProvider.ConcreteCompressionLimit(purpose), "-", "Предельная сжатая деформация бетона snapshot."
    If materialProvider.ConcreteTensionLimitEnabled(purpose) Then
        PutProperty data, rowIndex, runID, "ALL", prefix & "ConcreteTensionLimit", materialProvider.ConcreteTensionLimit(purpose), "-", "Предельная растянутая деформация бетона snapshot."
    End If
    PutProperty data, rowIndex, runID, "ALL", prefix & "SteelCompressionLimit", materialProvider.SteelCompressionLimit(purpose), "-", "Предельная сжатая деформация арматуры snapshot."
    PutProperty data, rowIndex, runID, "ALL", prefix & "SteelTensionLimit", materialProvider.SteelTensionLimit(purpose), "-", "Предельная растянутая деформация арматуры snapshot."
End Sub

' Записывает минимальные свойства imported-preview: Bounds, количество
' элементов и единицы, достаточные для проверки геометрии на схеме.
Private Sub WriteSectionPropertiesPreview(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 35, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Geometry.Source", section.SourceType, "-", "Источник геометрии расчетной модели."

    Dim concreteProps As CSectionPropertiesCalculator
    Set concreteProps = New CSectionPropertiesCalculator
    concreteProps.CalculateConcrete section
    PutConcreteSectionProperties data, rowIndex, runID, concreteProps, units

    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов арматуры."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения на момент импорта."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны на момент импорта."

    target.Resize(UBound(data, 1), SECTION_PROPERTY_COLUMNS).Value2 = data
    FormatResultsBlock target, UBound(data, 1), SECTION_PROPERTY_COLUMNS, "Свойства всего сечения"
End Sub

' ------------------------------
' Semantic-аннотации схемы
' ------------------------------
' Эти строки отвечают не за расчет, а за оформление последнего рассчитанного
' сечения: контуры, размеры и групповые подписи арматуры.
Private Sub WriteSectionAnnotations(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim annotations As CSectionAnnotations
    Set annotations = section.Annotations

    Dim rowCount As Long
    rowCount = 1 + annotations.Count
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    For i = 1 To annotations.Count
        PutAnnotationRow data, rowIndex, runID, annotations, i, units
    Next i

    target.Resize(rowCount, ANNOTATION_COLUMNS).Value2 = data
    FormatResultsBlock target, rowCount, ANNOTATION_COLUMNS, "Аннотации схемы"
End Sub

' Оставляет таблицу результатов элементов с одной шапкой.
' Для preview нет LC, поэтому записывать Stress/Strain нельзя.
Private Sub WriteEmptyElementResults(ByVal workbook As Object, ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 1, 1 To ELEMENT_RESULT_COLUMNS)
    PutElementResultHeaders data, units

    target.Resize(1, ELEMENT_RESULT_COLUMNS).Value2 = data
    FormatResultsBlock target, 1, ELEMENT_RESULT_COLUMNS, "Результаты элементов по сочетаниям"
End Sub

' Оставляет таблицу аннотаций с одной шапкой.
' У импортированной AutoCAD-сетки semantic-граней обычно нет, поэтому preview
' не должен наследовать старые размеры и подписи.
Private Sub WriteEmptySectionAnnotations(ByVal workbook As Object, ByVal runID As Variant)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 1, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    target.Resize(1, ANNOTATION_COLUMNS).Value2 = data
    FormatResultsBlock target, 1, ANNOTATION_COLUMNS, "Аннотации схемы"
End Sub

' Оформляет один Results-диапазон как самостоятельную таблицу.
' Заголовок ставится прямо над якорем, а чтение Results ниже выполняется
' явным anchored-reader-ом, поэтому внешний вид больше не зависит от CurrentRegion.
Private Sub FormatResultsBlock(ByVal target As Object, ByVal rowCount As Long, ByVal columnCount As Long, ByVal titleText As String)
    Dim tableRange As Object
    Set tableRange = target.Resize(rowCount, columnCount)

    tableRange.HorizontalAlignment = -4108
    tableRange.VerticalAlignment = -4108
    tableRange.Font.Name = "Arial"
    tableRange.Font.Size = 9
    tableRange.Borders.LineStyle = 1
    tableRange.Borders.Color = RGB(200, 200, 200)
    tableRange.Rows.Item(1).Interior.Color = RGB(217, 217, 217)
    tableRange.Rows.Item(1).Font.Bold = True
    tableRange.Rows.Item(1).HorizontalAlignment = -4131
    tableRange.Columns.Item(1).NumberFormat = "dd.mm.yy"

    AlignCommentColumns tableRange

    If target.Row > 2 Then
        Dim titleRange As Object
        Set titleRange = target.Offset(-1, 0).Resize(1, columnCount)
        titleRange.UnMerge
        titleRange.Clear
        titleRange.Merge
        titleRange.Value2 = titleText
        titleRange.Font.Name = "Arial"
        titleRange.Font.Size = 9
        titleRange.Font.Bold = True
        titleRange.HorizontalAlignment = -4108
        titleRange.VerticalAlignment = -4108
        titleRange.Interior.Color = RGB(217, 217, 217)
    End If
End Sub

' Ищет текстовые диагностические столбцы и выравнивает их вправо.
' Comment и PhysicalState читаются удобнее как текстовые признаки, поэтому
' они не центрируются вместе с числовыми данными таблиц Results.
Private Sub AlignCommentColumns(ByVal tableRange As Object)
    Dim headers As Variant
    headers = tableRange.Rows.Item(1).Value2

    Dim colIndex As Long
    For colIndex = 1 To tableRange.Columns.Count
        If ShouldRightAlignResultColumn(ResultHeaderBase(CStr(headers(1, colIndex)))) Then
            If tableRange.Rows.Count > 1 Then
                tableRange.Offset(1, colIndex - 1).Resize(tableRange.Rows.Count - 1, 1).HorizontalAlignment = -4152
            End If
        End If
    Next colIndex
End Sub

' Определяет текстовые столбцы Results, которые должны читаться как подпись,
' а не как числовой параметр. Пока это комментарии и физическое состояние
' элемента в rngNDMElementResults.
Private Function ShouldRightAlignResultColumn(ByVal headerBase As String) As Boolean
    ShouldRightAlignResultColumn = _
        (StrComp(headerBase, "Comment", vbTextCompare) = 0 Or _
         StrComp(headerBase, "PhysicalState", vbTextCompare) = 0)
End Function

' Возвращает имя столбца без единицы измерения.
' Например, заголовок "Stress, MPa" для служебных проверок считается просто
' "Stress"; это позволяет одинаково находить Comment и другие столбцы при
' разных пользовательских единицах вывода.
Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

Private Sub PutCombinationElementResults(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal section As CSectionModel, ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal runID As Variant, ByVal units As CUnitSystem)
    Dim i As Long
    Dim strain As Double
    Dim stress As Double
    Dim concrete As Object
    Dim steel As Object

    Set concrete = batch.ConcreteMaterialForCombination(combinationIndex)
    Set steel = batch.SteelMaterialForCombination(combinationIndex)

    For i = 1 To section.ConcreteCount
        strain = batch.Epsilon0(combinationIndex) + batch.KappaX(combinationIndex) * section.ConcreteY(i) + _
            batch.KappaY(combinationIndex) * section.ConcreteX(i)
        stress = concrete.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), section.ConcreteID(i), _
            strain, stress, PhysicalState("Concrete", strain, stress), units
        rowIndex = rowIndex + 1
    Next i

    For i = 1 To section.RebarCount
        strain = batch.Epsilon0(combinationIndex) + batch.KappaX(combinationIndex) * section.RebarY(i) + _
            batch.KappaY(combinationIndex) * section.RebarX(i)
        stress = steel.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), section.RebarID(i), _
            strain, stress, PhysicalState("Rebar", strain, stress), units
        rowIndex = rowIndex + 1
    Next i
End Sub

Private Sub PutElementRow(ByRef data() As Variant, ByVal rowIndex As Long, ByVal runID As Variant, _
        ByVal combinationID As String, ByVal elementID As String, _
        ByVal strain As Double, ByVal stress As Double, ByVal state As String, ByVal units As CUnitSystem)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = combinationID
    data(rowIndex, 3) = elementID
    data(rowIndex, 4) = strain
    data(rowIndex, 5) = OutputStress(stress, units)
    data(rowIndex, 6) = state
End Sub

Private Function PhysicalState(ByVal materialType As String, ByVal strain As Double, ByVal stress As Double) As String
    If Abs(stress) <= PHYSICAL_TOLERANCE Then
        If StrComp(materialType, "Concrete", vbTextCompare) = 0 And strain > PHYSICAL_TOLERANCE Then
            PhysicalState = "InactiveTensionConcrete"
        Else
            PhysicalState = "NearZero"
        End If
    ElseIf stress < 0# Then
        PhysicalState = "Compression"
    Else
        PhysicalState = "Tension"
    End If
End Function

Private Sub PutElementResultHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "ElementID"
    data(1, 4) = "Strain"
    data(1, 5) = "Stress, " & OutputStressUnit(units)
    data(1, 6) = "PhysicalState"
End Sub

Private Sub PutGeometryHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "ElementID"
    data(1, 3) = "MaterialType"
    data(1, 4) = "X, " & OutputLengthUnit(units)
    data(1, 5) = "Y, " & OutputLengthUnit(units)
    data(1, 6) = "Area, " & OutputAreaUnit(units)
    data(1, 7) = "ShapeType"
    data(1, 8) = "Width, " & OutputLengthUnit(units)
    data(1, 9) = "Height, " & OutputLengthUnit(units)
    data(1, 10) = "Diameter, " & OutputLengthUnit(units)
    data(1, 11) = "Rotation, rad"
    data(1, 12) = "LocalIx, " & OutputLengthUnit(units) & "4"
    data(1, 13) = "LocalIy, " & OutputLengthUnit(units) & "4"
    data(1, 14) = "LocalIxy, " & OutputLengthUnit(units) & "4"
    data(1, 15) = "Comment"
End Sub

Private Sub PutSectionPropertyHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "Parameter"
    data(1, 4) = "Value"
    data(1, 5) = "Unit"
    data(1, 6) = "Comment"
End Sub

Private Sub PutAnnotationHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "AnnotationType"
    data(1, 3) = "AnnotationID"
    data(1, 4) = "StartX"
    data(1, 5) = "StartY"
    data(1, 6) = "EndX"
    data(1, 7) = "EndY"
    data(1, 8) = "OutsideNormalX"
    data(1, 9) = "OutsideNormalY"
    data(1, 10) = "Text"
    data(1, 11) = "Value"
    data(1, 12) = "Unit"
    data(1, 13) = "Comment"
End Sub

Private Sub PutProperty(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal loadCase As String, ByVal parameter As String, ByVal value As Variant, _
        ByVal unitText As String, ByVal comment As String)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = loadCase
    data(rowIndex, 3) = parameter
    data(rowIndex, 4) = value
    data(rowIndex, 5) = unitText
    data(rowIndex, 6) = comment
    rowIndex = rowIndex + 1
End Sub

Private Sub PutAnnotationRow(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As Variant, _
        ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, ByVal units As CUnitSystem)
    Dim annotationType As String
    annotationType = annotations.AnnotationType(annotationIndex)

    data(rowIndex, 1) = runID
    data(rowIndex, 2) = annotationType
    data(rowIndex, 3) = annotations.AnnotationID(annotationIndex)
    data(rowIndex, 4) = OutputLength(annotations.StartX(annotationIndex), units)
    data(rowIndex, 5) = OutputLength(annotations.StartY(annotationIndex), units)
    data(rowIndex, 6) = OutputLength(annotations.EndX(annotationIndex), units)
    data(rowIndex, 7) = OutputLength(annotations.EndY(annotationIndex), units)
    data(rowIndex, 8) = annotations.NormalX(annotationIndex)
    data(rowIndex, 9) = annotations.NormalY(annotationIndex)
    data(rowIndex, 10) = OutputAnnotationText(annotations, annotationIndex, units)
    data(rowIndex, 11) = OutputAnnotationValue(annotations, annotationIndex, units)
    data(rowIndex, 12) = OutputAnnotationUnit(annotations, annotationIndex, units)
    data(rowIndex, 13) = OutputAnnotationComment(annotations, annotationIndex)
    rowIndex = rowIndex + 1
End Sub

Private Function OutputAnnotationText(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As String
    Dim annotationType As String
    annotationType = annotations.AnnotationType(annotationIndex)
    If StrComp(annotationType, "DIMENSION", vbTextCompare) <> 0 Then
        OutputAnnotationText = annotations.Text(annotationIndex)
        Exit Function
    End If

    Dim valueText As String
    valueText = FormatNumberInvariant(OutputLength(annotations.Value(annotationIndex), units))
    Dim unitText As String
    unitText = OutputLengthUnit(units)
    If StrComp(Trim$(annotations.Text(annotationIndex)), ApproxSign(), vbTextCompare) = 0 Then
        OutputAnnotationText = ApproxSign() & " " & valueText & " " & unitText
    ElseIf Len(Trim$(annotations.Text(annotationIndex))) > 0 Then
        OutputAnnotationText = annotations.Text(annotationIndex) & "=" & valueText & " " & unitText
    Else
        OutputAnnotationText = valueText & " " & unitText
    End If
End Function

Private Function OutputAnnotationComment(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long) As String
    Dim commentValue As String
    commentValue = Trim$(annotations.Comment(annotationIndex))

    If StrComp(commentValue, "Generated section contour segment.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Сегмент контура параметрического сечения."
    ElseIf StrComp(commentValue, "Generated circle contour.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Контур круглого параметрического сечения."
    ElseIf StrComp(commentValue, "Section dimension.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Размер сечения для отображения на схеме."
    ElseIf StrComp(commentValue, "Grouped rebar annotation.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Групповая подпись арматуры по смысловой грани."
    ElseIf StrComp(commentValue, "Overall concrete width.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Общая ширина бетонного сечения."
    ElseIf StrComp(commentValue, "Overall concrete height.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Общая высота бетонного сечения."
    ElseIf StrComp(commentValue, "Upper rectangle width.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Ширина верхнего прямоугольника Г-образного сечения."
    ElseIf StrComp(commentValue, "Lower rectangle height.", vbTextCompare) = 0 Then
        OutputAnnotationComment = "Высота нижнего прямоугольника Г-образного сечения."
    ElseIf Len(commentValue) > 0 Then
        OutputAnnotationComment = commentValue
    Else
        OutputAnnotationComment = "Semantic-аннотация схемы."
    End If
End Function

' Для AutoCAD-сетки у программы нет смысловых граней, но есть надежные Bounds
' бетонных элементов. Поэтому сохраняем только две приблизительные размерные
' аннотации: общую ширину и высоту. Знак приблизительного равенства показывает, что размер получен
' по дискретной сетке Region, а не из параметрической геометрии.
Private Sub EnsureAutoCADBoundsAnnotations(ByVal section As CSectionModel)
    If section Is Nothing Then Exit Sub
    If StrComp(section.SourceType, "AutoCADImport", vbTextCompare) <> 0 Then Exit Sub
    If section.AnnotationCount > 0 Then Exit Sub
    If section.ConcreteCount <= 0 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    section.Annotations.AddDimension "DIM_AUTO_BOUNDS_B", minX, minY, maxX, minY, 0#, -1#, _
        ApproxSign(), maxX - minX, "Приблизительная общая ширина по Bounds импортированного бетона."
    section.Annotations.AddDimension "DIM_AUTO_BOUNDS_H", minX, minY, minX, maxY, -1#, 0#, _
        ApproxSign(), maxY - minY, "Приблизительная общая высота по Bounds импортированного бетона."
End Sub

' Возвращает знак приблизительного равенства без литерала в исходнике. Так VBA-модуль
' не зависит от кодировки файла, а на лист Results попадает настоящий знак приблизительного равенства.
Private Function ApproxSign() As String
    ApproxSign = ChrW$(&H2248)
End Function

Private Function OutputAnnotationValue(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As Variant
    If StrComp(annotations.Unit(annotationIndex), "Length", vbTextCompare) = 0 Then
        OutputAnnotationValue = OutputLength(annotations.Value(annotationIndex), units)
    ElseIf Len(annotations.Unit(annotationIndex)) = 0 Then
        OutputAnnotationValue = vbNullString
    Else
        OutputAnnotationValue = annotations.Value(annotationIndex)
    End If
End Function

Private Function OutputAnnotationUnit(ByVal annotations As CSectionAnnotations, ByVal annotationIndex As Long, _
        ByVal units As CUnitSystem) As String
    If StrComp(annotations.Unit(annotationIndex), "Length", vbTextCompare) = 0 Then
        OutputAnnotationUnit = OutputLengthUnit(units)
    Else
        OutputAnnotationUnit = annotations.Unit(annotationIndex)
    End If
End Function

Private Sub GetConcreteBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4460, "CNDMResultsWriter", "В модели нет бетонных элементов."

    minX = section.ConcreteX(1) - ConcreteWidth(section, 1) / 2#
    maxX = section.ConcreteX(1) + ConcreteWidth(section, 1) / 2#
    minY = section.ConcreteY(1) - ConcreteHeight(section, 1) / 2#
    maxY = section.ConcreteY(1) + ConcreteHeight(section, 1) / 2#

    Dim i As Long
    For i = 1 To section.ConcreteCount
        minX = MinDouble(minX, section.ConcreteX(i) - ConcreteWidth(section, i) / 2#)
        maxX = MaxDouble(maxX, section.ConcreteX(i) + ConcreteWidth(section, i) / 2#)
        minY = MinDouble(minY, section.ConcreteY(i) - ConcreteHeight(section, i) / 2#)
        maxY = MaxDouble(maxY, section.ConcreteY(i) + ConcreteHeight(section, i) / 2#)
    Next i
End Sub

Private Sub GetSectionBoundsUnused(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4460, "CNDMResultsWriter", "В модели нет бетонных элементов."

    minX = section.ConcreteX(1) - ConcreteWidth(section, 1) / 2#
    maxX = section.ConcreteX(1) + ConcreteWidth(section, 1) / 2#
    minY = section.ConcreteY(1) - ConcreteHeight(section, 1) / 2#
    maxY = section.ConcreteY(1) + ConcreteHeight(section, 1) / 2#

    Dim i As Long
    For i = 1 To section.ConcreteCount
        minX = MinDouble(minX, section.ConcreteX(i) - ConcreteWidth(section, i) / 2#)
        maxX = MaxDouble(maxX, section.ConcreteX(i) + ConcreteWidth(section, i) / 2#)
        minY = MinDouble(minY, section.ConcreteY(i) - ConcreteHeight(section, i) / 2#)
        maxY = MaxDouble(maxY, section.ConcreteY(i) + ConcreteHeight(section, i) / 2#)
    Next i
    For i = 1 To section.RebarCount
        minX = MinDouble(minX, section.RebarX(i) - section.RebarDiameter(i) / 2#)
        maxX = MaxDouble(maxX, section.RebarX(i) + section.RebarDiameter(i) / 2#)
        minY = MinDouble(minY, section.RebarY(i) - section.RebarDiameter(i) / 2#)
        maxY = MaxDouble(maxY, section.RebarY(i) + section.RebarDiameter(i) / 2#)
    Next i
End Sub

Private Function ConcreteWidth(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteWidth = section.ConcreteWidth(index)
    If ConcreteWidth <= 0# Then ConcreteWidth = Sqr(section.ConcreteArea(index))
End Function

Private Function ConcreteHeight(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteHeight = section.ConcreteHeight(index)
    If ConcreteHeight <= 0# Then ConcreteHeight = Sqr(section.ConcreteArea(index))
End Function

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

Private Function OutputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputFourthPowerLength = value Else OutputFourthPowerLength = units.InternalFourthPowerLengthToOutput(value)
End Function

Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value Else OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

Private Function OutputSignN(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignN = "Tension" Else OutputSignN = units.OutputSignN
End Function

Private Function OutputSignMx(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMx = "+Y tension" Else OutputSignMx = units.OutputSignMx
End Function

Private Function OutputSignMy(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMy = "+X tension" Else OutputSignMy = units.OutputSignMy
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.###"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatNumberInvariant = textValue
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
