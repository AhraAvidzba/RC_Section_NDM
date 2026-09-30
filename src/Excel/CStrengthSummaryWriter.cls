VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStrengthSummaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подробная таблица НДС по прочности и несущей способности на листе Results
' ==========================================================================
' Класс пишет отдельный блок от якоря rngStrengthSummaryAnchor: прямое НДС
' заданного сочетания по модели прочности и расчет предельной несущей
' способности по той же strength-модели. Расчет здесь не выполняется: writer
' берет уже сохраненные named states StrengthState и CapacityState, переводит
' величины в пользовательские единицы и оформляет таблицу. Верхний
' rngBatchSummary остается компактной сводкой, а этот блок раскрывает именно
' результаты НДС по модели прочности, включая несущую способность.
'
' Для отсутствующих физических зон writer оставляет пустые ячейки. Например,
' если в диаграмме бетона растяжение отключено или все сечение растянуто, то
' сжатые/растянутые деформации и напряжения не подменяются ближайшими
' значениями другого знака.

Private Const HEADER_ROWS As Long = 4 ' Строки шапки над первой строкой LC.
Private Const TOTAL_COLUMNS As Long = 49 ' Полная ширина с двумя пустыми разделителями.

Private Const COL_LEFT_FIRST As Long = 1
Private Const COL_LEFT_LAST As Long = 9
Private Const COL_SEPARATOR1 As Long = 10
Private Const COL_STATE_FIRST As Long = 11
Private Const COL_STATE_LAST As Long = 30
Private Const COL_SEPARATOR2 As Long = 31
Private Const COL_CAPACITY_FIRST As Long = 32
Private Const COL_CAPACITY_LAST As Long = 49

Private Const COL_ID As Long = 1
Private Const COL_RESULT_COMMENT As Long = 2
Private Const COL_STATUS As Long = 3
Private Const COL_REF_N As Long = 4
Private Const COL_REF_MX As Long = 5
Private Const COL_REF_MY As Long = 6
Private Const COL_LOAD_N As Long = 7
Private Const COL_LOAD_MX As Long = 8
Private Const COL_LOAD_MY As Long = 9

Private Const COL_LIM_EBC As Long = 11
Private Const COL_LIM_EBT As Long = 12
Private Const COL_LIM_ESC As Long = 13
Private Const COL_LIM_EST As Long = 14
Private Const COL_STATE_EPS0 As Long = 15
Private Const COL_STATE_KX As Long = 16
Private Const COL_STATE_KY As Long = 17
Private Const COL_STATE_EBC As Long = 18
Private Const COL_STATE_EBT As Long = 19
Private Const COL_STATE_ESC As Long = 20
Private Const COL_STATE_EST As Long = 21
Private Const COL_STATE_SBC As Long = 22
Private Const COL_STATE_SBT As Long = 23
Private Const COL_STATE_SSC As Long = 24
Private Const COL_STATE_SST As Long = 25
Private Const COL_STATE_H As Long = 26
Private Const COL_STATE_X As Long = 27
Private Const COL_STATE_DIAGEXT As Long = 28
Private Const COL_STATE_RESERVE As Long = 29
Private Const COL_STATE_STATUS As Long = 30

Private Const COL_CAP_EPS0 As Long = 32
Private Const COL_CAP_KX As Long = 33
Private Const COL_CAP_KY As Long = 34
Private Const COL_CAP_LIMIT_STATE As Long = 35
Private Const COL_CAP_LOAD_PATH As Long = 36
Private Const COL_CAP_METHOD As Long = 37
Private Const COL_CAP_LAMBDA As Long = 38
Private Const COL_CAP_N As Long = 39
Private Const COL_CAP_MX As Long = 40
Private Const COL_CAP_MY As Long = 41
Private Const COL_CAP_EBC As Long = 42
Private Const COL_CAP_EBT As Long = 43
Private Const COL_CAP_ESC As Long = 44
Private Const COL_CAP_EST As Long = 45
Private Const COL_CAP_H As Long = 46
Private Const COL_CAP_X As Long = 47
Private Const COL_CAP_RESERVE As Long = 48
Private Const COL_CAP_STATUS As Long = 49

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Серый фон шапки, как в остальных подробных блоках.
Private Const RESULTS_COLUMN_WIDTH As Double = 10# ' На Results все столбцы держим одной ширины.
Private Const SMALL_VALUE As Double = 0.000000001 ' Численный порог только для пустых расчетных полей writer-а.
Private Const HEADER_NOTES_ROW As Long = 3
Private Const HEADER_LABEL_ROW As Long = 4

' Возвращает число строк шапки над rngStrengthSummaryAnchor.
Public Function HeaderRowsAboveAnchor() As Long
    HeaderRowsAboveAnchor = HEADER_ROWS
End Function

' Возвращает полную высоту блока: шапка плюс строки сочетаний.
Public Function RequiredRows(ByVal combinationCount As Long) As Long
    If combinationCount < 1 Then combinationCount = 1
    RequiredRows = HEADER_ROWS + combinationCount
End Function

' Возвращает число строк данных по высоте rngLoadCombinations, чтобы пустые
' доступные строки тоже получили рамки и не выглядели как остаток старого вывода.
Public Function RequiredDataRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredDataRowsForWorkbook = combinationCount
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1

    On Error GoTo SafeExit
    If workbook Is Nothing Then GoTo SafeExit
    Dim loads As Object
    Set loads = workbook.Names.Item("rngLoadCombinations").RefersToRange
    If loads.Rows.Count > 1 Then
        If loads.Rows.Count - 1 > RequiredDataRowsForWorkbook Then _
            RequiredDataRowsForWorkbook = loads.Rows.Count - 1
    End If
SafeExit:
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1
End Function

' Возвращает полную высоту блока для фактической книги.
Public Function RequiredRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredRowsForWorkbook = RequiredRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает ширину блока вместе с пустыми разделителями.
Public Function RequiredColumns() As Long
    RequiredColumns = TOTAL_COLUMNS
End Function

' Записывает подробную таблицу НДС/Capacity по модели прочности.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal section As CSectionModel = Nothing)
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 3996, "CStrengthSummaryWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3997, "CStrengthSummaryWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = TargetRange(workbook, RequiredDataRowsForWorkbook(workbook, batch.Count))
    PrepareTarget target
    WriteBlockTitle target, "Расчет НДС и несущей способности по модели прочности"
    WriteHeaders target, units
    WriteData target, batch, units, section
    FormatTable target, batch
    Exit Sub

Failed:
    Err.Raise Err.Number, "CStrengthSummaryWriter.WriteSummary", Err.Description
End Sub

' Очищает старый detailed-блок прочности вместе с объединениями и форматами.
Public Sub ClearSummary(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3998, "CStrengthSummaryWriter", "Книга Excel не передана."
    PrepareTarget ClearTargetRange(workbook)
End Sub

' Возвращает текущий диапазон таблицы: шапка над якорем и строки LC.
Private Function TargetRange(ByVal workbook As Object, Optional ByVal dataRows As Long = 1) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3999, "CStrengthSummaryWriter", _
        "rngStrengthSummaryAnchor должен иметь место для шапки над собой."
    Set TargetRange = anchor.Offset(-HEADER_ROWS, 0).Resize(RequiredRows(dataRows), TOTAL_COLUMNS)
End Function

' Возвращает область безопасной очистки до следующего detailed-блока Results.
Private Function ClearTargetRange(ByVal workbook As Object) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3999, "CStrengthSummaryWriter", _
        "rngStrengthSummaryAnchor должен иметь место для шапки над собой."

    Dim topRow As Long
    topRow = anchor.Row - HEADER_ROWS

    Dim rowsToClear As Long
    rowsToClear = RequiredRows(1)
    On Error GoTo UseDefaultHeight
    Dim crackWriter As CCrackSummaryWriter
    Set crackWriter = New CCrackSummaryWriter
    Dim crackAnchor As Object
    Set crackAnchor = workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Dim crackTopRow As Long
    crackTopRow = crackAnchor.Row - crackWriter.HeaderRowsAboveAnchor
    If crackTopRow > topRow Then rowsToClear = crackTopRow - topRow

UseDefaultHeight:
    If topRow + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then rowsToClear = anchor.Worksheet.Rows.Count - topRow + 1
    Set ClearTargetRange = anchor.Worksheet.Cells(topRow, anchor.Column).Resize(rowsToClear, TOTAL_COLUMNS)
End Function

' Снимает старые объединения и очищает значения/форматы перед новой записью.
Private Sub PrepareTarget(ByVal target As Object)
    ClearBlockTitleRow target
    target.UnMerge
    target.Clear
End Sub

' Записывает отдельную строку-название над detailed-блоком, не вмешиваясь в многоуровневую шапку таблицы.
Private Sub WriteBlockTitle(ByVal target As Object, ByVal titleText As String)
    If target.Row <= 1 Then Exit Sub

    Dim titleRow As Long
    titleRow = target.Row - 1
    target.Worksheet.Rows.Item(titleRow).Interior.Color = RGB(189, 215, 238)
    target.Worksheet.Rows.Item(titleRow).RowHeight = 15
    With target.Worksheet.Cells(titleRow, target.Column)
        .Value2 = titleText
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
End Sub

' Очищает строку-название вместе со старой заливкой, чтобы при переносе якоря не оставалось старых следов.
Private Sub ClearBlockTitleRow(ByVal target As Object)
    If target.Row <= 1 Then Exit Sub
    target.Worksheet.Rows.Item(target.Row - 1).Clear
    target.Worksheet.Rows.Item(target.Row - 1).RowHeight = 15
End Sub

' Формирует многоуровневую шапку таблицы.
Private Sub WriteHeaders(ByVal target As Object, ByVal units As CUnitSystem)
    Dim headers() As Variant
    ReDim headers(1 To HEADER_ROWS, 1 To TOTAL_COLUMNS)

    headers(1, COL_ID) = "Combination ID"
    headers(1, COL_RESULT_COMMENT) = "ResultComment"
    headers(1, COL_STATUS) = "общий статус"
    headers(1, COL_REF_N) = "относительно центральных осей приведенного сечения"
    headers(1, COL_LOAD_N) = "в точке приложения нагрузок"
    headers(1, COL_LIM_EBC) = "Общие данные"
    headers(1, COL_STATE_EPS0) = "НДС от заданного сочетания нагрузок (по модели прочности)"
    headers(1, COL_CAP_EPS0) = "Расчет предельной несущей способности по выбранной " & ChrW$(&H3BB) & "-траектории (по модели прочности)"

    headers(2, COL_STATE_EPS0) = "плоскость деформации"
    headers(2, COL_STATE_EBC) = "фактические макс. деформации"
    headers(2, COL_STATE_SBC) = "фактические напряжения"
    headers(2, COL_STATE_H) = "параметры сечения и статус"
    headers(2, COL_CAP_EPS0) = "плоскость деформации"
    headers(2, COL_CAP_LIMIT_STATE) = "параметры расчета"
    headers(2, COL_CAP_LAMBDA) = "предельные компоненты"
    headers(2, COL_CAP_EBC) = "фактические макс. деформации"
    headers(2, COL_CAP_H) = "параметры сечения при " & ChrW$(&H3BB) & "ult и статус"

    PutHeaderNotes headers
    PutColumnLabels headers, units
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).Value2 = headers
    MergeHeaders target
End Sub

' Заполняет строку обычных пояснений над нижним ярусом обозначений.
Private Sub PutHeaderNotes(ByRef headers() As Variant)
    headers(HEADER_NOTES_ROW, COL_RESULT_COMMENT) = "пояснение расчетного статуса блока прочности"
    headers(HEADER_NOTES_ROW, COL_REF_N) = "нагрузки, переданные в расчет прочности, в центре приведенного сечения"
    headers(HEADER_NOTES_ROW, COL_REF_MX) = "Mx там же, вокруг оси X принятой системы"
    headers(HEADER_NOTES_ROW, COL_REF_MY) = "My там же, вокруг оси Y принятой системы"
    headers(HEADER_NOTES_ROW, COL_LOAD_N) = "исходная N в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_LOAD_MX) = "исходный Mx в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_LOAD_MY) = "исходный My в пользовательской точке приложения"
    headers(HEADER_NOTES_ROW, COL_LIM_EBC) = "предельная деформация бетона при сжатии"
    headers(HEADER_NOTES_ROW, COL_LIM_EBT) = "предельная деформация бетона при растяжении; пусто, если растянутый бетон не работает"
    headers(HEADER_NOTES_ROW, COL_LIM_ESC) = "предельная деформация арматуры при сжатии"
    headers(HEADER_NOTES_ROW, COL_LIM_EST) = "предельная деформация арматуры при растяжении"
    headers(HEADER_NOTES_ROW, COL_STATE_EPS0) = ChrW$(&H3B5) & "0 в точке (0;0) внутренней системы координат модели; от этой точки заданы координаты всех элементов сечения"
    headers(HEADER_NOTES_ROW, COL_STATE_KX) = "кривизна kx при координате y во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_STATE_KY) = "кривизна ky при координате x во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_STATE_EBC) = "максимальная сжимающая деформация бетона, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_EBT) = "максимальная растягивающая деформация бетона, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_ESC) = "максимальная сжимающая деформация арматуры, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_EST) = "максимальная растягивающая деформация арматуры, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_SBC) = "максимальное сжатие бетона"
    headers(HEADER_NOTES_ROW, COL_STATE_SBT) = "максимальное растяжение бетона; пусто, если диаграмма не несет растяжение"
    headers(HEADER_NOTES_ROW, COL_STATE_SSC) = "максимальное сжатие арматуры"
    headers(HEADER_NOTES_ROW, COL_STATE_SST) = "максимальное растяжение арматуры"
    headers(HEADER_NOTES_ROW, COL_STATE_H) = "полная высота сечения по нормали к нейтральной линии"
    headers(HEADER_NOTES_ROW, COL_STATE_X) = "высота сжатой зоны по нормали к нейтральной линии"
    headers(HEADER_NOTES_ROW, COL_STATE_DIAGEXT) = "yes - равновесие найдено через продолжение диаграммы; no - продолжение было разрешено, но не понадобилось; - - продолжение выключено"
    headers(HEADER_NOTES_ROW, COL_STATE_RESERVE) = "минимальный запас по предельным деформациям"
    headers(HEADER_NOTES_ROW, COL_STATE_STATUS) = "статус фактического НДС по модели прочности"
    headers(HEADER_NOTES_ROW, COL_CAP_EPS0) = ChrW$(&H3B5) & "0 в точке (0;0) внутренней системы координат модели; от этой точки заданы координаты всех элементов сечения"
    headers(HEADER_NOTES_ROW, COL_CAP_KX) = "кривизна kx при координате y во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_CAP_KY) = "кривизна ky при координате x во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_CAP_LIMIT_STATE) = "причина найденного предельного состояния"
    headers(HEADER_NOTES_ROW, COL_CAP_LOAD_PATH) = "какие компоненты масштабировались"
    headers(HEADER_NOTES_ROW, COL_CAP_METHOD) = "фактический метод поиска"
    headers(HEADER_NOTES_ROW, COL_CAP_LAMBDA) = "предельный множитель выбранной траектории"
    headers(HEADER_NOTES_ROW, COL_CAP_N) = "полная N при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_MX) = "полный Mx при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_MY) = "полный My при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_EBC) = "сжатая деформация бетона в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_EBT) = "растянутая деформация бетона в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_ESC) = "сжатая деформация арматуры в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_EST) = "растянутая деформация арматуры в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_H) = "полная высота по нормали к нейтральной линии CapacityState"
    headers(HEADER_NOTES_ROW, COL_CAP_X) = "высота сжатой зоны CapacityState"
    headers(HEADER_NOTES_ROW, COL_CAP_RESERVE) = "запас по найденному множителю " & ChrW$(&H3BB)
    headers(HEADER_NOTES_ROW, COL_CAP_STATUS) = "статус расчета несущей способности"
End Sub

' Заполняет нижний ярус обозначений с единицами пользователя.
Private Sub PutColumnLabels(ByRef headers() As Variant, ByVal units As CUnitSystem)
    headers(HEADER_LABEL_ROW, COL_REF_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MY) = "My, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MY) = "My, " & OutputMomentUnit(units)

    headers(HEADER_LABEL_ROW, COL_LIM_EBC) = ChrW$(&H3B5) & "bc,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_EBT) = ChrW$(&H3B5) & "bt,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_ESC) = ChrW$(&H3B5) & "sc,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_EST) = ChrW$(&H3B5) & "st,ult"
    headers(HEADER_LABEL_ROW, COL_STATE_EPS0) = ChrW$(&H3B5) & "0"
    headers(HEADER_LABEL_ROW, COL_STATE_KX) = ChrW$(&H3BA) & "x, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_KY) = ChrW$(&H3BA) & "y, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_STATE_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_STATE_ESC) = ChrW$(&H3B5) & "sc"
    headers(HEADER_LABEL_ROW, COL_STATE_EST) = ChrW$(&H3B5) & "st"
    headers(HEADER_LABEL_ROW, COL_STATE_SBC) = ChrW$(&H3C3) & "bc"
    headers(HEADER_LABEL_ROW, COL_STATE_SBT) = ChrW$(&H3C3) & "bt"
    headers(HEADER_LABEL_ROW, COL_STATE_SSC) = ChrW$(&H3C3) & "sc"
    headers(HEADER_LABEL_ROW, COL_STATE_SST) = ChrW$(&H3C3) & "st"
    headers(HEADER_LABEL_ROW, COL_STATE_H) = "h, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_X) = "x, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_DIAGEXT) = "ExtUsed"
    headers(HEADER_LABEL_ROW, COL_STATE_RESERVE) = "запас (" & ChrW$(&H3B5) & ")"
    headers(HEADER_LABEL_ROW, COL_STATE_STATUS) = "статус"

    headers(HEADER_LABEL_ROW, COL_CAP_EPS0) = ChrW$(&H3B5) & "0"
    headers(HEADER_LABEL_ROW, COL_CAP_KX) = ChrW$(&H3BA) & "x, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_KY) = ChrW$(&H3BA) & "y, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_LIMIT_STATE) = "LimitState"
    headers(HEADER_LABEL_ROW, COL_CAP_LOAD_PATH) = "LoadPath"
    headers(HEADER_LABEL_ROW, COL_CAP_METHOD) = "Method"
    headers(HEADER_LABEL_ROW, COL_CAP_LAMBDA) = ChrW$(&H3BB) & "ult"
    headers(HEADER_LABEL_ROW, COL_CAP_N) = "Nult, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_MX) = "Mx,ult, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_MY) = "My,ult, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_CAP_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_CAP_ESC) = ChrW$(&H3B5) & "sc"
    headers(HEADER_LABEL_ROW, COL_CAP_EST) = ChrW$(&H3B5) & "st"
    headers(HEADER_LABEL_ROW, COL_CAP_H) = "h, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_X) = "x, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_RESERVE) = "запас (" & ChrW$(&H3BB) & ")"
    headers(HEADER_LABEL_ROW, COL_CAP_STATUS) = "статус"
End Sub

' Объединяет прямоугольные участки шапки.
Private Sub MergeHeaders(ByVal target As Object)
    MergeRange target, 1, COL_ID, HEADER_LABEL_ROW, COL_ID
    MergeRange target, 1, COL_RESULT_COMMENT, HEADER_LABEL_ROW, COL_RESULT_COMMENT
    MergeRange target, 1, COL_STATUS, HEADER_LABEL_ROW, COL_STATUS
    MergeRange target, 1, COL_REF_N, 2, COL_REF_MY
    MergeRange target, 1, COL_LOAD_N, 2, COL_LOAD_MY
    MergeRange target, 1, COL_LIM_EBC, 2, COL_LIM_EST
    MergeRange target, 1, COL_STATE_EPS0, 1, COL_STATE_STATUS
    MergeRange target, 2, COL_STATE_EPS0, 2, COL_STATE_KY
    MergeRange target, 2, COL_STATE_EBC, 2, COL_STATE_EST
    MergeRange target, 2, COL_STATE_SBC, 2, COL_STATE_SST
    MergeRange target, 2, COL_STATE_H, 2, COL_STATE_STATUS
    MergeRange target, 1, COL_CAP_EPS0, 1, COL_CAP_STATUS
    MergeRange target, 2, COL_CAP_EPS0, 2, COL_CAP_KY
    MergeRange target, 2, COL_CAP_LIMIT_STATE, 2, COL_CAP_METHOD
    MergeRange target, 2, COL_CAP_LAMBDA, 2, COL_CAP_MY
    MergeRange target, 2, COL_CAP_EBC, 2, COL_CAP_EST
    MergeRange target, 2, COL_CAP_H, 2, COL_CAP_STATUS
End Sub

' Записывает строки сочетаний одним массивом.
Private Sub WriteData(ByVal target As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows <= 0 Then Exit Sub

    Dim data() As Variant
    ReDim data(1 To dataRows, 1 To TOTAL_COLUMNS)

    Dim i As Long
    For i = 1 To batch.Count
        PutCombination data, OutputDataOffset(batch, i, dataRows), batch, i, units, section
    Next i
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).Value2 = data
End Sub

' Заполняет одну строку detailed-блока прочности.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim policy As CBatchStatusPolicy
    Set policy = StatusPolicy()

    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_RESULT_COMMENT) = MetaComment(batch.StrengthMeta(index))
    data(rowIndex, COL_STATUS) = StrengthOverallStatus(batch, index)
    data(rowIndex, COL_STATE_STATUS) = policy.NotApplicable
    data(rowIndex, COL_CAP_STATUS) = policy.NotApplicable

    If Not batch.StrengthDirectStateRequested(index) And Not batch.StrengthCapacityRequested(index) Then Exit Sub

    data(rowIndex, COL_REF_N) = OutputForce(batch.StrengthReferenceN(index), units)
    data(rowIndex, COL_REF_MX) = OutputMx(batch.StrengthReferenceMx(index), units)
    data(rowIndex, COL_REF_MY) = OutputMy(batch.StrengthReferenceMy(index), units)
    data(rowIndex, COL_LOAD_N) = OutputForce(batch.StrengthLoadPointN(index), units)
    data(rowIndex, COL_LOAD_MX) = OutputMx(batch.StrengthLoadPointMx(index), units)
    data(rowIndex, COL_LOAD_MY) = OutputMy(batch.StrengthLoadPointMy(index), units)

    PutStrengthLimits data, rowIndex, batch, index, units
    PutStrengthState data, rowIndex, batch, index, units, section, policy
    PutCapacityState data, rowIndex, batch, index, units, section, policy
End Sub

' Пишет предельные деформации модели прочности.
' Возвращает позицию LC внутри data-части подробной таблицы.
' Основной поток читает offset из rngLoadCombinations; прямые тесты без offset продолжают выводиться компактно.
Private Function OutputDataOffset(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal dataRows As Long) As Long
    OutputDataOffset = batch.SourceDataOffset(index)
    If OutputDataOffset < 1 Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = dataRows
End Function

Private Sub PutStrengthLimits(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    data(rowIndex, COL_LIM_EBC) = batch.StrengthConcreteCompressionLimitForCombination(index)
    If batch.StrengthConcreteTensionLimitEnabledForCombination(index) Then _
        data(rowIndex, COL_LIM_EBT) = batch.StrengthConcreteTensionLimitForCombination(index)
    data(rowIndex, COL_LIM_ESC) = batch.StrengthSteelCompressionLimitForCombination(index)
    data(rowIndex, COL_LIM_EST) = batch.StrengthSteelTensionLimitForCombination(index)
End Sub

' Пишет фактический StrengthState, если он действительно был рассчитан.
Private Sub PutStrengthState(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel, ByVal policy As CBatchStatusPolicy)
    If Not batch.StrengthDirectStateRequested(index) Then
        data(rowIndex, COL_STATE_STATUS) = policy.NotApplicable
        Exit Sub
    End If

    Dim stateResult As CSectionStateResult
    Set stateResult = batch.FindNamedState(index, sstStrengthState)
    If stateResult Is Nothing Then
        data(rowIndex, COL_STATE_STATUS) = DirectStatusText(batch, index, policy)
        Exit Sub
    End If

    PutStatePlane data, rowIndex, stateResult, units, COL_STATE_EPS0, COL_STATE_KX, COL_STATE_KY
    PutStateExtremes data, rowIndex, stateResult, units, _
        COL_STATE_EBC, COL_STATE_EBT, COL_STATE_ESC, COL_STATE_EST, _
        COL_STATE_SBC, COL_STATE_SBT, COL_STATE_SSC, COL_STATE_SST
    PutStateGeometry data, rowIndex, stateResult, units, section, COL_STATE_H, COL_STATE_X
    data(rowIndex, COL_STATE_DIAGEXT) = ExtensionUsedText(batch, index, stateResult)
    data(rowIndex, COL_STATE_RESERVE) = StrainReserve(batch, index, stateResult)
    data(rowIndex, COL_STATE_STATUS) = policy.DisplayStatus(stateResult.Status)
End Sub

' Пишет CapacityState и предельные компоненты найденной точки.
' Показывает использование продолжения диаграммы только тогда, когда само
' продолжение разрешено настройкой Solver.DirectState.DiagramExtension.
' При выключенной настройке выводится "-", потому что extension не мог
' применяться и обычная пара yes/no здесь не несет полезного смысла.
Private Function ExtensionUsedText(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal stateResult As CSectionStateResult) As String
    If Not batch.DirectStateDiagramExtensionEnabled(index) Then
        ExtensionUsedText = "-"
    ElseIf stateResult.ExtensionUsed Then
        ExtensionUsedText = "yes"
    Else
        ExtensionUsedText = "no"
    End If
End Function

Private Sub PutCapacityState(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel, ByVal policy As CBatchStatusPolicy)
    If Not batch.StrengthCapacityRequested(index) Then
        data(rowIndex, COL_CAP_STATUS) = policy.NotApplicable
        Exit Sub
    End If

    data(rowIndex, COL_CAP_LIMIT_STATE) = batch.CapacityLimitState(index)
    data(rowIndex, COL_CAP_LOAD_PATH) = batch.CapacityLoadPath(index)
    data(rowIndex, COL_CAP_METHOD) = batch.CapacitySolutionMethod(index)
    If batch.LambdaCapacity(index) > SMALL_VALUE Then data(rowIndex, COL_CAP_LAMBDA) = batch.LambdaCapacity(index)

    Dim stateResult As CSectionStateResult
    Set stateResult = batch.FindNamedState(index, sstCapacityState)
    If Not stateResult Is Nothing Then
        PutStatePlane data, rowIndex, stateResult, units, COL_CAP_EPS0, COL_CAP_KX, COL_CAP_KY
        PutCapacityLoadPointComponents data, rowIndex, batch, index, units
        PutStateExtremes data, rowIndex, stateResult, units, _
            COL_CAP_EBC, COL_CAP_EBT, COL_CAP_ESC, COL_CAP_EST, _
            0, 0, 0, 0
        PutStateGeometry data, rowIndex, stateResult, units, section, COL_CAP_H, COL_CAP_X
    End If
    If batch.StrengthCapacityReserve(index) > SMALL_VALUE Then _
        data(rowIndex, COL_CAP_RESERVE) = batch.StrengthCapacityReserve(index)
    data(rowIndex, COL_CAP_STATUS) = CapacityStatusText(batch, index, policy)
End Sub

' Выводит полный предельный вектор CapacityState в точке приложения нагрузки.
Private Sub PutCapacityLoadPointComponents(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    data(rowIndex, COL_CAP_N) = OutputForce(batch.NUltimate(index), units)
    data(rowIndex, COL_CAP_MX) = OutputMx(batch.MxUltimate(index), units)
    data(rowIndex, COL_CAP_MY) = OutputMy(batch.MyUltimate(index), units)
End Sub

' Пишет параметры плоскости деформаций saved-state: epsilon0 в начале
' внутренней системы координат модели и две кривизны в OUTPUT-единицах.
Private Sub PutStatePlane(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal colEps0 As Long, ByVal colKx As Long, ByVal colKy As Long)
    data(rowIndex, colEps0) = stateResult.Epsilon0
    data(rowIndex, colKx) = OutputCurvature(stateResult.KappaX, units)
    data(rowIndex, colKy) = OutputCurvature(stateResult.KappaY, units)
End Sub

' Пишет экстремумы состояния по флагам наличия физической зоны.
Private Sub PutStateExtremes(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal colEbc As Long, ByVal colEbt As Long, ByVal colEsc As Long, ByVal colEst As Long, _
        ByVal colSbc As Long, ByVal colSbt As Long, ByVal colSsc As Long, ByVal colSst As Long)
    If stateResult.HasConcreteCompression Then data(rowIndex, colEbc) = stateResult.MinConcreteStrain
    If stateResult.HasConcreteTension Then data(rowIndex, colEbt) = stateResult.MaxConcreteStrain
    If stateResult.HasSteelCompression Then data(rowIndex, colEsc) = stateResult.MinSteelStrain
    If stateResult.HasSteelTension Then data(rowIndex, colEst) = stateResult.MaxSteelStrain

    If colSbc > 0 And stateResult.HasConcreteCompression Then data(rowIndex, colSbc) = OutputStress(stateResult.MinConcreteStress, units)
    If colSbt > 0 And stateResult.HasConcreteTension Then data(rowIndex, colSbt) = OutputStress(stateResult.MaxConcreteStress, units)
    If colSsc > 0 And stateResult.HasSteelCompression Then data(rowIndex, colSsc) = OutputStress(stateResult.MinSteelStress, units)
    If colSst > 0 And stateResult.HasSteelTension Then data(rowIndex, colSst) = OutputStress(stateResult.MaxSteelStress, units)
End Sub

' Пишет h и x для плоскости деформаций named state.
Private Sub PutStateGeometry(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal section As CSectionModel, ByVal colH As Long, ByVal colX As Long)
    If section Is Nothing Then Exit Sub

    Dim depthH As Double
    Dim compressedDepth As Double
    CalculateDepthsPerpendicularToNeutral section, stateResult.Epsilon0, stateResult.KappaX, stateResult.KappaY, _
        depthH, compressedDepth
    If depthH > 0# Then data(rowIndex, colH) = OutputLength(depthH, units)
    If compressedDepth >= 0# Then data(rowIndex, colX) = OutputLength(compressedDepth, units)
End Sub

' Рассчитывает минимальный запас по доступным предельным деформациям.
Private Function StrainReserve(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal stateResult As CSectionStateResult) As Variant
    Dim reserve As Double
    reserve = 1E+100

    AddReserve reserve, stateResult.HasConcreteCompression, _
        batch.StrengthConcreteCompressionLimitForCombination(index), stateResult.MinConcreteStrain
    If batch.StrengthConcreteTensionLimitEnabledForCombination(index) Then _
        AddReserve reserve, stateResult.HasConcreteTension, _
            batch.StrengthConcreteTensionLimitForCombination(index), stateResult.MaxConcreteStrain
    AddReserve reserve, stateResult.HasSteelCompression, _
        batch.StrengthSteelCompressionLimitForCombination(index), stateResult.MinSteelStrain
    AddReserve reserve, stateResult.HasSteelTension, _
        batch.StrengthSteelTensionLimitForCombination(index), stateResult.MaxSteelStrain

    If reserve < 1E+90 Then StrainReserve = reserve
End Function

' Добавляет один частный запас, если зона и предельная деформация существуют.
Private Sub AddReserve(ByRef reserve As Double, ByVal hasValue As Boolean, ByVal limitValue As Double, ByVal actualValue As Double)
    If Not hasValue Then Exit Sub
    If Abs(limitValue) <= SMALL_VALUE Or Abs(actualValue) <= SMALL_VALUE Then Exit Sub
    reserve = MinDouble(reserve, Abs(limitValue) / Abs(actualValue))
End Sub

' Оформляет шапку, рамки, число и подсветку статусов.
Private Sub FormatTable(ByVal target As Object, ByVal batch As CBatchSectionCalculator)
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.HorizontalAlignment = -4108
    target.VerticalAlignment = -4108
    target.WrapText = False
    target.Rows.RowHeight = 15

    SetColumnWidths target
    FormatBlock target, COL_LEFT_FIRST, COL_LEFT_LAST
    FormatBlock target, COL_STATE_FIRST, COL_STATE_LAST
    FormatBlock target, COL_CAPACITY_FIRST, COL_CAPACITY_LAST
    target.Columns(COL_SEPARATOR1).Clear
    target.Columns(COL_SEPARATOR2).Clear

    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).HorizontalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).VerticalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).WrapText = True
    target.Rows(HEADER_NOTES_ROW).Font.Bold = False
    target.Rows(HEADER_NOTES_ROW).HorizontalAlignment = -4131
    target.Rows(HEADER_NOTES_ROW).WrapText = False
    target.Rows(HEADER_LABEL_ROW).WrapText = False

    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows > 0 Then
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).HorizontalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).VerticalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).NumberFormat = "0.000"
        target.Offset(HEADER_ROWS, COL_ID - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_RESULT_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
        target.Offset(HEADER_ROWS, COL_STATE_DIAGEXT - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_STATE_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_CAP_LIMIT_STATE - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_CAP_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        ApplyStrainNumberFormat target, dataRows
        ApplyStatusCells target, batch, dataRows
    End If
    ApplyHeaderSubscripts target
End Sub

' Форматирует относительные деформации и кривизны плоскости НДС с нужной точностью вывода.
Private Sub ApplyStrainNumberFormat(ByVal target As Object, ByVal dataRows As Long)
    target.Offset(HEADER_ROWS, COL_LIM_EBC - 1).Resize(dataRows, COL_LIM_EST - COL_LIM_EBC + 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_STATE_EPS0 - 1).Resize(dataRows, 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_STATE_KX - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
    target.Offset(HEADER_ROWS, COL_STATE_EBC - 1).Resize(dataRows, COL_STATE_EST - COL_STATE_EBC + 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_CAP_EPS0 - 1).Resize(dataRows, 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_CAP_KX - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
    target.Offset(HEADER_ROWS, COL_CAP_EBC - 1).Resize(dataRows, COL_CAP_EST - COL_CAP_EBC + 1).NumberFormat = "0.0000"
End Sub

' Задает одинаковую стандартную ширину всем столбцам Results, затронутым блоком.
Private Sub SetColumnWidths(ByVal target As Object)
    Dim columnIndex As Long
    For columnIndex = 1 To TOTAL_COLUMNS
        target.Worksheet.Columns(target.Column + columnIndex - 1).ColumnWidth = RESULTS_COLUMN_WIDTH
    Next columnIndex
End Sub

' Оформляет непрерывный блок таблицы, не затрагивая пустые разделители.
Private Sub FormatBlock(ByVal target As Object, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim block As Object
    Set block = target.Offset(0, firstCol - 1).Resize(target.Rows.Count, lastCol - firstCol + 1)
    block.Borders.LineStyle = 1
    block.Borders.Weight = 2
    block.Rows(1).Resize(HEADER_ROWS).Font.Bold = True
    target.Offset(0, firstCol - 1).Resize(HEADER_ROWS, lastCol - firstCol + 1).Interior.Color = HEADER_COLOR
End Sub

' Красит строковые блоки подробной таблицы по фактическому статусу каждой проверки.
Private Sub ApplyStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal dataRows As Long)
    If batch Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To batch.Count
        Dim rowOffset As Long
        rowOffset = OutputDataOffset(batch, i, dataRows)
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_STATE_STATUS
        ApplyStatusCell target, HEADER_ROWS + rowOffset, COL_CAP_STATUS
    Next i
End Sub

' Красит одну статусную ячейку detailed-блока, не затрагивая соседние расчетные значения строки.
Private Sub ApplyStatusCell(ByVal target As Object, ByVal rowIndex As Long, ByVal columnIndex As Long)
    Dim cell As Object
    Set cell = target.Cells(rowIndex, columnIndex)
    ApplyStatusFill cell, CStr(cell.Value2)
End Sub

' Проставляет нижние индексы в обозначениях нижнего яруса шапки.
Private Sub ApplyHeaderSubscripts(ByVal target As Object)
    Dim labels As Object
    Set labels = target.Rows(HEADER_LABEL_ROW)

    Dim cols As Variant
    cols = Array(COL_LIM_EBC, COL_LIM_EBT, COL_LIM_ESC, COL_LIM_EST, _
        COL_STATE_EPS0, COL_STATE_KX, COL_STATE_KY, _
        COL_STATE_EBC, COL_STATE_EBT, COL_STATE_ESC, COL_STATE_EST, _
        COL_STATE_SBC, COL_STATE_SBT, COL_STATE_SSC, COL_STATE_SST, _
        COL_CAP_LAMBDA, COL_CAP_N, COL_CAP_MX, COL_CAP_MY, _
        COL_CAP_EPS0, COL_CAP_KX, COL_CAP_KY, _
        COL_CAP_EBC, COL_CAP_EBT, COL_CAP_ESC, COL_CAP_EST)

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ApplySubscriptToken labels.Cells(1, CLng(cols(i))), "bc"
        ApplySubscriptToken labels.Cells(1, CLng(cols(i))), "bt"
        ApplySubscriptToken labels.Cells(1, CLng(cols(i))), "sc"
        ApplySubscriptToken labels.Cells(1, CLng(cols(i))), "st"
        ApplySubscriptToken labels.Cells(1, CLng(cols(i))), "ult"
    Next i
    ApplySubscriptToken labels.Cells(1, COL_CAP_LAMBDA), "ult"
    ApplySubscriptToken labels.Cells(1, COL_STATE_EPS0), "0"
    ApplySubscriptToken labels.Cells(1, COL_STATE_KX), "x"
    ApplySubscriptToken labels.Cells(1, COL_STATE_KY), "y"
    ApplySubscriptToken labels.Cells(1, COL_CAP_EPS0), "0"
    ApplySubscriptToken labels.Cells(1, COL_CAP_KX), "x"
    ApplySubscriptToken labels.Cells(1, COL_CAP_KY), "y"
End Sub

' Ставит нижний индекс найденному текстовому токену внутри одной ячейки.
Private Sub ApplySubscriptToken(ByVal cell As Object, ByVal token As String)
    Dim textValue As String
    textValue = CStr(cell.Value2)
    Dim startPos As Long
    startPos = InStr(1, textValue, token, vbTextCompare)
    Do While startPos > 0
        cell.Characters(startPos, Len(token)).Font.Subscript = True
        startPos = InStr(startPos + Len(token), textValue, token, vbTextCompare)
    Loop
End Sub

' Возвращает статус StrengthState для отдельного блока.
Private Function StrengthStateStatus(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As String
    If Not batch.StrengthDirectStateRequested(index) Then
        StrengthStateStatus = StatusPolicy().NotApplicable
        Exit Function
    End If

    Dim stateResult As CSectionStateResult
    Set stateResult = batch.FindNamedState(index, sstStrengthState)
    If Not stateResult Is Nothing Then
        StrengthStateStatus = StatusPolicy().DisplayStatus(stateResult.Status)
    Else
        StrengthStateStatus = DirectStatusText(batch, index, StatusPolicy())
    End If
End Function

' Возвращает общий статус именно detailed-блока прочности.
Private Function StrengthOverallStatus(ByVal batch As CBatchSectionCalculator, ByVal index As Long) As String
    Dim directStatus As String
    Dim capacityStatus As String
    directStatus = StatusPolicy().NotApplicable
    capacityStatus = StatusPolicy().NotApplicable

    If batch.StrengthDirectStateRequested(index) Then directStatus = StrengthStateStatus(batch, index)
    If batch.StrengthCapacityRequested(index) Then capacityStatus = CapacityStatusText(batch, index, StatusPolicy())
    StrengthOverallStatus = StatusPolicy().Aggregate(directStatus, capacityStatus, StatusPolicy().NotApplicable)
End Function

' Возвращает внешний статус прямого НДС по ResultMeta. Строковый статус остается
' только переходным fallback для старых тестовых объектов без result-tree.
Private Function DirectStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CBatchStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.DirectStateMeta(index)
    If Not meta Is Nothing Then
        DirectStatusText = policy.StatusFromMeta(meta)
    Else
        DirectStatusText = policy.DisplayStatus(batch.DirectStateStatus(index))
    End If
End Function

' Возвращает внешний статус несущей способности по ResultMeta.
Private Function CapacityStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CBatchStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.CapacityMeta(index)
    If Not meta Is Nothing Then
        CapacityStatusText = policy.StatusFromMeta(meta)
    Else
        CapacityStatusText = policy.DisplayStatus(batch.CapacityStatus(index))
    End If
End Function

' Возвращает готовое пользовательское пояснение из ResultMeta блока.
' Writer не анализирует технические StopReason и не собирает комментарий заново.
Private Function MetaComment(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then Exit Function
    MetaComment = meta.ResultComment
End Function

' Объединяет прямоугольный участок шапки.
Private Sub MergeRange(ByVal target As Object, ByVal firstRow As Long, ByVal firstCol As Long, _
        ByVal lastRow As Long, ByVal lastCol As Long)
    If lastRow < firstRow Or lastCol < firstCol Then Exit Sub
    target.Worksheet.Range(target.Cells(firstRow, firstCol), target.Cells(lastRow, lastCol)).Merge
End Sub

' Определяет габарит h и глубину сжатой зоны x вдоль нормали к нейтральной линии.
Private Sub CalculateDepthsPerpendicularToNeutral(ByVal section As CSectionModel, ByVal eps0 As Double, _
        ByVal kappaX As Double, ByVal kappaY As Double, ByRef depthH As Double, ByRef compressedDepth As Double)
    Dim normValue As Double
    normValue = Sqr(kappaX * kappaX + kappaY * kappaY)
    If normValue <= 0.000000000000001 Then Exit Sub

    Dim nx As Double
    Dim ny As Double
    nx = kappaY / normValue
    ny = kappaX / normValue

    Dim minS As Double
    Dim maxS As Double
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, nx, ny, False, minS, maxS

    depthH = maxS - minS
    If depthH <= 0# Then Exit Sub

    Dim neutralS As Double
    neutralS = -eps0 / normValue

    If eps0 + normValue * minS < 0# And eps0 + normValue * maxS < 0# Then
        compressedDepth = depthH
    ElseIf eps0 + normValue * minS >= 0# And eps0 + normValue * maxS >= 0# Then
        compressedDepth = 0#
    Else
        compressedDepth = MaxDouble(0#, MinDouble(maxS, neutralS) - minS)
    End If
End Sub

' Ниже собраны адаптеры единиц для strength-блока, чтобы все значения в таблице
' выводились теми же единицами и знаками, что и остальная программа.
Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

' Возвращает подпись выходной единицы длины для заголовков strength-таблицы.
Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

' Переводит напряжение из внутренних МПа в выбранную пользователем единицу.
Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

' Переводит кривизну плоскости деформаций из внутренних 1/мм
' в пользовательскую единицу вывода.
Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value Else OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

' Возвращает подпись выходной единицы кривизны для kx/ky.
Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

' Переводит продольную силу из внутренних Н в пользовательскую систему вывода.
Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Возвращает подпись выходной единицы силы для заголовков strength-таблицы.
Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Переводит Mx с учетом пользовательской системы знаков.
Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

' Переводит My с учетом пользовательской системы знаков.
Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

' Возвращает подпись выходной единицы момента для заголовков strength-таблицы.
Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

' Возвращает единый словарь пользовательских статусов для подробного блока прочности.
Private Function StatusPolicy() As CBatchStatusPolicy
    Static policy As CBatchStatusPolicy
    If policy Is Nothing Then Set policy = New CBatchStatusPolicy
    Set StatusPolicy = policy
End Function

' Локальный максимум нужен для компактных геометрических расчетов внутри writer-а.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxDouble = a Else MaxDouble = b
End Function

' Локальный минимум нужен для компактных геометрических расчетов внутри writer-а.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then MinDouble = a Else MinDouble = b
End Function
