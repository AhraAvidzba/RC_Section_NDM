VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemSettingsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение настроек Config и параметров материалов
' ==========================================================================
' Reader проходит по именованным диапазонам Config, собирает пары ключ/значение
' и разворачивает специальные таблицы материалов в те же ключи словаря.
' Материальные диаграммы больше не читаются как пользовательский набор точек:
' точки строит CMaterialModelProvider по параметрам бетона/арматуры и выбранному
' расчетному режиму.
'
' Класс не делает инженерных расчетов. Его задача - надежно прочитать ввод и
' дать понятную ошибку, если обязательный ключ или таблица отсутствуют.

Private mKeys() As String ' Имена настроек Config: Solver.Method, Capacity.SolutionStrategy и т.п.
Private mValues() As String ' Значения, которые ввел пользователь или подставили формулы таблицы.
Private mDefaults() As String ' Значения по умолчанию из таблицы Config.
Private mSources() As String ' Комментарий/источник строки для диагностики.
Private mCount As Long ' Количество загруженных ключей.

' Загружает все настройки книги, нужные расчету и выводу.
' Метод объединяет разные именованные диапазоны Config в один внутренний
' словарь. Новые таблицы материалов имеют человекочитаемую форму, поэтому
' перед попаданием в словарь разворачиваются в ключи Concrete.* / Steel.*.
Public Sub LoadFromWorkbook(ByVal workbook As Object)
    Clear
    AppendSettingsRange workbook.Names.Item("rngUnitSettings").RefersToRange
    AppendSettingsRange workbook.Names.Item("rngSignConventionSettings").RefersToRange
    AppendSettingsRange workbook.Names.Item("rngSystemSettings").RefersToRange
    AppendNamedSettingsRangeIfExists workbook, "rngConcreteMaterialParameters"
    AppendNamedSettingsRangeIfExists workbook, "rngSteelMaterialParameters"
    AppendSettingsRange workbook.Names.Item("rngPlotAnnotationSettings").RefersToRange
    AppendGeometrySettingsRanges workbook
End Sub

' Подключает дополнительный именованный диапазон, если он уже есть в книге.
' Это позволяет старым тестовым книгам открываться с расчетными defaults,
' но новая сборка книги все равно создает эти диапазоны явно на Config.
Private Sub AppendNamedSettingsRangeIfExists(ByVal workbook As Object, ByVal rangeName As String)
    On Error GoTo MissingRange
    Dim settingsRange As Object
    Set settingsRange = workbook.Names.Item(rangeName).RefersToRange
    On Error GoTo 0
    AppendSettingsRange settingsRange
    Exit Sub
MissingRange:
    On Error GoTo 0
End Sub

' Находит все входные geometry-диапазоны Config и добавляет их настройки в общий словарь.
Private Sub AppendGeometrySettingsRanges(ByVal workbook As Object)
    Dim nameObject As Object
    ' Проходим по именованным диапазонам и выбираем только входные таблицы
    ' геометрии. rngNDMSectionGeometry пропускается: это Results, а не Config.
    For Each nameObject In workbook.Names
        Dim rangeName As String
        rangeName = LocalName(CStr(nameObject.Name))
        If IsInputGeometrySettingsRangeName(rangeName) Then
            On Error Resume Next
            Dim settingsRange As Object
            Set settingsRange = nameObject.RefersToRange
            Dim rangeError As Long
            rangeError = Err.Number
            On Error GoTo 0
            If rangeError = 0 And Not settingsRange Is Nothing Then AppendSettingsRange settingsRange
        End If
    Next nameObject
End Sub

' Убирает префикс имени книги/листа у именованного диапазона, оставляя локальное имя.
Private Function LocalName(ByVal workbookName As String) As String
    Dim bangPos As Long
    bangPos = InStrRev(workbookName, "!")
    If bangPos > 0 Then
        LocalName = Mid$(workbookName, bangPos + 1)
    Else
        LocalName = workbookName
    End If
End Function

' Определяет, является ли имя входной таблицей геометрии.
Private Function IsInputGeometrySettingsRangeName(ByVal rangeName As String) As Boolean
    Dim key As String
    key = LCase$(Trim$(rangeName))
    If key = "rngndmsectiongeometry" Then Exit Function
    IsInputGeometrySettingsRangeName = (Left$(key, 3) = "rng" And Right$(key, 8) = "geometry")
End Function

' Загружает настройки из одного диапазона.
' Этот короткий путь используют тесты и вспомогательные сценарии без полной книги.
Public Sub LoadFromRange(ByVal settingsRange As Object)
    Clear
    AppendSettingsRange settingsRange
End Sub

' Добавляет один диапазон настроек; специальные геометрические таблицы сначала распознаются по шапке.
Private Sub AppendSettingsRange(ByVal settingsRange As Object)
    Dim data As Variant
    data = settingsRange.Value2
    Dim columnCount As Long
    columnCount = UBound(data, 2)

    If IsHollowRectangleSettingsRange(data) Then
        AppendHollowRectangleSettings data
        Exit Sub
    End If
    If IsRoundedRectangleFaceSettingsRange(data) Then
        AppendRoundedRectangleFaceSettings data
        Exit Sub
    End If
    If IsRectSetFaceSettingsRange(data) Then
        AppendRectSetFaceSettings data
        Exit Sub
    End If
    If IsUnitSettingsRange(data) Then
        AppendUnitSettings data
        Exit Sub
    End If
    If IsSignConventionSettingsRange(data) Then
        AppendSignConventionSettings data
        Exit Sub
    End If
    If IsPlotAnnotationSettingsRange(data) Then
        AppendPlotAnnotationSettings data
        Exit Sub
    End If
    If IsConcreteMaterialParametersRange(data) Then
        AppendConcreteMaterialParameters data
        Exit Sub
    End If
    If IsSteelMaterialParametersRange(data) Then
        AppendSteelMaterialParameters data
        Exit Sub
    End If
    Dim rowIndex As Long
    ' Обычные таблицы читаются как ключ/значение. Строки с заголовками секций
    ' в квадратных скобках пропускаются.
    For rowIndex = 2 To UBound(data, 1)
        Dim keyValue As String
        keyValue = Trim$(CStr(data(rowIndex, 1)))
        If Len(keyValue) > 0 And Left$(keyValue, 1) <> "[" Then
            Dim sourceValue As String
            If columnCount >= 6 Then
                sourceValue = Trim$(CStr(data(rowIndex, 6)))
            ElseIf columnCount >= 4 Then
                sourceValue = Trim$(CStr(data(rowIndex, 4)))
            Else
                sourceValue = vbNullString
            End If
            AppendSetting keyValue, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 3))), sourceValue
        End If
    Next rowIndex
End Sub

' Распознает компактную таблицу параметров бетона:
' Параметр | Сжатие | Растяжение | Ед. | Комментарий | Справка.
Private Function IsConcreteMaterialParametersRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 4 Then Exit Function
    IsConcreteMaterialParametersRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Параметр", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "Сжатие", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "Растяжение", vbTextCompare) = 0 And _
        InStr(1, CStr(data(2, 1)), "Concrete.", vbTextCompare) > 0
End Function

' Разворачивает симметричную таблицу бетона в конкретные ключи Rb/Rbt/Eb/Ebt.
Private Sub AppendConcreteMaterialParameters(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim rowKey As String
        rowKey = Trim$(CStr(data(rowIndex, 1)))
        If Len(rowKey) = 0 Then GoTo NextRow

        Dim compressionValue As String
        Dim tensionValue As String
        Dim commentValue As String
        compressionValue = Trim$(CStr(data(rowIndex, 2)))
        tensionValue = Trim$(CStr(data(rowIndex, 3)))
        If UBound(data, 2) >= 5 Then commentValue = Trim$(CStr(data(rowIndex, 5)))

        Select Case LCase$(rowKey)
            Case "concrete.r.uls(i)", "concrete.r.i(uls)", "concrete.r.uls"
                AppendSetting "Concrete.Rb.ULS", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Rbt.ULS", tensionValue, tensionValue, commentValue
            Case "concrete.r.sls(ii)", "concrete.r.ii(sls)", "concrete.r.sls"
                AppendSetting "Concrete.Rb.SLS", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Rbt.SLS", tensionValue, tensionValue, commentValue
            Case "concrete.rb.mc2"
                AppendSetting "Concrete.Rb.mc2", compressionValue, compressionValue, commentValue
            Case "concrete.e"
                AppendSetting "Concrete.Eb", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Ebt", tensionValue, tensionValue, commentValue
            Case "concrete.twoline.eb1red"
                AppendSetting "Concrete.Eb1Red", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Ebt1Red", tensionValue, tensionValue, commentValue
            Case "concrete.threeline.eb0"
                AppendSetting "Concrete.Eb0", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Ebt0", tensionValue, tensionValue, commentValue
            Case "concrete.twothreeline.eb2"
                AppendSetting "Concrete.Eb2", compressionValue, compressionValue, commentValue
                AppendSetting "Concrete.Ebt2", tensionValue, tensionValue, commentValue
            Case Else
                If Len(compressionValue) > 0 Then AppendSetting rowKey, compressionValue, compressionValue, commentValue
        End Select
NextRow:
    Next rowIndex
End Sub

' Распознает таблицу параметров арматуры с аналогичной структурой.
Private Function IsSteelMaterialParametersRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 4 Then Exit Function
    IsSteelMaterialParametersRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Параметр", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "Сжатие", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "Растяжение", vbTextCompare) = 0 And _
        InStr(1, CStr(data(2, 1)), "Steel.", vbTextCompare) > 0
End Function

' Разворачивает симметричную таблицу арматуры в конкретные ключи Rsc/Rs/Esc/Es.
Private Sub AppendSteelMaterialParameters(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim rowKey As String
        rowKey = Trim$(CStr(data(rowIndex, 1)))
        If Len(rowKey) = 0 Then GoTo NextRow

        Dim compressionValue As String
        Dim tensionValue As String
        Dim commentValue As String
        compressionValue = Trim$(CStr(data(rowIndex, 2)))
        tensionValue = Trim$(CStr(data(rowIndex, 3)))
        If UBound(data, 2) >= 5 Then commentValue = Trim$(CStr(data(rowIndex, 5)))

        Select Case LCase$(rowKey)
            Case "steel.r.uls(i)", "steel.r.i(uls)", "steel.r.uls"
                AppendSetting "Steel.Rsc.ULS", compressionValue, compressionValue, commentValue
                AppendSetting "Steel.Rs.ULS", tensionValue, tensionValue, commentValue
            Case "steel.r.sls(ii)", "steel.r.ii(sls)", "steel.r.sls"
                AppendSetting "Steel.Rsc.SLS", compressionValue, compressionValue, commentValue
                AppendSetting "Steel.Rs.SLS", tensionValue, tensionValue, commentValue
            Case "steel.e"
                AppendSetting "Steel.Esc", compressionValue, compressionValue, commentValue
                AppendSetting "Steel.Es", tensionValue, tensionValue, commentValue
            Case "steel.twoline.es2"
                AppendSetting "Steel.TwoLine.Esc2", compressionValue, compressionValue, commentValue
                AppendSetting "Steel.TwoLine.Es2", tensionValue, tensionValue, commentValue
            Case "steel.threeline.es2"
                AppendSetting "Steel.ThreeLine.Esc2", compressionValue, compressionValue, commentValue
                AppendSetting "Steel.ThreeLine.Es2", tensionValue, tensionValue, commentValue
            Case Else
                If Len(compressionValue) > 0 Then AppendSetting rowKey, compressionValue, compressionValue, commentValue
        End Select
NextRow:
    Next rowIndex
End Sub

' Распознает таблицу единиц по строке заголовков Quantity/INPUT/INTERNAL/OUTPUT.
Private Function IsUnitSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 4 Then Exit Function
    IsUnitSettingsRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Quantity", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "INPUT", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "INTERNAL", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 4))), "OUTPUT", vbTextCompare) = 0
End Function

' Распознает таблицу пользовательской системы знаков.
Private Function IsSignConventionSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 3 Then Exit Function
    IsSignConventionSettingsRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Quantity", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "USER", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "INTERNAL", vbTextCompare) = 0
End Function

' Распознает компактный блок настроек аннотаций схемы.
' В нем одна строка параметра задает значения сразу для RebarLabels и Dimensions.
Private Function IsPlotAnnotationSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 3 Or UBound(data, 2) < 5 Then Exit Function
    IsPlotAnnotationSettingsRange = _
        StrComp(Trim$(CStr(data(2, 1))), "Enabled", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(3, 1))), "Placement", vbTextCompare) = 0
End Function

' Разворачивает таблицу аннотаций в обычные ключи Plot.RebarLabels.*
' и Plot.Dimensions.*, чтобы остальной код читал настройки единообразно.
Private Sub AppendPlotAnnotationSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim settingName As String
        settingName = Trim$(CStr(data(rowIndex, 1)))
        If Len(settingName) > 0 And Left$(settingName, 1) <> "[" Then
            AppendPlotAnnotationValue "Plot.RebarLabels." & settingName, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 5)))
            AppendPlotAnnotationValue "Plot.Dimensions." & settingName, Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 5)))
        End If
    Next rowIndex
End Sub

' Добавляет настройку аннотации, пропуская пустые значения и "-".
Private Sub AppendPlotAnnotationValue(ByVal keyValue As String, ByVal settingValue As String, ByVal commentValue As String)
    If Len(settingValue) = 0 Or settingValue = "-" Then Exit Sub
    AppendSetting keyValue, settingValue, settingValue, commentValue
End Sub

' Загружает INPUT/INTERNAL/OUTPUT единицы в общий словарь настроек.
Private Sub AppendUnitSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = UnitQuantityKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(quantity) > 0 Then
            AppendSetting "Units." & quantity & ".Input", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), "Пользовательская единица ввода"
            AppendSetting "Units." & quantity & ".Internal", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), "Фиксированная внутренняя единица"
            AppendSetting "Units." & quantity & ".Output", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), "Пользовательская единица вывода"
        End If
    Next rowIndex
End Sub

' Загружает выбранные пользователем правила знаков N, Mx и My.
Private Sub AppendSignConventionSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = SignQuantityKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(quantity) > 0 Then
            AppendSetting "Sign." & quantity & ".User", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), "Пользовательское правило знаков"
            AppendSetting "Sign." & quantity & ".Internal", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), "Фиксированное внутреннее правило знаков"
        End If
    Next rowIndex
End Sub

' Преобразует подпись строки таблицы Units в суффикс ключа Units.*.
Private Function UnitQuantityKey(ByVal caption As String) As String
    Select Case LCase$(Trim$(caption))
        Case "length": UnitQuantityKey = "Length"
        Case "area": UnitQuantityKey = "Area"
        Case "force": UnitQuantityKey = "Force"
        Case "moment": UnitQuantityKey = "Moment"
        Case "stress": UnitQuantityKey = "Stress"
        Case "curvature": UnitQuantityKey = "Curvature"
    End Select
End Function

' Преобразует подпись строки таблицы знаков в суффикс ключа Sign.*.
Private Function SignQuantityKey(ByVal caption As String) As String
    Select Case LCase$(Trim$(caption))
        Case "+n": SignQuantityKey = "N"
        Case "+mx": SignQuantityKey = "Mx"
        Case "+my": SignQuantityKey = "My"
    End Select
End Function

' Распознает таблицу HollowRectangle по строке H/B/R/H_o/B_o/R_o.
Private Function IsHollowRectangleSettingsRange(ByRef data As Variant) As Boolean
    IsHollowRectangleSettingsRange = (HollowRectangleGeometryHeaderRow(data) > 0)
End Function

' Ищет строку основных размеров пустотелого скругленного прямоугольника.
Private Function HollowRectangleGeometryHeaderRow(ByRef data As Variant) As Long
    If UBound(data, 2) < 6 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1) - 1
        If StrComp(Trim$(CStr(data(rowIndex, 1))), "H", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 2))), "B", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 3))), "R", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 4))), "H_o", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 5))), "B_o", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 6))), "R_o", vbTextCompare) = 0 Then
            HollowRectangleGeometryHeaderRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Читает таблицу HollowRectangle: общие смещения Opening, геометрию,
' наружную арматуру и строки Opening с автоматическим числом стержней.
Private Sub AppendHollowRectangleSettings(ByRef data As Variant)
    Dim geometryHeaderRow As Long
    geometryHeaderRow = HollowRectangleGeometryHeaderRow(data)
    If geometryHeaderRow = 0 Then Exit Sub

    AppendHollowRectangleCommonSettings data, geometryHeaderRow

    Dim valueRow As Long
    valueRow = geometryHeaderRow + 1
    AppendSetting "HollowRectangle.H", Trim$(CStr(data(valueRow, 1))), Trim$(CStr(data(valueRow, 1))), "HollowRectangle geometry"
    AppendSetting "HollowRectangle.B", Trim$(CStr(data(valueRow, 2))), Trim$(CStr(data(valueRow, 2))), "HollowRectangle geometry"
    AppendSetting "HollowRectangle.R", Trim$(CStr(data(valueRow, 3))), Trim$(CStr(data(valueRow, 3))), "HollowRectangle geometry"
    AppendSetting "HollowRectangle.OpeningH", Trim$(CStr(data(valueRow, 4))), Trim$(CStr(data(valueRow, 4))), "HollowRectangle geometry"
    AppendSetting "HollowRectangle.OpeningB", Trim$(CStr(data(valueRow, 5))), Trim$(CStr(data(valueRow, 5))), "HollowRectangle geometry"
    AppendSetting "HollowRectangle.OpeningR", Trim$(CStr(data(valueRow, 6))), Trim$(CStr(data(valueRow, 6))), "HollowRectangle geometry"

    Dim mainHeaderRow As Long
    mainHeaderRow = FindRowWithCaption(data, "Грань", 1, geometryHeaderRow + 1)
    Dim extraHeaderRow As Long
    extraHeaderRow = FindRowWithCaption(data, "d2", 2, geometryHeaderRow + 1)
    If mainHeaderRow > 0 Then AppendHollowRectangleMainRebars data, mainHeaderRow
    If extraHeaderRow > 0 Then AppendHollowRectangleExtraRebars data, extraHeaderRow
End Sub

' Читает строки смещения отверстия над блоком геометрии.
Private Sub AppendHollowRectangleCommonSettings(ByRef data As Variant, ByVal geometryHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = 1 To geometryHeaderRow - 1
        Dim keyValue As String
        keyValue = Trim$(CStr(data(rowIndex, 1)))
        If StrComp(keyValue, "HollowRectangle.InnerOffsetX", vbTextCompare) = 0 Or _
                StrComp(keyValue, "HollowRectangle.InnerOffsetY", vbTextCompare) = 0 Then
            Dim sourceValue As String
            If UBound(data, 2) >= 4 Then sourceValue = Trim$(CStr(data(rowIndex, 4)))
            AppendSetting keyValue, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), sourceValue
        End If
    Next rowIndex
End Sub

' Читает основной ряд арматуры HollowRectangle по видимым пользователю граням H/B и Opening.
Private Sub AppendHollowRectangleMainRebars(ByRef data As Variant, ByVal mainHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = mainHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = HollowRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendSetting "HollowRectangle." & faceKey & ".as", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 6)))
        AppendSetting "HollowRectangle." & faceKey & ".d", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 6)))
        If IsNumericSettingText(data(rowIndex, 4)) Then
            AppendSetting "HollowRectangle." & faceKey & ".n", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 6)))
        End If
    Next rowIndex
End Sub

' Читает дополнительные ряды HollowRectangle и сохраняет их в единый набор ключей по граням.
Private Sub AppendHollowRectangleExtraRebars(ByRef data As Variant, ByVal extraHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = extraHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = HollowRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendSetting "HollowRectangle." & faceKey & ".d_2row", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "HollowRectangle." & faceKey & ".loc_2row", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "HollowRectangle." & faceKey & ".bind_2row", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "HollowRectangle." & faceKey & ".d_3row", Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "HollowRectangle." & faceKey & ".loc_3row", Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "HollowRectangle." & faceKey & ".bind_3row", Trim$(CStr(data(rowIndex, 7))), Trim$(CStr(data(rowIndex, 7))), Trim$(CStr(data(rowIndex, 9)))
    Next rowIndex
End Sub

' Переводит подпись строки HollowRectangle в устойчивый внутренний ключ грани.
Private Function HollowRectangleFaceKey(ByVal faceCaption As String) As String
    Select Case LCase$(Trim$(faceCaption))
        Case "h - левая", "h-left", "h.left": HollowRectangleFaceKey = "H.Left"
        Case "h - правая", "h-right", "h.right": HollowRectangleFaceKey = "H.Right"
        Case "b - верхняя", "b-top", "b.top": HollowRectangleFaceKey = "B.Top"
        Case "b - нижняя", "b-bottom", "b.bottom": HollowRectangleFaceKey = "B.Bottom"
        Case "h_o - левая", "ho-left", "opening.h.left": HollowRectangleFaceKey = "Opening.H.Left"
        Case "h_o - правая", "ho-right", "opening.h.right": HollowRectangleFaceKey = "Opening.H.Right"
        Case "b_o - верхняя", "bo-top", "opening.b.top": HollowRectangleFaceKey = "Opening.B.Top"
        Case "b_o - нижняя", "bo-bottom", "opening.b.bottom": HollowRectangleFaceKey = "Opening.B.Bottom"
    End Select
End Function

' Проверяет, можно ли строку настройки трактовать как число, не считая прочерки активными значениями.
Private Function IsNumericSettingText(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = Trim$(CStr(value))
    If Len(textValue) = 0 Or textValue = "-" Or UCase$(textValue) = "TODO" Then Exit Function
    IsNumericSettingText = IsNumeric(Replace$(textValue, ".", ApplicationDecimalSeparator()))
End Function

' Распознает новую табличную компоновку RoundedRectangle: B/H, левая/правая
' сторона и строки арматуры H/B.
Private Function IsRoundedRectangleFaceSettingsRange(ByRef data As Variant) As Boolean
    IsRoundedRectangleFaceSettingsRange = (RoundedRectangleGeometryHeaderRow(data) > 0)
End Function

' Ищет строку заголовков B/H в таблице скругленного прямоугольника.
Private Function RoundedRectangleGeometryHeaderRow(ByRef data As Variant) As Long
    If UBound(data, 2) < 3 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1) - 1
        If StrComp(Trim$(CStr(data(rowIndex, 2))), "B", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 3))), "H", vbTextCompare) = 0 Then
            RoundedRectangleGeometryHeaderRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Читает текущую таблицу RoundedRectangle и разворачивает видимые подписи
' граней в устойчивые ключи настроек.
Private Sub AppendRoundedRectangleFaceSettings(ByRef data As Variant)
    Dim geometryHeaderRow As Long
    geometryHeaderRow = RoundedRectangleGeometryHeaderRow(data)
    If geometryHeaderRow = 0 Then Exit Sub

    Dim valueRow As Long
    valueRow = geometryHeaderRow + 1
    AppendSetting "RoundedRectangle.B", Trim$(CStr(data(valueRow, 2))), Trim$(CStr(data(valueRow, 2))), "RoundedRectangle geometry"
    AppendSetting "RoundedRectangle.H", Trim$(CStr(data(valueRow, 3))), Trim$(CStr(data(valueRow, 3))), "RoundedRectangle geometry"

    Dim sideHeaderRow As Long
    sideHeaderRow = FindRowWithCaption(data, "слева", 2, geometryHeaderRow + 1)
    If sideHeaderRow > 0 And UBound(data, 1) >= sideHeaderRow + 4 Then
        AppendSetting "RoundedRectangle.Left.Type", Trim$(CStr(data(sideHeaderRow + 1, 2))), Trim$(CStr(data(sideHeaderRow + 1, 2))), Trim$(CStr(data(sideHeaderRow + 1, 5)))
        AppendSetting "RoundedRectangle.Right.Type", Trim$(CStr(data(sideHeaderRow + 1, 3))), Trim$(CStr(data(sideHeaderRow + 1, 3))), Trim$(CStr(data(sideHeaderRow + 1, 5)))
        AppendSetting "RoundedRectangle.Left.W", Trim$(CStr(data(sideHeaderRow + 2, 2))), Trim$(CStr(data(sideHeaderRow + 2, 2))), Trim$(CStr(data(sideHeaderRow + 2, 5)))
        AppendSetting "RoundedRectangle.Right.W", Trim$(CStr(data(sideHeaderRow + 2, 3))), Trim$(CStr(data(sideHeaderRow + 2, 3))), Trim$(CStr(data(sideHeaderRow + 2, 5)))
        AppendSetting "RoundedRectangle.Left.R1", Trim$(CStr(data(sideHeaderRow + 3, 2))), Trim$(CStr(data(sideHeaderRow + 3, 2))), Trim$(CStr(data(sideHeaderRow + 3, 5)))
        AppendSetting "RoundedRectangle.Right.R1", Trim$(CStr(data(sideHeaderRow + 3, 3))), Trim$(CStr(data(sideHeaderRow + 3, 3))), Trim$(CStr(data(sideHeaderRow + 3, 5)))
        AppendSetting "RoundedRectangle.Left.R2", Trim$(CStr(data(sideHeaderRow + 4, 2))), Trim$(CStr(data(sideHeaderRow + 4, 2))), Trim$(CStr(data(sideHeaderRow + 4, 5)))
        AppendSetting "RoundedRectangle.Right.R2", Trim$(CStr(data(sideHeaderRow + 4, 3))), Trim$(CStr(data(sideHeaderRow + 4, 3))), Trim$(CStr(data(sideHeaderRow + 4, 5)))
    End If

    Dim mainHeaderRow As Long
    mainHeaderRow = FindRowWithCaption(data, "Грань", 1, geometryHeaderRow + 1)
    Dim extraHeaderRow As Long
    extraHeaderRow = FindRowWithCaption(data, "d2", 2, geometryHeaderRow + 1)
    If mainHeaderRow > 0 Then AppendRoundedRectangleMainRebars data, mainHeaderRow
    If extraHeaderRow > 0 Then AppendRoundedRectangleExtraRebars data, extraHeaderRow
End Sub

' Ищет служебную строку шапки внутри табличной геометрии, начиная ниже уже найденного блока размеров.
Private Function FindRowWithCaption(ByRef data As Variant, ByVal caption As String, ByVal columnIndex As Long, ByVal startRow As Long) As Long
    Dim rowIndex As Long
    For rowIndex = startRow To UBound(data, 1)
        If StrComp(Trim$(CStr(data(rowIndex, columnIndex))), caption, vbTextCompare) = 0 Then
            FindRowWithCaption = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Читает основной ряд арматуры RoundedRectangle по четырем смысловым граням H/B.
Private Sub AppendRoundedRectangleMainRebars(ByRef data As Variant, ByVal mainHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = mainHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = RoundedRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendSetting "RoundedRectangle." & faceKey & ".as", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 6)))
        AppendSetting "RoundedRectangle." & faceKey & ".d", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 6)))
        AppendSetting "RoundedRectangle." & faceKey & ".n", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 6)))
    Next rowIndex
End Sub

' Читает дополнительные ряды RoundedRectangle и сохраняет привязку каждого ряда к своей грани.
Private Sub AppendRoundedRectangleExtraRebars(ByRef data As Variant, ByVal extraHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = extraHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = RoundedRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendSetting "RoundedRectangle." & faceKey & ".d_2row", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "RoundedRectangle." & faceKey & ".loc_2row", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "RoundedRectangle." & faceKey & ".bind_2row", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "RoundedRectangle." & faceKey & ".d_3row", Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "RoundedRectangle." & faceKey & ".loc_3row", Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 9)))
        AppendSetting "RoundedRectangle." & faceKey & ".bind_3row", Trim$(CStr(data(rowIndex, 7))), Trim$(CStr(data(rowIndex, 7))), Trim$(CStr(data(rowIndex, 9)))
    Next rowIndex
End Sub

' Переводит подпись строки RoundedRectangle в устойчивый внутренний ключ грани.
Private Function RoundedRectangleFaceKey(ByVal faceCaption As String) As String
    Select Case LCase$(Trim$(faceCaption))
        Case "h - левая", "h-left", "h.left": RoundedRectangleFaceKey = "H.Left"
        Case "h - правая", "h-right", "h.right": RoundedRectangleFaceKey = "H.Right"
        Case "b - верхняя", "b-top", "b.top": RoundedRectangleFaceKey = "B.Top"
        Case "b - нижняя", "b-bottom", "b.bottom": RoundedRectangleFaceKey = "B.Bottom"
    End Select
End Function

' Распознает табличную компоновку rngRectSetGeometry по заголовкам H1/B1/H2/B2.
Private Function IsRectSetFaceSettingsRange(ByRef data As Variant) As Boolean
    If RectSetGeometryHeaderRow(data) > 0 Then
        IsRectSetFaceSettingsRange = True
        Exit Function
    End If

    If UBound(data, 1) < 12 Or UBound(data, 2) < 7 Then Exit Function
    IsRectSetFaceSettingsRange = _
        StrComp(Trim$(CStr(data(1, 2))), "H1", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "B1", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 4))), "H2", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 5))), "B2", vbTextCompare) = 0
End Function

' Ищет строку заголовков H1/B1/H2/B2 в секционной таблице RectSet.
Private Function RectSetGeometryHeaderRow(ByRef data As Variant) As Long
    If UBound(data, 2) < 4 Then Exit Function

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1) - 1
        If StrComp(Trim$(CStr(data(rowIndex, 1))), "H1", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 2))), "B1", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 3))), "H2", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(rowIndex, 4))), "B2", vbTextCompare) = 0 Then
            RectSetGeometryHeaderRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

' Читает табличную компоновку RectSet и разворачивает ее в ключи настроек.
Private Sub AppendRectSetFaceSettings(ByRef data As Variant)
    Dim geometryHeaderRow As Long
    geometryHeaderRow = RectSetGeometryHeaderRow(data)
    If geometryHeaderRow > 0 And UBound(data, 1) >= geometryHeaderRow + 21 And UBound(data, 2) >= 9 Then
        AppendRectSetCommonSettings data, geometryHeaderRow
        AppendRectSetSectionedSettings data, geometryHeaderRow
        Exit Sub
    End If

    Dim faceColumns(1 To 4) As Long
    Dim faceNames(1 To 4) As String
    faceColumns(1) = 2: faceNames(1) = "H1"
    faceColumns(2) = 3: faceNames(2) = "B1"
    faceColumns(3) = 4: faceNames(3) = "H2"
    faceColumns(4) = 5: faceNames(4) = "B2"

    Dim i As Long
    ' Для каждой стороны H1/B1/H2/B2 разворачиваем пары _1/_2 в обычные ключи
    ' RectSet.<грань>.<параметр>. Solver и builder дальше не знают о форме таблицы.
    For i = 1 To 4
        Dim col As Long
        Dim faceName As String
        col = faceColumns(i)
        faceName = faceNames(i)
        AppendSetting "RectSet." & faceName, Trim$(CStr(data(2, col))), Trim$(CStr(data(2, col))), Trim$(CStr(data(2, 7)))
        AppendSetting "RectSet." & faceName & ".as_1", Trim$(CStr(data(3, col))), Trim$(CStr(data(3, col))), Trim$(CStr(data(3, 7)))
        AppendSetting "RectSet." & faceName & ".as_2", Trim$(CStr(data(4, col))), Trim$(CStr(data(4, col))), Trim$(CStr(data(4, 7)))
        AppendSetting "RectSet." & faceName & ".d_1", Trim$(CStr(data(5, col))), Trim$(CStr(data(5, col))), Trim$(CStr(data(5, 7)))
        AppendSetting "RectSet." & faceName & ".d_2", Trim$(CStr(data(6, col))), Trim$(CStr(data(6, col))), Trim$(CStr(data(6, 7)))
        AppendSetting "RectSet." & faceName & ".n_1", Trim$(CStr(data(7, col))), Trim$(CStr(data(7, col))), Trim$(CStr(data(7, 7)))
        AppendSetting "RectSet." & faceName & ".n_2", Trim$(CStr(data(8, col))), Trim$(CStr(data(8, col))), Trim$(CStr(data(8, 7)))
        AppendSetting "RectSet." & faceName & ".StartOffset1", Trim$(CStr(data(9, col))), Trim$(CStr(data(9, col))), Trim$(CStr(data(9, 7)))
        AppendSetting "RectSet." & faceName & ".EndOffset1", Trim$(CStr(data(10, col))), Trim$(CStr(data(10, col))), Trim$(CStr(data(10, 7)))
        AppendSetting "RectSet." & faceName & ".StartOffset2", Trim$(CStr(data(11, col))), Trim$(CStr(data(11, col))), Trim$(CStr(data(11, 7)))
        AppendSetting "RectSet." & faceName & ".EndOffset2", Trim$(CStr(data(12, col))), Trim$(CStr(data(12, col))), Trim$(CStr(data(12, 7)))
        If UBound(data, 1) >= 19 Then
            AppendSetting "RectSet." & faceName & ".d_2row_1", Trim$(CStr(data(14, col))), Trim$(CStr(data(14, col))), Trim$(CStr(data(14, 7)))
            AppendSetting "RectSet." & faceName & ".d_2row_2", Trim$(CStr(data(15, col))), Trim$(CStr(data(15, col))), Trim$(CStr(data(15, 7)))
            AppendSetting "RectSet." & faceName & ".d_3row_1", Trim$(CStr(data(16, col))), Trim$(CStr(data(16, col))), Trim$(CStr(data(16, 7)))
            AppendSetting "RectSet." & faceName & ".d_3row_2", Trim$(CStr(data(17, col))), Trim$(CStr(data(17, col))), Trim$(CStr(data(17, 7)))
            AppendSetting "RectSet." & faceName & ".loc_2row", Trim$(CStr(data(18, col))), Trim$(CStr(data(18, col))), Trim$(CStr(data(18, 7)))
            AppendSetting "RectSet." & faceName & ".loc_3row", Trim$(CStr(data(19, col))), Trim$(CStr(data(19, col))), Trim$(CStr(data(19, 7)))
            If UBound(data, 1) >= 21 Then
                AppendSetting "RectSet." & faceName & ".bind_2row", Trim$(CStr(data(20, col))), Trim$(CStr(data(20, col))), Trim$(CStr(data(20, 7)))
                AppendSetting "RectSet." & faceName & ".bind_3row", Trim$(CStr(data(21, col))), Trim$(CStr(data(21, col))), Trim$(CStr(data(21, 7)))
            End If
        End If
    Next i
End Sub

' Читает общие настройки RectSet над блоком H1/B1/H2/B2.
Private Sub AppendRectSetCommonSettings(ByRef data As Variant, ByVal geometryHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = 1 To geometryHeaderRow - 1
        Dim keyValue As String
        keyValue = Trim$(CStr(data(rowIndex, 1)))
        If StrComp(keyValue, "RectSet.SectionType", vbTextCompare) = 0 Or _
                StrComp(keyValue, "RectSet.UpperOffsetX", vbTextCompare) = 0 Then
            Dim sourceValue As String
            If UBound(data, 2) >= 4 Then sourceValue = Trim$(CStr(data(rowIndex, 4)))
            AppendSetting keyValue, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), sourceValue
        End If
    Next rowIndex
End Sub

' Читает текущую табличную компоновку RectSet-сечения:
' отдельный блок геометрии, блок основного армирования и блок дополнительных рядов.
Private Sub AppendRectSetSectionedSettings(ByRef data As Variant, Optional ByVal geometryHeaderRow As Long = 2)
    Dim valueRow As Long
    valueRow = geometryHeaderRow + 1
    AppendSetting "RectSet.H1", Trim$(CStr(data(valueRow, 1))), Trim$(CStr(data(valueRow, 1))), "RectSet geometry"
    AppendSetting "RectSet.B1", Trim$(CStr(data(valueRow, 2))), Trim$(CStr(data(valueRow, 2))), "RectSet geometry"
    AppendSetting "RectSet.H2", Trim$(CStr(data(valueRow, 3))), Trim$(CStr(data(valueRow, 3))), "RectSet geometry"
    AppendSetting "RectSet.B2", Trim$(CStr(data(valueRow, 4))), Trim$(CStr(data(valueRow, 4))), "RectSet geometry"

    Dim faceNames(1 To 8) As String
    Dim sideIndexes(1 To 8) As Long
    FillRectSetSectionedFaceOrder faceNames, sideIndexes

    Dim i As Long
    ' Восемь строк граней превращаются в ключи с суффиксами _1/_2.
    For i = 1 To 8
        Dim rowIndex As Long
        rowIndex = geometryHeaderRow + 3 + i
        AppendSetting "RectSet." & faceNames(i) & ".as_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "RectSet." & faceNames(i) & ".d_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "RectSet." & faceNames(i) & ".n_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "RectSet." & faceNames(i) & ".StartOffset" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "RectSet." & faceNames(i) & ".EndOffset" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 8)))

        Dim extraRow As Long
        extraRow = geometryHeaderRow + 13 + i
        AppendSetting "RectSet." & faceNames(i) & ".d_2row_" & CStr(sideIndexes(i)), Trim$(CStr(data(extraRow, 2))), Trim$(CStr(data(extraRow, 2))), Trim$(CStr(data(extraRow, 9)))
        AppendSetting "RectSet." & faceNames(i) & ".d_3row_" & CStr(sideIndexes(i)), Trim$(CStr(data(extraRow, 5))), Trim$(CStr(data(extraRow, 5))), Trim$(CStr(data(extraRow, 9)))
        If sideIndexes(i) = 1 Then
            AppendSetting "RectSet." & faceNames(i) & ".loc_2row", Trim$(CStr(data(extraRow, 3))), Trim$(CStr(data(extraRow, 3))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "RectSet." & faceNames(i) & ".bind_2row", Trim$(CStr(data(extraRow, 4))), Trim$(CStr(data(extraRow, 4))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "RectSet." & faceNames(i) & ".loc_3row", Trim$(CStr(data(extraRow, 6))), Trim$(CStr(data(extraRow, 6))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "RectSet." & faceNames(i) & ".bind_3row", Trim$(CStr(data(extraRow, 7))), Trim$(CStr(data(extraRow, 7))), Trim$(CStr(data(extraRow, 9)))
        End If
    Next i
End Sub

' Задает порядок строк в секционной таблице RectSet.
' _1 означает левую/верхнюю грань, _2 - правую/нижнюю.
Private Sub FillRectSetSectionedFaceOrder(ByRef faceNames() As String, ByRef sideIndexes() As Long)
    faceNames(1) = "H1": sideIndexes(1) = 1
    faceNames(2) = "H1": sideIndexes(2) = 2
    faceNames(3) = "B1": sideIndexes(3) = 1
    faceNames(4) = "B1": sideIndexes(4) = 2
    faceNames(5) = "H2": sideIndexes(5) = 1
    faceNames(6) = "H2": sideIndexes(6) = 2
    faceNames(7) = "B2": sideIndexes(7) = 1
    faceNames(8) = "B2": sideIndexes(8) = 2
End Sub

' Добавляет одну пару ключ-значение в reader.
' Массивы параллельные: это проще и быстрее для VBA, чем словарь с Variant-объектами,
' а поиск по ключу остается достаточно дешевым для числа настроек Config.
Private Sub AppendSetting(ByVal keyValue As String, ByVal settingValue As String, _
        ByVal defaultValue As String, ByVal sourceValue As String)
    mCount = mCount + 1
    ReDim Preserve mKeys(1 To mCount)
    ReDim Preserve mValues(1 To mCount)
    ReDim Preserve mDefaults(1 To mCount)
    ReDim Preserve mSources(1 To mCount)
    mKeys(mCount) = keyValue
    mValues(mCount) = settingValue
    mDefaults(mCount) = defaultValue
    mSources(mCount) = sourceValue
End Sub

' Сбрасывает все прочитанные настройки и точки диаграмм.
Public Sub Clear()
    Erase mKeys
    Erase mValues
    Erase mDefaults
    Erase mSources
    mCount = 0
End Sub

' Возвращает строковое значение настройки.
' Пустая строка и TODO трактуются как отсутствие пользовательского значения,
' поэтому вызывающий код получает свой defaultValue.
Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetString = defaultValue
    ElseIf Len(mValues(index)) = 0 Or UCase$(mValues(index)) = "TODO" Then
        GetString = defaultValue
    Else
        GetString = mValues(index)
    End If
End Function

' Возвращает сырую строку без подстановки default при пустом значении.
' Это нужно для настроек, где пустое значение является ошибкой ввода.
Public Function GetRawString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetRawString = defaultValue
    Else
        GetRawString = mValues(index)
    End If
End Function

' Возвращает настройку как Double, учитывая десятичный разделитель Excel.
Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
    Dim textValue As String
    textValue = GetString(key, vbNullString)
    If Len(textValue) = 0 Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
    End If
End Function

' Читает обязательную числовую настройку без скрытой подстановки значения.
' Используется для характеристик материалов: если пользователь очистил ячейку
' или ключ не попал в Config, расчет должен остановиться как ошибка исходных данных.
Public Function GetRequiredDouble(ByVal key As String) As Double
    Dim textValue As String
    textValue = Trim$(GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        Err.Raise vbObjectError + 4310, "CSystemSettingsReader", _
            "Обязательная числовая настройка " & key & " не задана."
    End If
    GetRequiredDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
End Function

' Читает обязательную текстовую настройку без default-значения.
' Нужна для пользовательских переключателей, которые являются частью
' расчетной материальной модели, например профиля арматуры.
Public Function GetRequiredString(ByVal key As String) As String
    Dim textValue As String
    textValue = Trim$(GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        Err.Raise vbObjectError + 4311, "CSystemSettingsReader", _
            "Обязательная настройка " & key & " не задана."
    End If
    GetRequiredString = textValue
End Function

' Возвращает настройку как Long через тот же числовой путь, что и GetDouble.
Public Function GetLong(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    GetLong = CLng(GetDouble(key, CDbl(defaultValue)))
End Function

' Читает булеву настройку из привычных вариантов Yes/No, True/False, 1/0, Да/Нет.
Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim textValue As String
    textValue = UCase$(GetString(key, IIf(defaultValue, "YES", "NO")))
    GetBoolean = (textValue = "YES" Or textValue = "TRUE" Or textValue = "1" Or textValue = "ДА")
End Function

' Возвращает комментарий/источник настройки для диагностических сообщений.
Public Function GetSource(ByVal key As String) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetSource = vbNullString
    Else
        GetSource = mSources(index)
    End If
End Function

' Проверяет наличие ключа в загруженных настройках.
Public Function HasKey(ByVal key As String) As Boolean
    HasKey = (FindKey(key) > 0)
End Function

' Количество ключей удобно для тестов структуры Config.
Public Function KeyCount() As Long
    KeyCount = mCount
End Function

' Возвращает имя ключа по порядковому номеру для проверок дубликатов и тестов.
Public Function KeyAt(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4300, "CSystemSettingsReader", "Индекс настройки вне допустимого диапазона."
    KeyAt = mKeys(index)
End Function

' Считает повторяющиеся ключи Config.
' Дубликаты опасны тем, что один и тот же параметр может быть прочитан из двух
' таблиц с разными значениями.
Public Function DuplicateCount() As Long
    Dim duplicates As Long
    Dim i As Long
    Dim j As Long
    For i = 1 To mCount
        For j = i + 1 To mCount
            If StrComp(mKeys(i), mKeys(j), vbTextCompare) = 0 Then
                duplicates = duplicates + 1
                Exit For
            End If
        Next j
    Next i
    DuplicateCount = duplicates
End Function

' Ищет ключ линейным проходом по небольшому массиву настроек.
' Для текущего размера Config словарь не нужен, зато порядок и дубликаты легко тестировать.
Private Function FindKey(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mKeys(i), key, vbTextCompare) = 0 Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = 0
End Function

' Возвращает десятичный разделитель текущего Excel.
' Это позволяет читать точки диаграмм и настройки с точкой или запятой в разных локалях.
Private Function ApplicationDecimalSeparator() As String
    ApplicationDecimalSeparator = Mid$(CStr(1.1), 2, 1)
End Function






