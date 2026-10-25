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
' Чтение настроек Config и характеристик материалов
' ==========================================================================
' Собирает ключи и значения из именованных диапазонов, включая специальные
' таблицы материалов. По прочитанным параметрам CMaterialModelProvider
' строит расчётные диаграммы бетона и арматуры.
' При отсутствии обязательного ключа или таблицы выдаёт понятную ошибку ввода.

Private mKeys() As String ' Имена настроек Config: Solver.Method, Capacity.SolutionStrategy и т.п.
Private mValues() As String ' Значения, которые ввел пользователь или подставили формулы таблицы.
Private mValueErrors() As String ' Отложенные ошибки условных полей: выдаются при запросе активного параметра без подстановки значения.
Private mDefaults() As String ' Сопутствующий текст, переданный при AppendSetting; сам по себе не заменяет пустое значение настройки.
Private mSources() As String ' Комментарий/источник строки для диагностики.
Private mLocations() As String ' Лист, адрес и видимый столбец прочитанного значения; только диагностика ввода.
Private mLoadingSheet As String ' Имя листа текущей входной таблицы для диагностики.
Private mLoadingRow As Long ' Первая строка таблицы; считывается один раз, без COM на каждый ключ.
Private mLoadingColumn As Long ' Первый столбец таблицы для вычисления абсолютного адреса значения.
Private mCount As Long ' Количество загруженных ключей.

' Загружает настройки расчёта и вывода из именованных таблиц Config
' в общие массивы ключей, значений и адресов. Таблицы материалов
' разворачивает в ключи Concrete.* и Steel.* без обращения ядра к Excel.
Public Sub LoadFromWorkbook(ByVal workbook As Object)
    Clear
    AppendRequiredNamedSettingsRange workbook, "rngUnitSettings", "Единицы измерения"
    AppendRequiredNamedSettingsRange workbook, "rngSignConventionSettings", "Система знаков"
    ValidateWorkbookUnitAndSignKeys
    AppendRequiredNamedSettingsRange workbook, "rngSystemSettings", "Системные настройки"
    AppendNamedSettingsRangeIfExists workbook, "rngConcreteMaterialParameters"
    AppendNamedSettingsRangeIfExists workbook, "rngSteelMaterialParameters"
    AppendRequiredNamedSettingsRange workbook, "rngPlotAnnotationSettings", "Подписи и размеры схемы"
    ValidateDiagramExtensionSetting True
End Sub

' Различает поврежденное имя обязательной таблицы и ошибку ее значения.
' При утрате имени дает действие для восстановления книги вместо COM-ошибки 1004.
Private Sub AppendRequiredNamedSettingsRange(ByVal workbook As Object, ByVal rangeName As String, ByVal caption As String)
    Dim settingsRange As Object
    On Error GoTo MissingRange
    Set settingsRange = workbook.Names.Item(rangeName).RefersToRange
    On Error GoTo 0
    AppendSettingsRange settingsRange, rangeName
    Exit Sub
MissingRange:
    Err.Raise vbObjectError + 4316, "CSystemSettingsReader.LoadFromWorkbook", _
        "На листе Config недоступна обязательная таблица """ & caption & """ (имя " & rangeName & _
        "). Восстановите таблицу и ссылку на нее в диспетчере имен Excel или используйте исправный шаблон книги."
End Sub

' После чтения таблиц проверяет обязательные настройки полной книги.
' Отсутствующая или повреждённая строка Config вызывает ошибку без подстановки
' значения по умолчанию. Отдельный LoadFromRange может читать лишь часть настроек.
Private Sub ValidateWorkbookUnitAndSignKeys()
    Dim key As Variant
    For Each key In Array("Units.Length.Input", "Units.Length.Output", "Units.Area.Input", "Units.Area.Output", _
            "Units.Force.Input", "Units.Force.Output", "Units.Moment.Input", "Units.Moment.Output", _
            "Units.Stress.Input", "Units.Stress.Output", "Units.Curvature.Input", "Units.Curvature.Output", _
            "Sign.N.User", "Sign.Mx.User", "Sign.My.User")
        If Not HasKey(CStr(key)) Then
            Err.Raise vbObjectError + 4313, "CSystemSettingsReader.LoadFromWorkbook", _
                InputErrorMessage(CStr(key), "Обязательная строка настройки отсутствует.", _
                    "Восстановите строку в указанной таблице и заполните пользовательское значение.")
        End If
    Next key
End Sub

' Подключает необязательную таблицу, если её имя есть в книге.
' Таблица может отсутствовать до запроса её данных расчётом,
' но существующее повреждённое имя не скрывается как отсутствие ввода.
Private Sub AppendNamedSettingsRangeIfExists(ByVal workbook As Object, ByVal rangeName As String)
    On Error GoTo MissingName
    Dim nameObject As Object
    Set nameObject = workbook.Names.Item(rangeName)
    On Error GoTo BrokenRange
    Dim settingsRange As Object
    Set settingsRange = nameObject.RefersToRange
    On Error GoTo 0
    AppendSettingsRange settingsRange, rangeName
    Exit Sub
MissingName:
    On Error GoTo 0
    Exit Sub
BrokenRange:
    Err.Raise vbObjectError + 4316, "CSystemSettingsReader.LoadFromWorkbook", _
        "Таблица параметров материала (именованный диапазон " & rangeName & ") недоступна. " & _
        "Восстановите ссылку на один непрерывный прямоугольный диапазон со всей таблицей " & _
        "в диспетчере имен Excel или используйте исправный шаблон книги."
End Sub

' Загружает настройки из одного диапазона.
' Этот короткий путь используют тесты и вспомогательные сценарии без полной книги.
Public Sub LoadFromRange(ByVal settingsRange As Object)
    Clear
    AppendSettingsRange settingsRange
    ValidateDiagramExtensionSetting
End Sub

' Проверяет текущий переключатель без преобразования структуры и переименования ключей.
' Для неполного автономного LoadFromRange добавляет Yes; в полной книге строка обязательна.
Private Sub ValidateDiagramExtensionSetting(Optional ByVal requireWorkbookValue As Boolean = False)
    Const key As String = "General.DiagramExtension"
    Dim index As Long
    index = FindKey(key)
    If index > 0 Then
        mValues(index) = DiagramExtensionValue(mValues(index), key)
    ElseIf requireWorkbookValue Then
        Err.Raise vbObjectError + 4313, "CSystemSettingsReader", _
            InputErrorMessage(key, "Обязательная строка настройки отсутствует.", _
                "Восстановите строку в системных настройках Config и выберите Yes или No.")
    Else
        AppendSetting key, "Yes", "Yes", "Численное продолжение активных диаграмм."
    End If
End Sub

' Читает поддерживаемые варианты да/нет и возвращает Yes или No.
' Пустое или неизвестное значение вызывает ошибку Config.
Private Function DiagramExtensionValue(ByVal rawValue As String, ByVal settingKey As String) As String
    Select Case UCase$(Trim$(rawValue))
        Case "YES", "TRUE", "1", "ДА": DiagramExtensionValue = "Yes"
        Case "NO", "FALSE", "0", "НЕТ": DiagramExtensionValue = "No"
        Case Else
            Err.Raise vbObjectError + 4310, "CSystemSettingsReader", _
                InputErrorMessage(settingKey, "Значение переключателя не распознано.", "Выберите Yes или No.")
    End Select
End Function

' Добавляет одну непрерывную таблицу общих настроек. Несмежные области отвергаются до Value2:
' иначе Excel возвращает только первую область и часть ввода теряется.
Private Sub AppendSettingsRange(ByVal settingsRange As Object, Optional ByVal rangeName As String = vbNullString)
    If settingsRange Is Nothing Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
        "Таблица настроек Config не передана."
    mLoadingSheet = CStr(settingsRange.Worksheet.Name)
    mLoadingRow = settingsRange.Row
    mLoadingColumn = settingsRange.Column
    Dim tableDescription As String
    tableDescription = "Таблица настроек"
    If Len(rangeName) > 0 Then tableDescription = tableDescription & " " & rangeName
    If settingsRange.Areas.Count <> 1 Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
        tableDescription & " (" & InputCellLocation(1, 1) & _
        ") содержит несколько отдельных областей. Восстановите ссылку на один непрерывный " & _
        "прямоугольный диапазон со всей таблицей в диспетчере имен Excel; " & _
        "заголовок и все строки данных должны находиться в этой области."
    Dim data As Variant
    data = settingsRange.Value2
    If Not IsArray(data) Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
        "Таблица настроек (" & InputCellLocation(1, 1) & ") имеет неполную структуру. " & _
        "Восстановите заголовок, строки данных, минимум три столбца и ссылку на таблицу в диспетчере имен Excel."
    If UBound(data, 1) < 2 Or UBound(data, 2) < 3 Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
        "Таблица настроек (" & InputCellLocation(1, 1) & ") имеет неполную структуру. " & _
        "Восстановите заголовок, строки данных, минимум три столбца и ссылку на таблицу в диспетчере имен Excel."
    Dim columnCount As Long
    columnCount = UBound(data, 2)

    Select Case LCase$(rangeName) ' Имя известных таблиц сохраняет их смысл при повреждении шапки. Сначала называем конкретную ячейку структуры, затем читаем значения; такой дефект не должен выглядеть как исчезновение обязательных ключей.
        Case "rngunitsettings"
            RequireSettingsHeaders data, "Единицы измерения", Array("Quantity", "INPUT", "INTERNAL", "OUTPUT"), 4
            AppendUnitSettings data
            Exit Sub
        Case "rngsignconventionsettings"
            RequireSettingsHeaders data, "Система знаков", Array("Quantity", "USER", "INTERNAL"), 3
            AppendSignConventionSettings data
            Exit Sub
        Case "rngconcretematerialparameters"
            RequireSettingsHeaders data, "Параметры бетона", Array("Параметр", "Сжатие", "Растяжение"), 4
            AppendConcreteMaterialParameters data
            Exit Sub
        Case "rngsteelmaterialparameters"
            RequireSettingsHeaders data, "Параметры арматуры", Array("Параметр", "Сжатие", "Растяжение"), 4
            AppendSteelMaterialParameters data
            Exit Sub
    End Select

    If StrComp(rangeName, "rngPlotAnnotationSettings", vbTextCompare) = 0 Then ' Тип именованной таблицы сохраняется даже при удалении строки параметра. Например, отсутствие Placement не должно превратить остальные поля таблицы оформления в обычную таблицу пар ключ/значение.
        If columnCount < 5 Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
            "Таблица аннотаций схемы на листе " & mLoadingSheet & " имеет неполную структуру. Восстановите колонки параметров, арматуры, размеров, единиц и комментария."
        AppendPlotAnnotationSettings data
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
    For rowIndex = 2 To UBound(data, 1) ' Обычные таблицы читаются как ключ/значение. Строки с заголовками секций в квадратных скобках пропускаются.
        Dim keyValue As String
        keyValue = ConfigCellText(data(rowIndex, 1), "ключ настройки", rowIndex, vbNullString, False, 1)
        If Len(keyValue) > 0 And Left$(keyValue, 1) <> "[" Then
            Dim sourceValue As String
            If columnCount >= 6 Then
                sourceValue = ConfigCellText(data(rowIndex, 6), "комментарий", rowIndex, keyValue, True)
            ElseIf columnCount >= 4 Then
                sourceValue = ConfigCellText(data(rowIndex, 4), "комментарий", rowIndex, keyValue, True)
            Else
                sourceValue = vbNullString
            End If
            If Left$(keyValue, 5) = "Plot." Or keyValue = "Mesh.StepX" Or keyValue = "Mesh.StepY" Then ' Ручные шаги не являются входом Auto; в User их ошибки выдаются при обязательном чтении.
                AppendDeferredInputValue data, keyValue, rowIndex, 2, sourceValue
            Else
                AppendSetting keyValue, _
                    ConfigCellText(data(rowIndex, 2), "значение", rowIndex, keyValue, False, 2), _
                    ConfigCellText(data(rowIndex, 3), "единица измерения", rowIndex, keyValue, True), _
                    sourceValue, rowIndex, 2
            End If
        End If
    Next rowIndex
End Sub

' Проверяет столбцы данных в известных именованных таблицах.
' Обязательные заголовки определяют способ чтения значений. Адрес ошибки
' вычисляется по текущей привязке имени диапазона без фиксированных координат.
Private Sub RequireSettingsHeaders(ByRef data As Variant, ByVal tableCaption As String, _
        ByVal headers As Variant, ByVal minimumColumns As Long)
    If UBound(data, 2) < minimumColumns Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
        "Таблица """ & tableCaption & """ (" & InputCellLocation(1, 1) & ") имеет неполную структуру. " & _
        "Восстановите все колонки таблицы и ссылку на нее в диспетчере имен Excel."
    Dim index As Long, actual As String
    For index = LBound(headers) To UBound(headers)
        actual = ConfigCellText(data(1, index + 1), "заголовок таблицы", 1, tableCaption, False, index + 1)
        If StrComp(actual, CStr(headers(index)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
                "Таблица """ & tableCaption & """ (" & InputCellLocation(1, index + 1) & _
                "): в заголовке колонки должно быть """ & CStr(headers(index)) & """, сейчас """ & actual & _
                """. Восстановите заголовок и расположение колонок по исправному шаблону книги."
        End If
    Next index
End Sub

' Возвращает текст ячейки Config, сохраняя различие между обычным значением
' и ошибкой формулы Excel. Ошибка расчетного поля получает точный адрес;
' нерасчетные комментарии могут игнорировать ее вместо текста Error 2015.
Private Function ConfigCellText(ByVal cellValue As Variant, ByVal columnName As String, _
        ByVal rowIndex As Long, ByVal keyValue As String, _
        Optional ByVal allowFormulaErrorAsEmpty As Boolean = False, Optional ByVal columnIndex As Long = 0) As String
    If IsError(cellValue) Then
        If allowFormulaErrorAsEmpty Then Exit Function
        If Len(keyValue) = 0 Then keyValue = "строка " & CStr(rowIndex)
        Err.Raise vbObjectError + 4309, "CSystemSettingsReader", _
            "Настройка " & keyValue & " (" & InputCellLocation(rowIndex, columnIndex) & _
            ") содержит ошибку Excel в колонке """ & columnName & _
            """. Исправьте формулу или введите допустимое значение вместо ошибки Excel."
    End If
    ConfigCellText = Trim$(CStr(cellValue))
End Function

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
        rowKey = ConfigCellText(data(rowIndex, 1), "параметр бетона", rowIndex, vbNullString, False, 1)
        If Len(rowKey) = 0 Then GoTo NextRow

        Dim compressionValue As String
        Dim tensionValue As String
        Dim commentValue As String
        compressionValue = ConfigCellText(data(rowIndex, 2), "Сжатие", rowIndex, rowKey, False, 2)
        tensionValue = vbNullString
        If StrComp(rowKey, "Concrete.Rb.mc2", vbTextCompare) <> 0 Then _
            tensionValue = ConfigCellText(data(rowIndex, 3), "Растяжение", rowIndex, rowKey, False, 3) ' Rb,mc2 относится только к сжатию; неприменимая клетка растяжения не является входным параметром, даже если в ней ошибочная формула.
        If UBound(data, 2) >= 5 Then commentValue = ConfigCellText(data(rowIndex, 5), "комментарий", rowIndex, rowKey, True, 5)

        Select Case LCase$(rowKey)
            Case "concrete.r.uls(i)", "concrete.r.i(uls)", "concrete.r.uls"
                AppendSetting "Concrete.Rb.ULS", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Rbt.ULS", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "concrete.r.sls(ii)", "concrete.r.ii(sls)", "concrete.r.sls"
                AppendSetting "Concrete.Rb.SLS", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Rbt.SLS", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "concrete.rb.mc2"
                AppendSetting "Concrete.Rb.mc2", compressionValue, compressionValue, commentValue, rowIndex, 2
            Case "concrete.e"
                AppendSetting "Concrete.Eb", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Ebt", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "concrete.twoline.eb1red"
                AppendSetting "Concrete.Eb1Red", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Ebt1Red", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "concrete.threeline.eb0"
                AppendSetting "Concrete.Eb0", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Ebt0", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "concrete.twothreeline.eb2"
                AppendSetting "Concrete.Eb2", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Concrete.Ebt2", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case Else
                If Len(compressionValue) > 0 Then AppendSetting rowKey, compressionValue, compressionValue, commentValue, rowIndex, 2
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
        rowKey = ConfigCellText(data(rowIndex, 1), "параметр арматуры", rowIndex, vbNullString, False, 1)
        If Len(rowKey) = 0 Then GoTo NextRow

        Dim compressionValue As String
        Dim tensionValue As String
        Dim commentValue As String
        compressionValue = ConfigCellText(data(rowIndex, 2), "Сжатие", rowIndex, rowKey, False, 2)
        tensionValue = ConfigCellText(data(rowIndex, 3), "Растяжение", rowIndex, rowKey, False, 3)
        If UBound(data, 2) >= 5 Then commentValue = ConfigCellText(data(rowIndex, 5), "комментарий", rowIndex, rowKey, True, 5)

        Select Case LCase$(rowKey)
            Case "steel.r.uls(i)", "steel.r.i(uls)", "steel.r.uls"
                AppendSetting "Steel.Rsc.ULS", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Steel.Rs.ULS", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "steel.r.sls(ii)", "steel.r.ii(sls)", "steel.r.sls"
                AppendSetting "Steel.Rsc.SLS", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Steel.Rs.SLS", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "steel.e"
                AppendSetting "Steel.Esc", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Steel.Es", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "steel.twoline.es2"
                AppendSetting "Steel.TwoLine.Esc2", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Steel.TwoLine.Es2", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case "steel.threeline.es2"
                AppendSetting "Steel.ThreeLine.Esc2", compressionValue, compressionValue, commentValue, rowIndex, 2
                AppendSetting "Steel.ThreeLine.Es2", tensionValue, tensionValue, commentValue, rowIndex, 3
            Case Else
                If Len(compressionValue) > 0 Then AppendSetting rowKey, compressionValue, compressionValue, commentValue, rowIndex, 2
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

' Разворачивает общую толщину контура и отдельные параметры подписей/размеров.
Private Sub AppendPlotAnnotationSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim settingName As String
        settingName = ConfigCellText(data(rowIndex, 1), "параметр аннотации", rowIndex, vbNullString, False, 1)
        If Len(settingName) > 0 And Left$(settingName, 1) <> "[" Then
            If settingName = "ContourLineWeight" Then
                AppendDeferredInputValue data, "Plot.ContourLineWeight", rowIndex, 2, _
                    ConfigCellText(data(rowIndex, 5), "комментарий", rowIndex, "Plot.ContourLineWeight", True, 5)
            Else
                AppendPlotAnnotationValue data, "Plot.RebarLabels." & settingName, rowIndex, 2
                AppendPlotAnnotationValue data, "Plot.Dimensions." & settingName, rowIndex, 3
            End If
        End If
    Next rowIndex
End Sub

' Добавляет только применимое поле таблицы аннотаций. Пустое применимое поле
' сохраняет адрес и остается ошибкой активного потребителя; неподдерживаемые
' сочетания колонки и параметра не становятся настройками даже при ошибке ячейки.
Private Sub AppendPlotAnnotationValue(ByRef data As Variant, ByVal keyValue As String, _
        ByVal rowIndex As Long, ByVal columnIndex As Long)
    Dim settingName As String
    settingName = Trim$(CStr(data(rowIndex, 1)))
    If columnIndex = 2 Then
        Select Case settingName
            Case "ExtensionLineWeight", "ExtensionLineColor", "ArrowType", "ArrowSize": Exit Sub
        End Select
    ElseIf settingName = "LineEnabled" Then
        Exit Sub
    End If
    AppendDeferredInputValue data, keyValue, rowIndex, columnIndex, _
        ConfigCellText(data(rowIndex, 5), "комментарий", rowIndex, keyValue, True, 5)
End Sub

' Разворачивает таблицу INPUT/INTERNAL/OUTPUT в общие массивы ключей Units.*.
Private Sub AppendUnitSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = UnitQuantityKey(ConfigCellText(data(rowIndex, 1), "физическая величина", rowIndex, vbNullString, False, 1))
        If Len(quantity) > 0 Then
            Dim value As String
            value = ConfigCellText(data(rowIndex, 2), "INPUT", rowIndex, "Units." & quantity & ".Input", False, 2)
            AppendSetting "Units." & quantity & ".Input", value, value, "Пользовательская единица ввода", rowIndex, 2
            value = ConfigCellText(data(rowIndex, 3), "INTERNAL", rowIndex, "Units." & quantity & ".Internal", False, 3)
            AppendSetting "Units." & quantity & ".Internal", value, value, "Фиксированная внутренняя единица", rowIndex, 3
            value = ConfigCellText(data(rowIndex, 4), "OUTPUT", rowIndex, "Units." & quantity & ".Output", False, 4)
            AppendSetting "Units." & quantity & ".Output", value, value, "Пользовательская единица вывода", rowIndex, 4
        End If
    Next rowIndex
End Sub

' Загружает выбранные пользователем правила знаков N, Mx и My.
Private Sub AppendSignConventionSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = SignQuantityKey(ConfigCellText(data(rowIndex, 1), "компонента нагрузки", rowIndex, vbNullString, False, 1))
        If Len(quantity) > 0 Then
            Dim value As String
            value = ConfigCellText(data(rowIndex, 2), "USER", rowIndex, "Sign." & quantity & ".User", False, 2)
            AppendSetting "Sign." & quantity & ".User", value, value, "Пользовательское правило знаков", rowIndex, 2
            value = ConfigCellText(data(rowIndex, 3), "INTERNAL", rowIndex, "Sign." & quantity & ".Internal", False, 3)
            AppendSetting "Sign." & quantity & ".Internal", value, value, "Фиксированное внутреннее правило знаков", rowIndex, 3
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

' Сохраняет условный ввод вместе с возможной ошибкой ячейки.
' Ошибка выдаётся при запросе активного параметра; выключенные аннотации
' и автоматический шаг сетки не требуют чтения неиспользуемых ручных полей.
Private Sub AppendDeferredInputValue(ByRef data As Variant, ByVal key As String, _
        ByVal rowIndex As Long, ByVal columnIndex As Long, Optional ByVal source As String = vbNullString)
    Dim value As String, valueError As String
    If IsError(data(rowIndex, columnIndex)) Then
        valueError = "Настройка " & key & " (" & InputCellLocation(rowIndex, columnIndex) & _
            ") содержит ошибку Excel в колонке ""значение"". " & _
            "Исправьте формулу или введите допустимое значение вместо ошибки Excel."
    Else
        value = ConfigCellText(data(rowIndex, columnIndex), "значение", rowIndex, key, False, columnIndex)
    End If
    AppendSetting key, value, value, source, rowIndex, columnIndex, valueError
End Sub

' Добавляет имя настройки и значение в общие массивы. Для небольшого
' числа настроек Config поиск по ним достаточно быстр и не требует
' отдельных объектов на каждое значение.
Private Sub AppendSetting(ByVal keyValue As String, ByVal settingValue As String, _
        ByVal defaultValue As String, ByVal sourceValue As String, _
        Optional ByVal valueRow As Long = 0, Optional ByVal valueColumn As Long = 0, _
        Optional ByVal valueError As String = vbNullString)
    mCount = mCount + 1
    ReDim Preserve mKeys(1 To mCount)
    ReDim Preserve mValues(1 To mCount)
    ReDim Preserve mValueErrors(1 To mCount)
    ReDim Preserve mDefaults(1 To mCount)
    ReDim Preserve mSources(1 To mCount)
    ReDim Preserve mLocations(1 To mCount)
    mKeys(mCount) = keyValue
    mValues(mCount) = settingValue
    mValueErrors(mCount) = valueError
    mDefaults(mCount) = defaultValue
    mSources(mCount) = sourceValue
    If valueRow > 0 And valueColumn > 0 Then mLocations(mCount) = InputCellLocation(valueRow, valueColumn)
End Sub

' Фиксирует фактический адрес значения, а не предполагаемую позицию Config.
' Адрес используется только в диагностике: перестановка строк таблицы допустима.
Private Function InputCellLocation(ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If Len(mLoadingSheet) = 0 Or columnIndex < 1 Then
        InputCellLocation = "лист Config"
        Exit Function
    End If
    Dim column As Long, address As String
    column = mLoadingColumn + columnIndex - 1
    Do While column > 0
        address = Chr$(65 + (column - 1) Mod 26) & address
        column = (column - 1) \ 26
    Loop
    InputCellLocation = "лист " & mLoadingSheet & ", ячейка " & address & CStr(mLoadingRow + rowIndex - 1)
End Function

' Добавляет место исправления к ошибке конкретной настройки. Причину и
' допустимые значения определяет ее владелец; текст не назначает ResultCode/status.
' Для отсутствующего ключа адрес не выдумывается: указаны раздел и восстановление строки.
Public Function InputErrorMessage(ByVal key As String, ByVal reason As String, ByVal correction As String) As String
    Dim index As Long, location As String
    index = FindKey(key)
    location = InputLocation(key)
    InputErrorMessage = "Настройка " & key & " (" & location & "): " & reason
    If index > 0 Then
        If Len(mValues(index)) > 0 Then
            InputErrorMessage = InputErrorMessage & " Введённое значение: """ & mValues(index) & """."
        Else
            InputErrorMessage = InputErrorMessage & " Ячейка значения не заполнена."
        End If
    End If
    If Len(Trim$(correction)) > 0 Then InputErrorMessage = InputErrorMessage & " " & correction
End Function

' Возвращает сохранённое место задания настройки как строку. Расчётный класс
' может сохранить эту ссылку для последующей ошибки, не сохраняя объекты Excel.
Public Function InputLocation(ByVal key As String) As String
    Dim index As Long
    index = FindKey(key)
    If index > 0 Then InputLocation = mLocations(index)
    If Len(InputLocation) = 0 Then InputLocation = InputSettingsSection(key)
End Function

' Называет видимую таблицу для утраченной строки. Это навигация по Config,
' а не подстановка значения отсутствующей настройки или таблицы.
Private Function InputSettingsSection(ByVal key As String) As String
    Dim sectionName As String
    Select Case True
        Case Left$(key, 9) = "Concrete.": sectionName = "параметры бетона"
        Case Left$(key, 6) = "Steel.": sectionName = "параметры арматуры"
        Case Left$(key, 6) = "Units.": sectionName = "единицы измерения, столбцы INPUT/OUTPUT"
        Case Left$(key, 5) = "Sign.": sectionName = "система знаков, столбец USER"
        Case Left$(key, 10) = "SLS.Crack.": sectionName = "настройки расчета трещин"
        Case Left$(key, 10) = "Stability.": sectionName = "настройки устойчивости"
        Case Left$(key, 7) = "Solver.": sectionName = "настройки решателя"
        Case Left$(key, 9) = "Capacity.": sectionName = "настройки несущей способности"
        Case Left$(key, 15) = "AutoCAD.Import.": sectionName = "импорт геометрии из AutoCAD"
        Case Left$(key, 5) = "Plot.": sectionName = "оформление Excel-схемы"
        Case Left$(key, 9) = "Geometry.", Left$(key, 5) = "Mesh.": sectionName = "геометрия и расчетная сетка"
        Case Else: sectionName = "таблица исходных настроек"
    End Select
    InputSettingsSection = "лист Config, раздел """ & sectionName & """"
End Function

' Сбрасывает прочитанные ключи, значения, адреса и отложенные ошибки ввода.
' Следующее чтение начинается с пустого набора настроек.
Public Sub Clear()
    Erase mKeys
    Erase mValues
    Erase mValueErrors
    Erase mDefaults
    Erase mSources
    Erase mLocations
    mLoadingSheet = vbNullString
    mLoadingRow = 0
    mLoadingColumn = 0
    mCount = 0
End Sub

' Возвращает строковое значение настройки.
' Пустая строка и TODO трактуются как отсутствие пользовательского значения,
' поэтому вызывающий код получает свой defaultValue.
Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    AssertValueReadable index
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
    AssertValueReadable index
    If index = 0 Then
        GetRawString = defaultValue
    Else
        GetRawString = mValues(index)
    End If
End Function

' Выдает сохраненную ошибку активной ячейки до любых правил пустоты/default.
' Индекс 0 означает отсутствующий ключ; его обязательность проверяет вызывающий API.
Private Sub AssertValueReadable(ByVal index As Long)
    If index = 0 Then Exit Sub
    If Len(mValueErrors(index)) > 0 Then Err.Raise vbObjectError + 4309, _
        "CSystemSettingsReader", mValueErrors(index)
End Sub

' Возвращает Double с учетом десятичного разделителя Excel. Для optional
' параметров пустое значение допускается; allowBlankValue=False запрещает
' скрытый default для явно заданной пустой/TODO ячейки, но не для отсутствующего key.
Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#, _
        Optional ByVal allowBlankValue As Boolean = True) As Double
    Dim textValue As String, parsed As Double
    If Not allowBlankValue And HasKey(key) Then
        textValue = Trim$(GetRawString(key))
        If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
            Err.Raise vbObjectError + 4310, "CSystemSettingsReader.GetDouble", _
                InputErrorMessage(key, "Числовое значение не задано.", "Введите число или формулу, возвращающую число; пустое значение и TODO недопустимы.")
        End If
    Else
        textValue = GetString(key, vbNullString)
    End If
    If Len(textValue) = 0 Then
        GetDouble = defaultValue
    Else
        On Error GoTo InvalidNumber
        If Not TryReadStrictDouble(textValue, parsed) Then GoTo InvalidNumber
        GetDouble = parsed
    End If
    Exit Function

InvalidNumber:
    Err.Raise vbObjectError + 4312, "CSystemSettingsReader.GetDouble", _
        InputErrorMessage(key, "Не удалось прочитать число: """ & textValue & """.", _
            "Введите конечное число допустимого диапазона или исправьте формулу.")
End Function

' Читает обязательную числовую настройку без скрытой подстановки значения.
' Используется для обязательных исходных данных. Пустая ячейка и потерянная
' строка различаются в сообщении; обе причины остаются ошибкой ввода без default.
Public Function GetRequiredDouble(ByVal key As String) As Double
    Dim textValue As String, parsed As Double
    textValue = Trim$(GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        Err.Raise vbObjectError + 4310, "CSystemSettingsReader", _
            InputErrorMessage(key, "Обязательное числовое значение не задано.", RequiredInputCorrection(key, True))
    End If
    On Error GoTo InvalidNumber
    If Not TryReadStrictDouble(textValue, parsed) Then GoTo InvalidNumber
    GetRequiredDouble = parsed
    Exit Function

InvalidNumber:
    Err.Raise vbObjectError + 4313, "CSystemSettingsReader.GetRequiredDouble", _
        InputErrorMessage(key, "Не удалось прочитать обязательное число: """ & textValue & """.", _
            "Введите конечное число допустимого диапазона или исправьте формулу.")
End Function

' Читает обязательную текстовую настройку без default-значения.
' Нужна для пользовательских переключателей, которые являются частью
' расчетной материальной модели, например профиля арматуры.
Public Function GetRequiredString(ByVal key As String) As String
    Dim textValue As String
    textValue = Trim$(GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        Err.Raise vbObjectError + 4311, "CSystemSettingsReader", _
            InputErrorMessage(key, "Обязательное значение не задано.", RequiredInputCorrection(key, False))
    End If
    GetRequiredString = textValue
End Function

' Читает обязательный вариант из заданного списка и возвращает его принятое написание.
' Пустота, TODO и неизвестное значение дают ошибку с адресом ячейки,
' а не выбор первого элемента по умолчанию.
Public Function GetRequiredChoice(ByVal key As String, ByVal choices As Variant) As String
    Dim value As String, choice As Variant
    value = GetRequiredString(key)
    For Each choice In choices
        If StrComp(value, CStr(choice), vbTextCompare) = 0 Then
            GetRequiredChoice = CStr(choice)
            Exit Function
        End If
    Next choice
    Err.Raise vbObjectError + 4311, "CSystemSettingsReader.GetRequiredChoice", _
        InputErrorMessage(key, "Недопустимый вариант: """ & value & """.", "Выберите из списка: " & Join(choices, ", ") & ".")
End Function

' Подсказывает действие отдельно для утраченной строки и незаполненного значения.
' Число/текст проверяются здесь, а инженерный диапазон остается у потребителя.
Private Function RequiredInputCorrection(ByVal key As String, ByVal numericValue As Boolean) As String
    If Not HasKey(key) Then RequiredInputCorrection = "Восстановите строку " & key & " в указанном разделе. "
    If numericValue Then
        RequiredInputCorrection = RequiredInputCorrection & "Введите число или формулу, возвращающую число; пустое значение и TODO недопустимы."
    Else
        RequiredInputCorrection = RequiredInputCorrection & "Заполните значение или выберите допустимый вариант из списка; пустое значение и TODO недопустимы."
    End If
End Function

' Читает целое Long без округления дробной части и переполнения.
' Допустимость пустоты задаёт общий способ чтения числа;
' знак и инженерный диапазон проверяет использующий параметр расчёт.
Public Function GetLong(ByVal key As String, Optional ByVal defaultValue As Long = 0, _
        Optional ByVal allowBlankValue As Boolean = True) As Long
    GetLong = CheckedLong(key, GetDouble(key, CDbl(defaultValue), allowBlankValue))
End Function

' Читает обязательный счетчик без подстановки отсутствующего/пустого значения.
' Применяет ту же проверку представимости и отсутствия округления, что GetLong.
Public Function GetRequiredLong(ByVal key As String) As Long
    GetRequiredLong = CheckedLong(key, GetRequiredDouble(key))
End Function

' Проверяет числовой ввод перед CLng. Знак и допустимый инженерный диапазон
' остаются у владельца настройки; здесь исключены Overflow и округление дроби.
Private Function CheckedLong(ByVal key As String, ByVal value As Double) As Long
    If value < -2147483648# Or value > 2147483647# Then Err.Raise vbObjectError + 4315, "CSystemSettingsReader.GetLong", _
        InputErrorMessage(key, "Целое число выходит за допустимый диапазон Long.", "Введите целое число от -2147483648 до 2147483647 с учетом допустимого диапазона этой настройки.")
    If value <> Fix(value) Then Err.Raise vbObjectError + 4315, "CSystemSettingsReader.GetLong", _
        InputErrorMessage(key, "Введено дробное значение вместо целого; автоматическое округление не выполняется.", "Введите целое число.")
    CheckedLong = CLng(value)
End Function

' Читает переключатель Yes/No и поддерживаемые варианты True/False, 1/0, да/нет.
' Значение по умолчанию допустимо только при отсутствии ключа; пустой
' или неверный явный ввод остаётся ошибкой.
Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    If Not HasKey(key) Then
        GetBoolean = defaultValue
        Exit Function
    End If
    Dim textValue As String
    textValue = UCase$(Trim$(GetRawString(key)))
    Select Case textValue
        Case "YES", "TRUE", "1", "ДА": GetBoolean = True
        Case "NO", "FALSE", "0", "НЕТ": GetBoolean = False
        Case Else
            Err.Raise vbObjectError + 4314, "CSystemSettingsReader.GetBoolean", _
                InputErrorMessage(key, "Значение переключателя не распознано.", "Выберите Yes или No.")
    End Select
End Function

' Читает обязательный переключатель с теми же вариантами, что GetBoolean.
' Отсутствующий ключ и пустая ячейка не заменяются значением по умолчанию.
Public Function GetRequiredBoolean(ByVal key As String) As Boolean
    Call GetRequiredString(key)
    GetRequiredBoolean = GetBoolean(key)
End Function

' Читает пользовательский RGB без округления и молчаливой подстановки цвета.
' Каналы должны быть целыми 0..255; настоящий черный цвет возвращается как 0.
' Общая диагностика хранит фактический адрес, в том числе после переноса Name.
Public Function GetRequiredRgbColor(ByVal key As String) As Long
    Dim textValue As String, parts As Variant, channels(0 To 2) As Long
    Dim index As Long, value As Double
    textValue = GetRequiredString(key)
    On Error GoTo InvalidColor
    parts = Split(Replace$(textValue, ";", ","), ",")
    If UBound(parts) <> 2 Then GoTo InvalidColor
    For index = 0 To 2
        If Len(Trim$(CStr(parts(index)))) = 0 Then GoTo InvalidColor
        If Not TryReadStrictDouble(CStr(parts(index)), value) Then GoTo InvalidColor
        If value < 0# Or value > 255# Then GoTo InvalidColor
        If value <> Fix(value) Then GoTo InvalidColor
        channels(index) = CLng(value)
    Next index
    GetRequiredRgbColor = RGB(channels(0), channels(1), channels(2))
    Exit Function
InvalidColor:
    Err.Raise vbObjectError + 4316, "CSystemSettingsReader.GetRequiredRgbColor", _
        InputErrorMessage(key, "Цвет RGB задан неверно: """ & textValue & """.", _
            "Введите три целых числа от 0 до 255 через запятую, например 0,0,0 для черного цвета.")
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

' Возвращает число прочитанных ключей Config для отчёта и проверки структуры настроек.
Public Function KeyCount() As Long
    KeyCount = mCount
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

' Ищет первый ключ без учёта регистра и возвращает его индекс.
' Ноль означает отсутствие; повторы отдельно учитывает DuplicateCount.
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

' Возвращает ключ по порядковому номеру 1..KeyCount. Рабочий пакетный
' расчёт использует его для переноса адресов настроек в результаты ошибок.
Public Function KeyAt(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4300, "CSystemSettingsReader", "Индекс настройки вне допустимого диапазона." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mCount) & "."
    KeyAt = mKeys(index)
End Function


