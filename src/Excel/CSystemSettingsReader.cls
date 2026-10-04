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
Private mValueErrors() As String ' Отложенные ошибки геометрии/оформления: выдаются при чтении активного параметра, не заменяются default.
Private mDefaults() As String ' Значения по умолчанию из таблицы Config.
Private mSources() As String ' Комментарий/источник строки для диагностики.
Private mLocations() As String ' Лист, адрес и видимый столбец прочитанного значения; только диагностика ввода.
Private mLoadingSheet As String ' Имя листа текущей входной таблицы для диагностики.
Private mLoadingRow As Long ' Первая строка таблицы; считывается один раз, без COM на каждый ключ.
Private mLoadingColumn As Long ' Первый столбец таблицы для вычисления абсолютного адреса значения.
Private mCount As Long ' Количество загруженных ключей.
Private mMigrationWarning As String ' Предупреждение о конфликте старого и канонического ключа Extension.

' Загружает все настройки книги, нужные расчету и выводу.
' Метод объединяет разные именованные диапазоны Config в один внутренний
' словарь. Новые таблицы материалов имеют человекочитаемую форму, поэтому
' перед попаданием в словарь разворачиваются в ключи Concrete.* / Steel.*.
Public Sub LoadFromWorkbook(ByVal workbook As Object)
    Clear
    AppendRequiredNamedSettingsRange workbook, "rngUnitSettings", "Единицы измерения"
    AppendRequiredNamedSettingsRange workbook, "rngSignConventionSettings", "Система знаков"
    ValidateWorkbookUnitAndSignKeys
    AppendRequiredNamedSettingsRange workbook, "rngSystemSettings", "Системные настройки"
    AppendNamedSettingsRangeIfExists workbook, "rngConcreteMaterialParameters"
    AppendNamedSettingsRangeIfExists workbook, "rngSteelMaterialParameters"
    AppendRequiredNamedSettingsRange workbook, "rngPlotAnnotationSettings", "Подписи и размеры схемы"
    AppendGeometrySettingsRanges workbook
    NormalizeDiagramExtensionSetting True
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

' Проверяет обязательные пользовательские ключи полной книги после чтения
' обеих таблиц. Поврежденная или отсутствующая строка Config не становится
' optional default адаптера; автономный LoadFromRange остается неполным reader-ом.
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

' Подключает необязательный для reader диапазон, если его имя есть в книге.
' Отсутствующая таблица допустима до запроса материала потребителем, но
' существующее поврежденное имя не должно скрываться как отсутствие ввода.
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
    NormalizeDiagramExtensionSetting
End Sub

' Читает старый формат только на границе Excel. До любых defaults сохраняет
' явно введенное значение, затем удаляет старый ключ из runtime-настроек.
' Конфликт не скрывается: новый ключ имеет приоритет, причина доступна отчету.
Private Sub NormalizeDiagramExtensionSetting(Optional ByVal requireWorkbookValue As Boolean = False)
    Const canonicalKey As String = "General.DiagramExtension"
    Const legacyKey As String = "Solver.DirectState.DiagramExtension"
    Dim canonicalIndex As Long
    Dim legacyIndex As Long
    Dim valueText As String
    Dim legacyValue As String
    canonicalIndex = FindKey(canonicalKey)
    legacyIndex = FindKey(legacyKey)
    If canonicalIndex > 0 Then
        valueText = DiagramExtensionValue(mValues(canonicalIndex), canonicalKey)
        If legacyIndex > 0 Then
            legacyValue = UCase$(Trim$(mValues(legacyIndex)))
            Select Case legacyValue
                Case "YES", "TRUE", "1", "ДА": legacyValue = "Yes"
                Case "NO", "FALSE", "0", "НЕТ": legacyValue = "No"
            End Select
            If StrComp(valueText, legacyValue, vbTextCompare) <> 0 Then
                mMigrationWarning = "Настройки численного продолжения диаграмм заданы дважды. " & _
                    "Использовано значение " & valueText & " из General.DiagramExtension; " & _
                    "старый Solver.DirectState.DiagramExtension больше не управляет расчетом."
            End If
        End If
        mValues(canonicalIndex) = valueText
    ElseIf legacyIndex > 0 Then
        valueText = DiagramExtensionValue(mValues(legacyIndex), legacyKey)
        mKeys(legacyIndex) = canonicalKey
        mValues(legacyIndex) = valueText
        legacyIndex = 0
    Else
        If requireWorkbookValue Then Err.Raise vbObjectError + 4313, "CSystemSettingsReader", _
            InputErrorMessage(canonicalKey, "Обязательная строка настройки отсутствует.", _
                "Восстановите строку в системных настройках Config и выберите Yes или No.")
        ' Неполный автономный LoadFromRange сохраняет контракт reader-а;
        ' в полной пользовательской книге отсутствие строки является ошибкой.
        AppendSetting canonicalKey, "Yes", "Yes", "Численное продолжение активных диаграмм."
    End If
    If legacyIndex > 0 Then RemoveSettingAt legacyIndex
End Sub

' Возвращает каноническое Yes/No, сохраняя поддержанные логические формы.
' Пустое или неизвестное явно введенное значение является ошибкой Config.
Private Function DiagramExtensionValue(ByVal rawValue As String, ByVal settingKey As String) As String
    Select Case UCase$(Trim$(rawValue))
        Case "YES", "TRUE", "1", "ДА": DiagramExtensionValue = "Yes"
        Case "NO", "FALSE", "0", "НЕТ": DiagramExtensionValue = "No"
        Case Else
            Err.Raise vbObjectError + 4310, "CSystemSettingsReader", _
                InputErrorMessage(settingKey, "Значение переключателя не распознано.", "Выберите Yes или No.")
    End Select
End Function

' Убирает мигрированный старый ключ из уже прочитанных массивов. Остальные
' настройки сохраняют порядок и значения; runtime не видит двух переключателей.
Private Sub RemoveSettingAt(ByVal index As Long)
    Dim i As Long
    For i = index To mCount - 1
        mKeys(i) = mKeys(i + 1)
        mValues(i) = mValues(i + 1)
        mValueErrors(i) = mValueErrors(i + 1)
        mDefaults(i) = mDefaults(i + 1)
        mSources(i) = mSources(i + 1)
        mLocations(i) = mLocations(i + 1)
    Next i
    mCount = mCount - 1
End Sub

Public Property Get MigrationWarning() As String
    MigrationWarning = mMigrationWarning
End Property

' Добавляет одну непрерывную таблицу настроек; специальные геометрические
' форматы распознаются по шапке. Несмежные области отвергаются до Value2:
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

    ' Имя известных таблиц сохраняет их смысл при повреждении шапки.
    ' Сначала называем конкретную ячейку структуры, затем читаем значения;
    ' такой дефект не должен выглядеть как исчезновение обязательных ключей.
    Select Case LCase$(rangeName)
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

    ' Named-таблица сохраняет свой тип даже после утраты строки параметра.
    ' Иначе отсутствие Placement превращало все поля в чужую scalar-таблицу.
    If StrComp(rangeName, "rngPlotAnnotationSettings", vbTextCompare) = 0 Then
        If columnCount < 5 Then Err.Raise vbObjectError + 4316, "CSystemSettingsReader", _
            "Таблица аннотаций схемы на листе " & mLoadingSheet & " имеет неполную структуру. Восстановите колонки параметров, арматуры, размеров, единиц и комментария."
        AppendPlotAnnotationSettings data
        Exit Sub
    End If

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
            If IsGeometrySettingKey(keyValue) Then
                AppendDeferredInputValue data, keyValue, rowIndex, 2, sourceValue
            ElseIf Left$(keyValue, 5) = "Plot." Then
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

' Проверяет только смысловые колонки известных именованных таблиц.
' Ожидаемые подписи совпадают с прежним распознаванием формата; изменено
' место диагностики, а не физический смысл параметров. Адрес вычисляется
' от текущей ссылки Name, без привязки к координатам шаблона Config.
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
        tensionValue = ConfigCellText(data(rowIndex, 3), "Растяжение", rowIndex, rowKey, False, 3)
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

' Разворачивает таблицу аннотаций в обычные ключи Plot.RebarLabels.*
' и Plot.Dimensions.*, чтобы остальной код читал настройки единообразно.
Private Sub AppendPlotAnnotationSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim settingName As String
        settingName = ConfigCellText(data(rowIndex, 1), "параметр аннотации", rowIndex, vbNullString, False, 1)
        If Len(settingName) > 0 And Left$(settingName, 1) <> "[" Then
            AppendPlotAnnotationValue data, "Plot.RebarLabels." & settingName, rowIndex, 2
            AppendPlotAnnotationValue data, "Plot.Dimensions." & settingName, rowIndex, 3
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

' Загружает INPUT/INTERNAL/OUTPUT единицы в общий словарь настроек.
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
    AppendDeferredInputValue data, "HollowRectangle.H", valueRow, 1
    AppendDeferredInputValue data, "HollowRectangle.B", valueRow, 2
    AppendDeferredInputValue data, "HollowRectangle.R", valueRow, 3
    AppendDeferredInputValue data, "HollowRectangle.OpeningH", valueRow, 4
    AppendDeferredInputValue data, "HollowRectangle.OpeningB", valueRow, 5
    AppendDeferredInputValue data, "HollowRectangle.OpeningR", valueRow, 6

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
            AppendDeferredInputValue data, keyValue, rowIndex, 2, sourceValue
        End If
    Next rowIndex
End Sub

' Читает основной ряд по наружным граням и граням отверстия. Пользовательское
' количество наружных стержней сохраняется без предварительного отбора чисел:
' ошибочный текст должен дойти до строгого GetLong, а не выключить грань.
' Количество на гранях отверстия служебное и определяется проекцией арматуры.
Private Sub AppendHollowRectangleMainRebars(ByRef data As Variant, ByVal mainHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = mainHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = HollowRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".as", rowIndex, 2, Trim$(CStr(data(rowIndex, 6)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".d", rowIndex, 3, Trim$(CStr(data(rowIndex, 6)))
        If Left$(faceKey, 8) <> "Opening." Then
            Dim countKey As String
            countKey = "HollowRectangle." & faceKey & ".n"
            AppendDeferredInputValue data, countKey, rowIndex, 4, Trim$(CStr(data(rowIndex, 6)))
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
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".d_2row", rowIndex, 2, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".loc_2row", rowIndex, 3, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".bind_2row", rowIndex, 4, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".d_3row", rowIndex, 5, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".loc_3row", rowIndex, 6, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "HollowRectangle." & faceKey & ".bind_3row", rowIndex, 7, Trim$(CStr(data(rowIndex, 9)))
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
    AppendDeferredInputValue data, "RoundedRectangle.B", valueRow, 2
    AppendDeferredInputValue data, "RoundedRectangle.H", valueRow, 3

    Dim sideHeaderRow As Long
    sideHeaderRow = FindRowWithCaption(data, "слева", 2, geometryHeaderRow + 1)
    If sideHeaderRow > 0 And UBound(data, 1) >= sideHeaderRow + 4 Then
        AppendDeferredInputValue data, "RoundedRectangle.Left.Type", sideHeaderRow + 1, 2
        AppendDeferredInputValue data, "RoundedRectangle.Right.Type", sideHeaderRow + 1, 3
        AppendDeferredInputValue data, "RoundedRectangle.Left.W", sideHeaderRow + 2, 2
        AppendDeferredInputValue data, "RoundedRectangle.Right.W", sideHeaderRow + 2, 3
        AppendDeferredInputValue data, "RoundedRectangle.Left.R1", sideHeaderRow + 3, 2
        AppendDeferredInputValue data, "RoundedRectangle.Right.R1", sideHeaderRow + 3, 3
        AppendDeferredInputValue data, "RoundedRectangle.Left.R2", sideHeaderRow + 4, 2
        AppendDeferredInputValue data, "RoundedRectangle.Right.R2", sideHeaderRow + 4, 3
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
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".as", rowIndex, 2, Trim$(CStr(data(rowIndex, 6)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".d", rowIndex, 3, Trim$(CStr(data(rowIndex, 6)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".n", rowIndex, 4, Trim$(CStr(data(rowIndex, 6)))
    Next rowIndex
End Sub

' Читает дополнительные ряды RoundedRectangle и сохраняет привязку каждого ряда к своей грани.
Private Sub AppendRoundedRectangleExtraRebars(ByRef data As Variant, ByVal extraHeaderRow As Long)
    Dim rowIndex As Long
    For rowIndex = extraHeaderRow + 1 To UBound(data, 1)
        Dim faceKey As String
        faceKey = RoundedRectangleFaceKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(faceKey) = 0 Then Exit For
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".d_2row", rowIndex, 2, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".loc_2row", rowIndex, 3, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".bind_2row", rowIndex, 4, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".d_3row", rowIndex, 5, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".loc_3row", rowIndex, 6, Trim$(CStr(data(rowIndex, 9)))
        AppendDeferredInputValue data, "RoundedRectangle." & faceKey & ".bind_3row", rowIndex, 7, Trim$(CStr(data(rowIndex, 9)))
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
        AppendDeferredInputValue data, "RectSet." & faceName, 2, col, Trim$(CStr(data(2, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".as_1", 3, col, Trim$(CStr(data(3, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".as_2", 4, col, Trim$(CStr(data(4, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".d_1", 5, col, Trim$(CStr(data(5, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".d_2", 6, col, Trim$(CStr(data(6, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".n_1", 7, col, Trim$(CStr(data(7, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".n_2", 8, col, Trim$(CStr(data(8, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".StartOffset1", 9, col, Trim$(CStr(data(9, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".EndOffset1", 10, col, Trim$(CStr(data(10, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".StartOffset2", 11, col, Trim$(CStr(data(11, 7)))
        AppendDeferredInputValue data, "RectSet." & faceName & ".EndOffset2", 12, col, Trim$(CStr(data(12, 7)))
        If UBound(data, 1) >= 19 Then
            AppendDeferredInputValue data, "RectSet." & faceName & ".d_2row_1", 14, col, Trim$(CStr(data(14, 7)))
            AppendDeferredInputValue data, "RectSet." & faceName & ".d_2row_2", 15, col, Trim$(CStr(data(15, 7)))
            AppendDeferredInputValue data, "RectSet." & faceName & ".d_3row_1", 16, col, Trim$(CStr(data(16, 7)))
            AppendDeferredInputValue data, "RectSet." & faceName & ".d_3row_2", 17, col, Trim$(CStr(data(17, 7)))
            AppendDeferredInputValue data, "RectSet." & faceName & ".loc_2row", 18, col, Trim$(CStr(data(18, 7)))
            AppendDeferredInputValue data, "RectSet." & faceName & ".loc_3row", 19, col, Trim$(CStr(data(19, 7)))
            If UBound(data, 1) >= 21 Then
                AppendDeferredInputValue data, "RectSet." & faceName & ".bind_2row", 20, col, Trim$(CStr(data(20, 7)))
                AppendDeferredInputValue data, "RectSet." & faceName & ".bind_3row", 21, col, Trim$(CStr(data(21, 7)))
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
            AppendDeferredInputValue data, keyValue, rowIndex, 2, sourceValue
        End If
    Next rowIndex
End Sub

' Читает текущую табличную компоновку RectSet-сечения:
' отдельный блок геометрии, блок основного армирования и блок дополнительных рядов.
Private Sub AppendRectSetSectionedSettings(ByRef data As Variant, Optional ByVal geometryHeaderRow As Long = 2)
    Dim valueRow As Long
    valueRow = geometryHeaderRow + 1
    AppendDeferredInputValue data, "RectSet.H1", valueRow, 1
    AppendDeferredInputValue data, "RectSet.B1", valueRow, 2
    AppendDeferredInputValue data, "RectSet.H2", valueRow, 3
    AppendDeferredInputValue data, "RectSet.B2", valueRow, 4

    Dim faceNames(1 To 8) As String
    Dim sideIndexes(1 To 8) As Long
    FillRectSetSectionedFaceOrder faceNames, sideIndexes

    Dim i As Long
    ' Восемь строк граней превращаются в ключи с суффиксами _1/_2.
    For i = 1 To 8
        Dim rowIndex As Long
        rowIndex = geometryHeaderRow + 3 + i
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".as_" & CStr(sideIndexes(i)), rowIndex, 2, Trim$(CStr(data(rowIndex, 8)))
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".d_" & CStr(sideIndexes(i)), rowIndex, 3, Trim$(CStr(data(rowIndex, 8)))
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".n_" & CStr(sideIndexes(i)), rowIndex, 4, Trim$(CStr(data(rowIndex, 8)))
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".StartOffset" & CStr(sideIndexes(i)), rowIndex, 5, Trim$(CStr(data(rowIndex, 8)))
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".EndOffset" & CStr(sideIndexes(i)), rowIndex, 6, Trim$(CStr(data(rowIndex, 8)))

        Dim extraRow As Long
        extraRow = geometryHeaderRow + 13 + i
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".d_2row_" & CStr(sideIndexes(i)), extraRow, 2, Trim$(CStr(data(extraRow, 9)))
        AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".d_3row_" & CStr(sideIndexes(i)), extraRow, 5, Trim$(CStr(data(extraRow, 9)))
        If sideIndexes(i) = 1 Then
            ValidateRectSetSharedSelectors data, extraRow, faceNames(i)
            AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".loc_2row", extraRow, 3, Trim$(CStr(data(extraRow, 9)))
            AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".bind_2row", extraRow, 4, Trim$(CStr(data(extraRow, 9)))
            AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".loc_3row", extraRow, 6, Trim$(CStr(data(extraRow, 9)))
            AppendDeferredInputValue data, "RectSet." & faceNames(i) & ".bind_3row", extraRow, 7, Trim$(CStr(data(extraRow, 9)))
        End If
    Next i
End Sub

' Проверяет общие для пары граней параметры дополнительных рядов RectSet.
' Повторное значение во второй строке допускается только совпадающее с первой;
' пустая служебная ячейка или "-" не задает отдельный выбор. Конфликт нельзя
' молча проигнорировать, поскольку независимый выбор по сторонам не поддержан.
Private Sub ValidateRectSetSharedSelectors(ByRef data As Variant, ByVal firstRow As Long, ByVal faceName As String)
    Dim column As Variant, firstValue As String, secondValue As String, parameterName As String
    Dim faceLabel As String, key As String
    Select Case UCase$(faceName)
        Case "H1": faceLabel = "вертикальных граней верхнего прямоугольника (H1)"
        Case "B1": faceLabel = "горизонтальных граней верхнего прямоугольника (B1)"
        Case "H2": faceLabel = "вертикальных граней нижнего прямоугольника (H2)"
        Case "B2": faceLabel = "горизонтальных граней нижнего прямоугольника (B2)"
        Case Else: faceLabel = "граней " & faceName
    End Select
    For Each column In Array(3, 4, 6, 7)
        Select Case CLng(column)
            Case 3: parameterName = "положение второго ряда": key = "loc_2row"
            Case 4: parameterName = "привязка второго ряда к первому": key = "bind_2row"
            Case 6: parameterName = "положение третьего ряда": key = "loc_3row"
            Case 7: parameterName = "привязка третьего ряда к первому": key = "bind_3row"
        End Select
        key = "RectSet." & faceName & "." & key
        ' Ошибку формулы основной ячейки проверит активный потребитель ряда.
        ' Проверка противоречивого отдельного ввода второй стороны сохраняется.
        If Not IsError(data(firstRow, CLng(column))) Then
            firstValue = ConfigCellText(data(firstRow, CLng(column)), parameterName, firstRow, key, False, CLng(column))
            secondValue = ConfigCellText(data(firstRow + 1, CLng(column)), parameterName, firstRow + 1, key, False, CLng(column))
            If Len(secondValue) > 0 And secondValue <> "-" Then
                If StrComp(firstValue, secondValue, vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 4317, "CSystemSettingsReader", _
                        "В таблице RectSet для " & faceLabel & " параметр """ & _
                        parameterName & """ задан по-разному в двух строках (" & _
                        InputCellLocation(firstRow, CLng(column)) & "; " & _
                        InputCellLocation(firstRow + 1, CLng(column)) & "). Это общий выбор для пары сторон; " & _
                        "укажите его в первой строке, а во второй оставьте тот же выбор или пустое значение."
                End If
            End If
        End If
    Next column
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

' Сохраняет значение геометрии или оформления вместе с фактическим адресом.
' Ошибка формулы выдается при обращении активного потребителя: чужая форма,
' выключенный ряд или группа аннотаций не мешают основному расчету, но чтение
' ошибочного параметра никогда не возвращает пустоту или default.
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

' Распознает геометрические ключи обычной двухколоночной компоновки, включая
' таблицу круга. Условия активности определяют builders, а не Excel-reader.
Private Function IsGeometrySettingKey(ByVal key As String) As Boolean
    Select Case LCase$(Left$(key, InStr(1, key & ".", ".", vbBinaryCompare) - 1))
        Case "circle", "rebar", "rectset", "roundedrectangle", "hollowrectangle"
            IsGeometrySettingKey = True
    End Select
End Function

' Добавляет одну пару ключ-значение в reader.
' Массивы параллельные: это проще и быстрее для VBA, чем словарь с Variant-объектами,
' а поиск по ключу остается достаточно дешевым для числа настроек Config.
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
    If index > 0 Then location = mLocations(index)
    If Len(location) = 0 Then location = InputSettingsSection(key)
    InputErrorMessage = "Настройка " & key & " (" & location & "): " & reason & " " & correction
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
        Case Left$(key, 7) = "Circle.", Left$(key, 6) = "Rebar.": sectionName = "круглое сечение"
        Case Left$(key, 8) = "RectSet.": sectionName = "RectSet-сечение"
        Case Left$(key, 17) = "RoundedRectangle.": sectionName = "скругленный прямоугольник"
        Case Left$(key, 16) = "HollowRectangle.": sectionName = "пустотелый прямоугольник"
        Case Left$(key, 9) = "Geometry.", Left$(key, 5) = "Mesh.": sectionName = "геометрия и расчетная сетка"
        Case Else: sectionName = "таблица исходных настроек"
    End Select
    InputSettingsSection = "лист Config, раздел """ & sectionName & """"
End Function

' Сбрасывает все прочитанные настройки и точки диаграмм.
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
    mMigrationWarning = vbNullString
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
    Dim textValue As String
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
        GetDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
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
    Dim textValue As String
    textValue = Trim$(GetRawString(key, vbNullString))
    If Len(textValue) = 0 Or UCase$(textValue) = "TODO" Then
        Err.Raise vbObjectError + 4310, "CSystemSettingsReader", _
            InputErrorMessage(key, "Обязательное числовое значение не задано.", RequiredInputCorrection(key, True))
    End If
    On Error GoTo InvalidNumber
    GetRequiredDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
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

' Читает обязательный селектор из явного набора значений. Возвращает запись
' из набора с каноническим регистром; пустота/TODO и неизвестный вариант
' дают ошибку ввода с адресом ячейки, а не выбор первого варианта по умолчанию.
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

' Читает Long без округления дроби и переполнения. Контракт пустой ячейки
' передается общему числовому reader-у; знак и инженерный диапазон проверяет consumer.
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

' Читает Yes/No и утвержденные aliases. Default допустим только для отсутствующего
' optional key; явное пустое/невалидное значение не превращается в False/default.
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

' Требует присутствующее значение переключателя и сохраняет утвержденные
' aliases Yes/No. Потерянная строка и пустая ячейка не выбирают default.
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
        value = CDbl(Trim$(CStr(parts(index))))
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






