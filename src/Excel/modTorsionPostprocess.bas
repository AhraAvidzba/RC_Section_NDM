Attribute VB_Name = "modTorsionPostprocess"
Option Explicit

' ==========================================================================
' Адаптер изолированной постпроверки T для Excel и файлового снимка
' ==========================================================================
' Единственный вход в этап — готовые геометрия и результаты batch. Ошибки
' пользовательских таблиц T становятся строками блока NMT, а не ошибками
' НДМ. Исходные настройки, семь таблиц НДМ и существующие сводки не изменяются.

' Читает только запрошенные строки T и сохраняет самостоятельные матрицы.
' Отсутствие таблиц T / NMT в книге означает отсутствие постпроверки.
Public Function CaptureTorsionPostprocess(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem, ByVal snapshot As CCalculationSnapshot) As String
    On Error GoTo Failed
    CaptureTorsionPostprocess = CaptureTorsionPostCore(workbook, section, batch, units, snapshot)
    Exit Function
Failed:
    CaptureTorsionPostprocess = "Не удалось сохранить постпроверку T (кручение): " & WorkbookDialogText(Err.Description, workbook) & _
        vbCrLf & "Расчёт и сохранение результатов НДМ продолжаются независимо."
End Function

' Отделяет обработку строк от внешней границы отказа: даже ошибка записи
' диагностической строки возвращается предупреждением вызывающему НДМ.
Private Function CaptureTorsionPostCore(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem, ByVal snapshot As CCalculationSnapshot) As String
    Dim source As Object, settings As CSystemSettingsReader
    Dim dictionary As Object, requested As Object, results As New Collection
    Dim torsionWriter As New CTorsionSummaryWriter
    Dim data As Variant, main As Variant, identifier As String, row As Long
    Dim result As CTorsionPostResult, calculator As CTorsionPostCalculator, preparationError As String
    On Error GoTo TableFailed
    Set source = OptionalTorsionRange(workbook, "rngTorsionCombinations")
    If source Is Nothing Then Exit Function
    If source.Areas.Count <> 1 Or source.Columns.Count <> 2 Or source.Rows.Count < 2 Then _
        Err.Raise 5, , "rngTorsionCombinations должен ссылаться на одну таблицу ID, T с заголовком и строками данных."
    data = source.Value2
    Set dictionary = CreateObject("Scripting.Dictionary"): dictionary.CompareMode = vbTextCompare
    Set requested = CreateObject("Scripting.Dictionary"): requested.CompareMode = vbTextCompare
    main = workbook.Names.Item("rngLoadCombinations").RefersToRange.Value2
    For row = 2 To UBound(main, 1)
        If Not IsError(main(row, 1)) And Not IsNull(main(row, 1)) Then
            identifier = Trim$(CStr(main(row, 1)))
            If Len(identifier) > 0 Then dictionary(identifier) = row
        End If
    Next row
    For row = 2 To UBound(data, 1)
        If TorsionRowRequested(data(row, 1), data(row, 2)) Then
            If Not IsError(data(row, 1)) And Not IsNull(data(row, 1)) Then
                identifier = Trim$(CStr(data(row, 1)))
                If requested.Exists(identifier) Then requested(identifier) = -1 Else requested.Add identifier, row
            End If
        End If
    Next row
    For row = 2 To UBound(data, 1)
        If TorsionRowRequested(data(row, 1), data(row, 2)) Then
            Set result = ReadTorsionPostRow(workbook, section, batch, units, source, data, row, main, dictionary, requested, _
                calculator, settings, preparationError)
            If Not result Is Nothing Then results.Add result
        End If
    Next row
    torsionWriter.CaptureSummary snapshot, results, units
    Dim summary As Variant, warnings As Long, policy As New CResultStatusPolicy
    For Each result In results
        summary = result.Summary
        If CStr(summary(1, 38)) <> policy.OK Then warnings = warnings + 1
    Next result
    If warnings > 0 Then CaptureTorsionPostCore = "Постпроверка T (кручение): строк с отказом, ошибкой или недоступной суммой — " & CStr(warnings) & ". Причины указаны в блоке NMT на Results."
    Exit Function
TableFailed:
    Dim reason As String: reason = Err.Description
    If Not source Is Nothing Then reason = reason & vbCrLf & source.Worksheet.Name & "!" & source.Address(False, False)
    Set result = New CTorsionPostResult
    result.InitializeFailure "[T]", WorkbookDialogText("Постпроверка T (кручение)." & vbCrLf & reason, workbook)
    results.Add result
    torsionWriter.CaptureSummary snapshot, results, units
    CaptureTorsionPostCore = "Постпроверка T (кручение) завершена с ошибкой ввода; причина выведена в отдельном блоке Results. Результаты НДМ сохранены."
End Function

' Определяет применимость строки до проверки ID и чтения настроек.
' Нулевой момент полностью отключает строку, включая проверку повторного ID.
Private Function TorsionRowRequested(ByVal identifier As Variant, ByVal torque As Variant) As Boolean
    If TorsionValueBlank(identifier) Or TorsionValueBlank(torque) Then Exit Function
    If Not IsError(torque) And Not IsNull(torque) Then
        If IsNumeric(torque) Then
            If CDbl(torque) = 0# Then Exit Function
        End If
    End If
    TorsionRowRequested = True
End Function

' Превращает дефект одной строки в InputErr, сохраняя выполнение других ID.
Private Function ReadTorsionPostRow(ByVal workbook As Object, ByVal section As CSectionModel, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal source As Object, ByRef data As Variant, ByVal row As Long, ByRef main As Variant, _
        ByVal knownIDs As Object, ByVal requested As Object, ByRef calculator As CTorsionPostCalculator, _
        ByRef settings As CSystemSettingsReader, ByRef preparationError As String) As CTorsionPostResult
    Dim identifier As String, torque As Double, result As CTorsionPostResult, state As CSectionStateResult
    Dim column As Long: column = 1
    On Error GoTo Failed
    If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then Err.Raise 5, , "ID сочетания возвращает ошибку Excel."
    identifier = Trim$(CStr(data(row, 1)))
    If Not knownIDs.Exists(identifier) Then Err.Raise 5, , "ID не найден в rngLoadCombinations. Выберите ID существующего сочетания."
    If CLng(requested(identifier)) = -1 Then Err.Raise 5, , "Для одного ID задано несколько строк T. Оставьте одну строку с итоговым моментом."
    column = 2
    If IsError(data(row, 2)) Or IsNull(data(row, 2)) Then Err.Raise 5, , "Момент T возвращает ошибку Excel."
    If Not IsNumeric(data(row, 2)) Then Err.Raise 5, , "Момент T должен быть числом."
    torque = CDbl(data(row, 2))
    If Abs(torque) >= 1E+100 Then Err.Raise 5, , "Модуль T превышает допустимый числовой диапазон."
    If torque = 0# Then Exit Function ' Нулевая строка не читает параметры хомута и не запрашивает НДС.
    torque = Abs(units.InputMomentMxToInternal(torque)) ' Для независимой проверки используется модуль T; знаки изгиба не влияют на результат.
    If calculator Is Nothing And Len(preparationError) = 0 Then PrepareTorsionPostCalculator workbook, section, units, calculator, settings, preparationError
    If Len(preparationError) > 0 Then Err.Raise 5, , preparationError
    Dim i As Long, stateName As String
    stateName = settings.GetRequiredString("Torsion.StateType")
    For i = 1 To batch.Count
        If StrComp(batch.CombinationID(i), identifier, vbTextCompare) = 0 Then
            Set state = batch.ResultAt(i).StateRepository.FindState(SectionStateTypeFromText(stateName))
            Exit For
        End If
    Next i
    Set result = calculator.Calculate(identifier, torque, state)
    Dim values As Variant, bars As Variant, addressed As CTorsionPostResult
    values = result.Summary
    If Len(CStr(values(1, 39))) > 0 Then
        values(1, 39) = WorkbookDialogText("Постпроверка T (кручение)." & vbCrLf & settings.InputErrorMessage("Torsion.StateType", CStr(values(1, 39)), _
            "Выберите сохранённое состояние или включите его получение в расчётном профиле сочетания."), workbook)
        bars = result.Bars
        Set addressed = New CTorsionPostResult: addressed.Initialize values, bars ' Адресная копия не меняет готовое НДС или неизменяемый результат вычислителя.
        Set result = addressed
    End If
    Set ReadTorsionPostRow = result
    Exit Function
Failed:
    Dim reason As String: reason = InputDiagnosticText(Err.Description, _
        source.Worksheet.Name & "!" & source.Cells(row, column).Address(False, False), _
        "Постпроверка T (кручение), ID «" & identifier & "»", data(row, column))
    reason = WorkbookDialogText(reason, workbook)
    Set result = New CTorsionPostResult
    result.InitializeFailure identifier, reason
    Set ReadTorsionPostRow = result
End Function

' Подготавливает автономный reader и числовые настройки; неиспользуемые
' параметры отключённой проверки бетона не запрашиваются.
Private Sub PrepareTorsionPostCalculator(ByVal workbook As Object, ByVal section As CSectionModel, ByVal units As CUnitSystem, _
        ByRef calculator As CTorsionPostCalculator, ByRef settings As CSystemSettingsReader, ByRef failure As String)
    Dim table As Object, p As TTorsionPostParameters
    On Error GoTo Failed
    Set table = OptionalTorsionRange(workbook, "rngTorsionSettings")
    If table Is Nothing Then Err.Raise 5, , "Не найдена таблица rngTorsionSettings. Восстановите таблицу настроек T из сборки."
    Set settings = New CSystemSettingsReader: settings.LoadFromRange table
    p.StateName = settings.GetRequiredChoice("Torsion.StateType", Array("StrengthState", "CapacityState", "PreCrackState", "PostCrackState", "CrackedState"))
    p.AxisOffset = units.InputLengthToInternal(TorsionSettingNumber(settings, "Torsion.HoopAxisOffset", 0#, True))
    p.HoopDiameter = units.InputLengthToInternal(TorsionSettingNumber(settings, "Torsion.HoopDiameter", 0#, True))
    p.HoopSpacing = units.InputLengthToInternal(TorsionSettingNumber(settings, "Torsion.HoopSpacing", 0#, True))
    p.SteelResistance = units.InputStressToInternal(TorsionSettingNumber(settings, "Torsion.Rs", 0#, True))
    p.HoopResistance = units.InputStressToInternal(TorsionSettingNumber(settings, "Torsion.Rsw", 0#, True))
    p.BarBand = units.InputLengthToInternal(TorsionSettingNumber(settings, "Torsion.BarBand", 0#, True))
    p.StrutAngle = TorsionSettingNumber(settings, "Torsion.StrutAngle", 0#, True, 90#, True)
    p.ConcreteMethod = settings.GetRequiredChoice("Torsion.ConcreteMethod", Array("Strut", "GorskyRectangle", "None"))
    If p.ConcreteMethod <> "None" Then p.ConcreteResistance = units.InputStressToInternal(TorsionSettingNumber(settings, "Torsion.Rb", 0#, True))
    If p.ConcreteMethod = "Strut" Then
        p.WallThickness = units.InputLengthToInternal(TorsionSettingNumber(settings, "Torsion.WallThickness", 0#, False))
        p.ConcreteEfficiency = TorsionSettingNumber(settings, "Torsion.ConcreteEfficiency", 0#, True, 1#)
    ElseIf p.ConcreteMethod = "GorskyRectangle" Then
        p.GorskyCoefficient = TorsionSettingNumber(settings, "Torsion.GorskyCoefficient", 0#, True)
    End If
    Set calculator = New CTorsionPostCalculator
    calculator.Prepare section, p
    Exit Sub
Failed:
    failure = Err.Description
    If Not settings Is Nothing And Not calculator Is Nothing Then
        Dim key As Variant, primary As Boolean: primary = True
        For Each key In Split(calculator.FailureSettingKeys, ",")
            If Len(CStr(key)) > 0 Then
                If primary Then
                    failure = settings.InputErrorMessage(CStr(key), failure, "Проверьте параметры геометрии T.")
                    primary = False
                Else
                    failure = failure & vbCrLf & settings.InputErrorMessage(CStr(key), "Связанный параметр геометрии T.", vbNullString)
                End If
            End If
        Next key
    End If
    Set calculator = Nothing
End Sub

' Проверяет конечное число и границы с адресной диагностикой общего reader.
Private Function TorsionSettingNumber(ByVal settings As CSystemSettingsReader, ByVal key As String, ByVal minimum As Double, _
        ByVal strictMinimum As Boolean, Optional ByVal maximum As Double = 1E+50, Optional ByVal strictMaximum As Boolean = False) As Double
    Dim value As Double: value = settings.GetRequiredDouble(key)
    If value < minimum Or value > maximum Or (strictMinimum And value = minimum) Or (strictMaximum And value = maximum) Then _
        Err.Raise 5, , settings.InputErrorMessage(key, "Число вне допустимого диапазона " & CStr(minimum) & "…" & CStr(maximum) & ".", "Проверьте границы в комментарии настройки.")
    TorsionSettingNumber = value
End Function

' Находит необязательную таблицу, не подменяя неправильную ссылку другой книгой.
Private Function OptionalTorsionRange(ByVal workbook As Object, ByVal name As String) As Object
    Dim item As Object, found As Object
    On Error Resume Next
    Set item = workbook.Names.Item(name)
    On Error GoTo 0
    If item Is Nothing Then Exit Function
    On Error GoTo BrokenReference
    Set found = item.RefersToRange
    If Not found.Worksheet.Parent Is workbook Then Err.Raise 5, , "Таблица " & name & " ссылается на другую книгу."
    Set OptionalTorsionRange = found
    Exit Function
BrokenReference:
    Err.Raise 5, , "Не удаётся прочитать таблицу " & name & " текущей книги. Проверьте ссылку в диспетчере имён. " & Err.Description
End Function

' Пустая ячейка отключает строку. CVErr и Null остаются активными ошибками,
' если у строки заполнен ID; формула с результатом "" считается пустой.
Private Function TorsionValueBlank(ByVal value As Variant) As Boolean
    If IsError(value) Or IsNull(value) Then Exit Function
    TorsionValueBlank = (Len(Trim$(CStr(value))) = 0)
End Function
