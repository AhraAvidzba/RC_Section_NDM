VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExecutionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Человекочитаемый отчет выполнения расчета
' ==========================================================================
' Класс собирает короткий пошаговый журнал пользовательского запуска: чтение
' настроек, построение модели, обработку сочетаний, запись Results и обновление
' схемы. Также форматирует готовую диагностику решателей для пользователя,
' не изменяя исходные журналы, результаты и инженерные статусы.
'
' Отчет накапливается в памяти и записывается одним действием в txt-файл рядом
' с книгой. Такой подход не замедляет горячие расчетные циклы и не размазывает
' работу с файловой системой по решателям.

Private Const REPORT_FILE_NAME As String = "RC_Section_NDM_execution_report.txt" ' Фиксированное имя последнего отчета рядом с книгой.

Private mEnabled As Boolean ' True, если пользователь включил запись отчета в Config.
Private mFilePath As String ' Полный путь к txt-файлу, который будет перезаписан в конце запуска.
Private mStartedAt As Date ' Время начала расчетного сценария.
Private mStartTimer As Double ' Значение Timer для вывода относительного времени шагов.
Private mUnits As CUnitSystem ' Выходные единицы и знаки физических величин отчета.
Private mLines As Collection ' Накопленные строки отчета в порядке выполнения.

' Готовит объект к работе по настройке General.ExecutionReportEnabled.
' Если флаг выключен, остальные методы остаются безопасными пустыми операциями.
' Потерянная или неверная активная настройка возвращает адресную ошибку ввода,
' а не молча выбирает отсутствие отчета. Повторная инициализация очищает буфер.
Public Sub Initialize(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    mEnabled = False
    mFilePath = vbNullString
    Set mLines = New Collection
    If workbook Is Nothing Then Exit Sub
    If settings Is Nothing Then Exit Sub

    mEnabled = settings.GetRequiredBoolean("General.ExecutionReportEnabled")
    If Not mEnabled Then Exit Sub

    Set mLines = New Collection
    mStartedAt = Now
    mStartTimer = Timer
    mFilePath = WorkbookFolder(workbook) & "\" & REPORT_FILE_NAME

    AddRawLine "ОТЧЕТ ВЫПОЛНЕНИЯ RC SECTION NDM"
    AddRawLine String$(72, "=")
    AddRawLine "Книга: " & SafeWorkbookName(workbook)
    AddRawLine "Папка книги: " & WorkbookFolder(workbook)
    AddRawLine "Начало: " & Format$(mStartedAt, "yyyy-mm-dd hh:nn:ss")
    AddRawLine "Отметки [с] и длительность журнала отсчитываются от указанного начала отчета; общий таймер макроса включает более раннюю подготовку."
    AddRawLine String$(72, "-")
End Sub

' Добавляет визуальный раздел, чтобы txt-файл читался как небольшой протокол,
' а не как сплошной поток строк.
Public Sub AddSection(ByVal title As String)
    If Not mEnabled Then Exit Sub
    AddRawLine vbNullString
    AddRawLine String$(72, "-")
    AddRawLine Trim$(title)
    AddRawLine String$(72, "-")
End Sub

' Записывает обычный шаг выполнения с отметкой времени от начала запуска.
Public Sub AddStep(ByVal text As String)
    If Not mEnabled Then Exit Sub
    AddRawLine TimePrefix() & " " & Trim$(text)
End Sub

' Записывает пару "параметр - значение". Метод удобен для краткой фиксации
' исходных режимов, числа элементов и итоговых статусов.
Public Sub AddValue(ByVal label As String, ByVal valueText As String)
    If Not mEnabled Then Exit Sub
    AddRawLine TimePrefix() & " " & Trim$(label) & ": " & Trim$(valueText)
End Sub

' Записывает заголовок и готовый многострочный журнал итераций в общий отчёт запуска.
' Подходит для решения НДС и одномерного поиска несущей способности.
Public Sub AddBlock(ByVal title As String, ByVal body As String)
    If Not mEnabled Then Exit Sub
    body = Trim$(body)
    If Len(body) = 0 Then Exit Sub

    AddRawLine TimePrefix() & " " & Trim$(title)
    AddRawLine body
End Sub

' Фиксирует ошибку перед сохранением отчета. Ошибка не подавляется: вызывающий
' код показывает пользователю обычное сообщение Excel.
Public Sub AddError(ByVal context As String, ByVal errorText As String)
    If Not mEnabled Then Exit Sub
    AddSection "Ошибка"
    AddStep Trim$(context) & ": " & Trim$(errorText)
End Sub

' Записывает накопленный отчет в txt-файл. Файл последнего запуска намеренно
' перезаписывается, чтобы в папке книги не копились устаревшие протоколы.
Public Sub Save(Optional ByVal finalStatus As String = vbNullString)
    If Not mEnabled Then Exit Sub

    AddRawLine vbNullString
    AddRawLine String$(72, "-")
    AddRawLine "Итог"
    AddRawLine String$(72, "-")
    If Len(Trim$(finalStatus)) > 0 Then AddRawLine Trim$(finalStatus)
    AddRawLine "Окончание: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AddRawLine "Длительность ведения журнала: " & Format$(ElapsedSeconds(), "0.000") & " с"

    Dim fso As Object
    Dim stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(mFilePath, True, True)
    stream.Write JoinLines()
    stream.Close
End Sub

' Выводит отдельный итог T/NMT из готовых таблиц снимка в их сохраненных выходных единицах.
' НДС и равновесие заново не решает; пустые результаты не превращает в нули.
Public Sub ReportTorsionResults(ByVal snapshot As CCalculationSnapshot)
    If Not mEnabled Then Exit Sub
    If Not snapshot.HasTable(TORSION_POST_DATA_KEY) Then Exit Sub
    AddSection "Кручение и проверочная сумма напряжений NM+T"
    AddStep "T (кручение) рассматривается как момент вокруг продольной оси сечения. При отсутствии поперечной силы перенос точки в плоскости сечения не меняет T. " & _
        "Сумма NM+T проверяется сложением напряжений; совместное равновесие NMT не рассчитывается."
    Dim data As Variant: data = snapshot.Table(TORSION_POST_DATA_KEY)
    Dim row As Long, count As Long, momentUnit As String
    momentUnit = "(единица снимка)": If Not mUnits Is Nothing Then momentUnit = mUnits.OutputMomentUnit
    For row = 1 To UBound(data, 1)
        If Len(CStr(data(row, tscCombinationID))) > 0 Then
            count = count + 1
            AddValue "Сочетание ID", CStr(data(row, tscCombinationID))
            AddValue "T вокруг продольной оси сечения", SavedNumber(data(row, tscTorque)) & " " & momentUnit
            AddValue "Общий статус проверки T/NMT", CStr(data(row, tscOverallStatus))
            AddValue "Продольная арматура от T: запас / статус", SavedNumber(data(row, tscLongitudinalReserve)) & " / " & CStr(data(row, tscLongitudinalStatus))
            AddValue "Хомут: запас / статус", SavedNumber(data(row, tscHoopReserve)) & " / " & CStr(data(row, tscHoopStatus))
            AddValue "Бетон: запас / статус", SavedNumber(data(row, tscConcreteReserve)) & " / " & CStr(data(row, tscConcreteStatus))
            AddValue "Сумма напряжений NM+T: запас / статус", CombinedReserveText(data, row) & " / " & CStr(data(row, tscCombinedStatus))
            If Len(CStr(data(row, tscResultComment))) > 0 Then AddValue "Пояснение T/NMT", CStr(data(row, tscResultComment))
        End If
    Next row
    If count = 0 Then AddStep "Активных сочетаний с T нет; кручение не рассчитывалось."
End Sub

' Форматирует сохраненную выходную величину без повторного преобразования единиц.
Private Function SavedNumber(ByVal value As Variant) As String
    If IsEmpty(value) Then SavedNumber = "не определялось": Exit Function
    If IsNumeric(value) Then SavedNumber = FormatNumberInvariant(CDbl(value)) Else SavedNumber = CStr(value)
End Function

' Отличает неограниченный запас при сжатии от недоступной суммы напряжений.
Private Function CombinedReserveText(ByRef data As Variant, ByVal row As Long) As String
    CombinedReserveText = SavedNumber(data(row, tscCombinedReserve))
    If Not IsEmpty(data(row, tscCombinedReserve)) Or IsEmpty(data(row, tscCombinedStress)) Then Exit Function
    If CDbl(data(row, tscCombinedStress)) <= 0# And CStr(data(row, tscCombinedStatus)) = "OK" Then _
        CombinedReserveText = "не ограничен по растяжению"
End Function

' Сообщает вызывающему коду, нужно ли дописывать путь к отчету в итоговое
' пользовательское сообщение.
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

' Полный путь к последнему отчету.
Public Property Get FilePath() As String
    FilePath = mFilePath
End Property

' Добавляет строку без форматирования во внутренний буфер отчета.
Private Sub AddRawLine(ByVal text As String)
    If mLines Is Nothing Then Set mLines = New Collection
    mLines.Add text
End Sub

' Собирает коллекцию строк в один текст с Windows-переносами строк.
Private Function JoinLines() As String
    Dim parts() As String
    ReDim parts(0 To mLines.Count - 1)

    Dim i As Long
    For i = 1 To mLines.Count
        parts(i - 1) = CStr(mLines.Item(i))
    Next i
    JoinLines = Join(parts, vbCrLf)
End Function

' Возвращает префикс вида [0.123 с] для строк пошагового отчета.
Private Function TimePrefix() As String
    TimePrefix = "[" & Format$(ElapsedSeconds(), "0.000") & " с]"
End Function

' Считает прошедшее время с учетом перехода Timer через полночь.
Private Function ElapsedSeconds() As Double
    ElapsedSeconds = Timer - mStartTimer
    If ElapsedSeconds < 0# Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

' Возвращает папку книги или текущую папку, если книга еще не сохранена.
Private Function WorkbookFolder(ByVal workbook As Object) As String
    WorkbookFolder = Trim$(CStr(workbook.Path))
    If Len(WorkbookFolder) = 0 Then WorkbookFolder = CurDir$
End Function

' Безопасно читает имя книги для отчета выполнения.
Private Function SafeWorkbookName(ByVal workbook As Object) As String
    On Error GoTo Failed
    SafeWorkbookName = CStr(workbook.Name)
    Exit Function
Failed:
    SafeWorkbookName = "(имя книги недоступно)"
End Function

' Подключает централизованный адаптер единиц для форматирования отчета.
' Это не меняет численные результаты и не включает запись отчета.
Public Property Set UnitSystem(ByVal units As CUnitSystem)
    Set mUnits = units
End Property

' Форматирует число независимо от регионального десятичного разделителя.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
    If Right$(FormatNumberInvariant, 1) = "." Then FormatNumberInvariant = Left$(FormatNumberInvariant, Len(FormatNumberInvariant) - 1)
End Function

' Добавляет понятное представление диагностики в отчёт. Исходный подробный журнал
' сохраняется в результате; сокращение повторов меняет только текст вывода.
Public Sub AddDiagnosticBlock(ByVal title As String, ByVal body As String)
    If Not mEnabled Then Exit Sub
    If Len(Trim$(body)) = 0 Then Exit Sub
    AddStep "Диагностика ниже дана во внутренних единицах и знаках: N, Nint, RN — Н; Mx, My, Mxint, Myint, RMx, RMy — Н*мм. " & _
        "N > 0 означает растяжение; Mx > 0 растягивает сторону +Y, My > 0 — сторону +X. " & _
        "Моменты и их невязки относятся к началу O(0; 0) внутренней расчетной системы сечения. " & _
        "Координаты — мм; epsilon0 относится к O, кривизны — 1/мм. Номер пробы или оценки плоскости не является lambda."
    AddBlock title, HumanReadableReportBlock(title, body)
End Sub

' Подписывает точку в той же расчетной системе, в которой хранятся координаты элементов сечения.
Public Function FormatPoint(ByVal x As Double, ByVal y As Double) As String
    FormatPoint = "X=" & FormatLength(x) & "; Y=" & FormatLength(y) & _
        " (внутренняя расчетная система координат элементов сечения)"
End Function

' Записывает усилия с явной точкой отсчета моментов; N при переносе точки не меняется.
Public Sub ReportLoads(ByVal title As String, ByVal n As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal pointName As String, ByVal x As Double, ByVal y As Double)
    If Not mEnabled Then Exit Sub
    AddValue title & " — моменты относительно " & pointName, _
        "N=" & FormatForce(n) & "; Mx=" & FormatMx(mx) & "; My=" & FormatMy(my)
    AddValue "Точка отсчета: " & pointName, FormatPoint(x, y)
End Sub

' Форматирует длину для отчёта в той же выходной единице, что таблица Results.
' Во внутренних расчётных данных длина остаётся в мм.
Public Function FormatLength(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatLength = FormatNumberInvariant(value) & " мм"
    Else
        FormatLength = FormatNumberInvariant(mUnits.InternalLengthToOutput(value)) & " " & mUnits.OutputLengthUnit
    End If
End Function

' Форматирует площадь в пользовательской выходной единице площади.
Public Function FormatArea(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatArea = FormatNumberInvariant(value) & " мм2"
    Else
        FormatArea = FormatNumberInvariant(mUnits.InternalAreaToOutput(value)) & " " & mUnits.OutputAreaUnit
    End If
End Function

' Форматирует обычную знаковую продольную силу N с учетом Sign.N.User.
Public Function FormatForce(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatForce = FormatNumberInvariant(value) & " Н"
    Else
        FormatForce = FormatNumberInvariant(mUnits.InternalForceToOutput(value)) & " " & mUnits.OutputForceUnit
    End If
End Function

' Форматирует положительную величину сжимающей силы, например Ncr или Nult.
' Внутренний знак восстанавливается как сжатие, чтобы пользователь видел
' тот же знак, что в таблице устойчивости.
Public Function FormatCompressionForce(ByVal compressionMagnitude As Double) As String
    If mUnits Is Nothing Then
        FormatCompressionForce = FormatNumberInvariant(-Abs(compressionMagnitude)) & " Н"
    Else
        FormatCompressionForce = FormatNumberInvariant(mUnits.InternalForceToOutput(-Abs(compressionMagnitude))) & _
            " " & mUnits.OutputForceUnit
    End If
End Function

' Форматирует уже подписанную сжимающую силу, где положительное значение
' означает сжатие в локальной нормативной формуле устойчивости.
Public Function FormatSignedCompressionForce(ByVal signedCompression As Double) As String
    If mUnits Is Nothing Then
        FormatSignedCompressionForce = FormatNumberInvariant(-signedCompression) & " Н"
    Else
        FormatSignedCompressionForce = FormatNumberInvariant(mUnits.InternalForceToOutput(-signedCompression)) & _
            " " & mUnits.OutputForceUnit
    End If
End Function

' Форматирует Mx с пользовательским знаком Sign.Mx.User.
Public Function FormatMx(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatMx = FormatNumberInvariant(value) & " Н*мм"
    Else
        FormatMx = FormatNumberInvariant(mUnits.InternalMomentMxToOutput(value)) & " " & mUnits.OutputMomentUnit
    End If
End Function

' Форматирует My с пользовательским знаком Sign.My.User.
Public Function FormatMy(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatMy = FormatNumberInvariant(value) & " Н*мм"
    Else
        FormatMy = FormatNumberInvariant(mUnits.InternalMomentMyToOutput(value)) & " " & mUnits.OutputMomentUnit
    End If
End Function

' Форматирует модуль момента без изменения знака по глобальным осям.
' Используется для M1/M2, Mxy,crc и других ненаправленных величин отчёта.
Public Function FormatMomentMagnitude(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatMomentMagnitude = FormatNumberInvariant(value) & " Н*мм"
    Else
        FormatMomentMagnitude = FormatNumberInvariant(mUnits.InternalMomentMagnitudeToOutput(value)) & _
            " " & mUnits.OutputMomentUnit
    End If
End Function

' Форматирует напряжение/модуль в пользовательской выходной единице напряжений.
Public Function FormatStress(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatStress = FormatNumberInvariant(value) & " МПа"
    Else
        FormatStress = FormatNumberInvariant(mUnits.InternalStressToOutput(value)) & " " & mUnits.OutputStressUnit
    End If
End Function

' Форматирует жесткость D устойчивости. Размерность задается как сила*длина2
' и использует ту же пару output-единиц силы и длины, что Results.
Public Function FormatStiffness(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatStiffness = FormatNumberInvariant(value) & " Н*мм2"
    Else
        FormatStiffness = FormatNumberInvariant(mUnits.InternalForceLengthSquaredToOutput(value)) & _
            " " & mUnits.OutputForceLengthSquaredUnit
    End If
End Function

' Единая короткая русская запись булевых значений для отчета.
Public Function YesNoRu(ByVal value As Boolean) As String
    If value Then
        YesNoRu = "да"
    Else
        YesNoRu = "нет"
    End If
End Function

' Русское название нормативного блока устойчивости для txt-отчета.
Public Function StabilityCodeDisplayName(ByVal codeText As String) As String
    Select Case UCase$(Trim$(codeText))
        Case "SP63"
            StabilityCodeDisplayName = "СП 63"
        Case "SP35"
            StabilityCodeDisplayName = "СП 35"
        Case Else
            StabilityCodeDisplayName = codeText
    End Select
End Function

' Русское описание внутренней ветви устойчивости. Расчетные идентификаторы
' остаются в коде, но отчет должен читаться без знания этих имен.
Public Function StabilityBranchDisplayName(ByVal branchText As String) As String
    Select Case LCase$(Trim$(branchText))
        Case "nocompression"
            StabilityBranchDisplayName = "устойчивость не требуется: нет сжатия"
        Case "sp63"
            StabilityBranchDisplayName = "СП 63: расчет eta и Ncr"
        Case "sp35-eta"
            StabilityBranchDisplayName = "СП 35: ветвь |ec| > r через eta и Ncr"
        Case "sp35-table"
            StabilityBranchDisplayName = "СП 35: табличная ветвь |ec| <= r через Nult,stab"
        Case "sp35-mixed"
            StabilityBranchDisplayName = "СП 35: плоскости попали в разные ветви"
        Case Else
            StabilityBranchDisplayName = branchText
    End Select
End Function

' Делает PASS/FAIL в отчете понятным русским текстом.
Public Function StabilityPassDisplayName(ByVal passed As Boolean) As String
    If passed Then
        StabilityPassDisplayName = "проверка пройдена"
    Else
        StabilityPassDisplayName = "проверка не пройдена"
    End If
End Function

' Преобразует служебную диагностику решателей в понятный текст отчёта.
' Подробные исходные DiagnosticLog сохраняются без изменений;
' повторяющиеся строки вроде initialGuess поясняются и сокращаются только при выводе.
Private Function HumanReadableReportBlock(ByVal title As String, ByVal body As String) As String
    Dim text As String
    text = body
    If Len(text) = 0 Then Exit Function

    If InStr(1, title, "Диагностика расчета трещин", vbTextCompare) > 0 Or _
            InStr(1, title, "Диагностика образования трещины", vbTextCompare) > 0 Then
        text = CompactCrackDiagnosticText(text)
    End If
    If InStr(1, title, "Сценарий прямого StateSolution", vbTextCompare) > 0 Or _
            InStr(1, title, "Итерации прямого CSectionSolver", vbTextCompare) > 0 Or _
            InStr(1, title, "Итерации внутреннего CSectionSolver", vbTextCompare) > 0 Then
        text = CompactStateSolutionDiagnosticText(text)
    End If
    If InStr(1, title, "Диагностика поиска несущей способности", vbTextCompare) > 0 Then
        text = CompactCapacityDiagnosticText(text)
    End If

    HumanReadableReportBlock = text
End Function

' Оставляет события расчёта трещин и итерации lambda_crc в исходном порядке,
' а одинаковые сообщения о начальных приближениях заменяет краткой сводкой.
Private Function CompactCrackDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim parts() As String, partCount As Long
    ReDim parts(0 To UBound(lines) + 2)
    Dim warmStartCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then ' Пустые строки внутри diagnostic-блока не несут смысла.
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        Else
            AppendReportPart parts, partCount, lineText
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportPart parts, partCount, "Для повторных НДС использовалась стартовая плоскость из предыдущего близкого решения; повторов: " & CStr(warmStartCount) & "."

    CompactCrackDiagnosticText = TranslateDiagnosticLine(JoinReportParts(parts, partCount))
End Function

' Сокращает повторяющиеся сообщения CStateSolutionRunner о начальных приближениях.
' Полная техническая диагностика остаётся в исходном результате.
Private Function CompactStateSolutionDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim parts() As String, partCount As Long
    ReDim parts(0 To UBound(lines) + 2)
    Dim warmStartCount As Long
    Dim activeGuessCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then ' Пропускаем.
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        ElseIf lineText = "initialGuess=directStateActiveSet" Then
            activeGuessCount = activeGuessCount + 1
        Else
            AppendReportPart parts, partCount, lineText
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportPart parts, partCount, "Стартовая плоскость: использовано предыдущее близкое решение; повторов: " & CStr(warmStartCount) & "."
    If activeGuessCount > 0 Then _
        AppendReportPart parts, partCount, "Стартовая плоскость: подобрана по активной зоне сечения; повторов: " & CStr(activeGuessCount) & "."

    CompactStateSolutionDiagnosticText = TranslateDiagnosticLine(JoinReportParts(parts, partCount))
End Function

' Готовит журнал поиска несущей способности для отчёта. Сохраняет
' каждую проверенную нагрузку lambda и сведения о повторных попытках НДС.
Private Function CompactCapacityDiagnosticText(ByVal body As String) As String
    If Len(body) = 0 Then Exit Function
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim parts() As String, partCount As Long
    ReDim parts(0 To UBound(lines))
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then ' Пропускаем.
        Else
            AppendReportPart parts, partCount, lineText
        End If
    Next i

    CompactCapacityDiagnosticText = TranslateDiagnosticLine(JoinReportParts(parts, partCount))
End Function

' Переводит маркеры всего блока, не изменяя исходную диагностику решателей.
Private Function TranslateDiagnosticLine(ByVal lineText As String) As String
    Dim translated As String
    translated = lineText
    ReplaceDiagnosticMarker translated, "ultimateTrustRegion flat trial rejected; original recovery remains available", _
        "Резервная проба на плато отклонена; прежний подбор шага сохранен."
    ReplaceDiagnosticMarker translated, "ultimateTrustRegion attempt=", "Резервный шаг на активном пределе: проба="
    ReplaceDiagnosticMarker translated, "; rho=", "; соответствие прогноза="
    ReplaceWholeDiagnosticLine translated, "ultimatePath pureAxial redirected=LoadMultiplier", "Силовая траектория lambda*N перенаправлена на LoadMultiplier, потому что UltimateStrain для нее вырождается."
    ReplaceWholeDiagnosticLine translated, "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier.", "Силовая траектория lambda*N перенаправлена на LoadMultiplier, потому что UltimateStrain для нее вырождается."
    ReplaceDiagnosticMarker translated, "ultimatePath iter=", "оценка предельной плоскости по общей lambda-траектории №="
    ReplaceDiagnosticMarker translated, "ultimate iter=", "оценка предельной плоскости по моментной траектории №="
    ReplaceDiagnosticMarker translated, "lambda_crc=", "Найден lambda_crc="
    ReplaceDiagnosticMarker translated, "lambda_crc_iter=", "итерация прямого поиска lambda_crc="
    ReplaceDiagnosticMarker translated, "lambda_crc_path_iter=", "lambda_crc проверяемой плоскости="
    ReplaceDiagnosticMarker translated, "lambda_crc_load_probe=", "точка LoadMultiplier lambda_crc="
    ReplaceDiagnosticMarker translated, "iterations=", "итераций="
    ReplaceDiagnosticMarker translated, "method=", "метод="
    ReplaceDiagnosticMarker translated, "formation=Ncrc", "Образование трещины найдено по центральному растяжению Ncrc"
    ReplaceDiagnosticMarker translated, "search=", "Поиск: "
    ReplaceDiagnosticMarker translated, "searchMethod=", "метод поиска="
    ReplaceDiagnosticMarker translated, "secantGuard=insideBracket", "защита Secant=точка возвращена внутрь интервала поиска"
    ReplaceDiagnosticMarker translated, "rawLambda=", "исходная lambda="
    ReplaceDiagnosticMarker translated, "guardedLambda=", "защищенная lambda="
    ReplaceDiagnosticMarker translated, "bracketLow=", "нижняя граница="
    ReplaceDiagnosticMarker translated, "bracketHigh=", "верхняя граница="
    ReplaceDiagnosticMarker translated, "probe=", "проба №="
    ReplaceDiagnosticMarker translated, "initialGuess=warmStartLoadState", "начальное приближение=сохраненное НДС близкой нагрузки"
    ReplaceDiagnosticMarker translated, "ultimateNewton step=", "шаг Ньютона предельной плоскости №="
    ReplaceDiagnosticMarker translated, "step=", "шаг="
    ReplaceDiagnosticMarker translated, "iter=", "итерация="
    ReplaceDiagnosticMarker translated, "directStateRetry=", "повтор прямого НДС="
    ReplaceDiagnosticMarker translated, "extensionRetry=", "повтор после выхода за диаграмму="
    ReplaceDiagnosticMarker translated, "directStateActiveSet", "активная зона прямого НДС"
    ReplaceDiagnosticMarker translated, "warmStartSolver", "предыдущее близкое решение"
    ReplaceDiagnosticMarker translated, "direct target", "целевая нагрузка"
    ReplaceDiagnosticMarker translated, "stepped", "ступенчатая нагрузка"
    ReplaceDiagnosticMarker translated, "attempt=", "повтор НДС="
    ReplaceDiagnosticMarker translated, "converged=True", "сошлось=да"
    ReplaceDiagnosticMarker translated, "converged=False", "сошлось=нет"
    ReplaceDiagnosticMarker translated, "state=", "состояние="
    ReplaceDiagnosticMarker translated, "stop=", "причина="
    ReplaceDiagnosticMarker translated, "lineSearchUsed=", "подбор длины шага использован="
    ReplaceDiagnosticMarker translated, "lineSearchIterations=", "итераций подбора длины шага="
    ReplaceDiagnosticMarker translated, "internalForceEvaluationCount=", "оценок внутренних усилий="
    ReplaceDiagnosticMarker translated, "internalNewtonCallCount=", "вызовов Newton="
    ReplaceDiagnosticMarker translated, "result=converged", "результат=сошлось"
    ReplaceDiagnosticMarker translated, "result=failed", "результат=не сошлось"
    ReplaceDiagnosticMarker translated, "rPath1=", "невязка пути 1="
    ReplaceDiagnosticMarker translated, "rPath2=", "невязка пути 2="
    ReplaceDiagnosticMarker translated, "rLimit=", "невязка предельной деформации="
    ReplaceDiagnosticMarker translated, "rDir=", "невязка направления="
    ReplaceDiagnosticMarker translated, "residual=", "невязка образования трещины="
    ReplaceDiagnosticMarker translated, "rN=", "невязка N="
    ReplaceDiagnosticMarker translated, "rMomentDir=", "невязка направления момента="
    ReplaceDiagnosticMarker translated, "eps_bt_ult=", "eps_bt,ult="
    ReplaceDiagnosticMarker translated, "residualN=", "невязка N="
    ReplaceDiagnosticMarker translated, "residualMx=", "невязка Mx="
    ReplaceDiagnosticMarker translated, "residualMy=", "невязка My="
    ReplaceDiagnosticMarker translated, "critical=", "критический элемент="
    ReplaceDiagnosticMarker translated, "minConcreteStrain=", "мин. деформация бетона="
    ReplaceDiagnosticMarker translated, "maxConcreteStrain=", "макс. деформация бетона="
    ReplaceDiagnosticMarker translated, "maxAbsSteelStrain=", "макс. модуль деформации арматуры="
    TranslateDiagnosticLine = translated
End Function

' Не копирует большой блок при отсутствии маркера; порядок замен сохранен.
Private Sub ReplaceDiagnosticMarker(ByRef text As String, ByVal marker As String, ByVal replacement As String)
    If InStr(1, text, marker, vbBinaryCompare) > 0 Then text = Replace$(text, marker, replacement, 1, -1, vbBinaryCompare)
End Sub

' Особое правило действует только на целую строку, включая первую и последнюю.
Private Sub ReplaceWholeDiagnosticLine(ByRef text As String, ByVal marker As String, ByVal replacement As String)
    Dim position As Long, nextPosition As Long, searchFrom As Long, startsLine As Boolean, endsLine As Boolean
    searchFrom = 1
    position = InStr(searchFrom, text, marker, vbBinaryCompare)
    Do While position > 0
        startsLine = (position = 1)
        If position > 2 Then startsLine = (Mid$(text, position - 2, 2) = vbCrLf)
        nextPosition = position + Len(marker)
        endsLine = (nextPosition > Len(text))
        If Not endsLine Then endsLine = (Mid$(text, nextPosition, 2) = vbCrLf)
        If startsLine And endsLine Then
            text = Left$(text, position - 1) & replacement & Mid$(text, nextPosition)
            searchFrom = position + Len(replacement)
        Else
            searchFrom = nextPosition
        End If
        If searchFrom > Len(text) Then Exit Do
        position = InStr(searchFrom, text, marker, vbBinaryCompare)
    Loop
End Sub

' Записывает подробности устойчивости по двум главным плоскостям,
' чтобы было видно, какой способ расчёта СП применён к сочетанию.
' Основная таблица результатов остаётся на Results.
Public Sub ReportStabilityDetails(ByVal stability As CStabilityResult)
    If Not mEnabled Then Exit Sub
    If stability Is Nothing Then Exit Sub

    If Len(stability.Meta.ResultComment) > 0 Then _
        AddValue "Пояснение результата устойчивости", stability.Meta.ResultComment
    If Not stability.PlaneApplicable1 And Not stability.PlaneApplicable2 Then
        AddStep "Проверка главных плоскостей не выполнена или не требуется; см. пояснение результата. Ее числовые параметры не выводятся."
        Exit Sub
    End If
    AddValue "Центр приведенного сечения для устойчивости", FormatPoint(stability.ReferenceX, stability.ReferenceY)
    AddStep "M1 и M2 относятся к внутренним главным центральным осям сечения; единицы моментов — выходные."
    AddValue "Продольная сила для проверки", FormatForce(stability.DesignN)
    AddValue "Моменты после учета eta относительно центра приведенного сечения, главные оси 1/2", _
        "M1=" & FormatMomentMagnitude(stability.DesignMoment1) & "; " & _
        "M2=" & FormatMomentMagnitude(stability.DesignMoment2)
    AddValue "Длительная часть продольной силы", FormatSignedCompressionForce(stability.SustainedN)
    AddValue "Длительные моменты относительно того же центра, главные оси 1/2", _
        "M1l=" & FormatMomentMagnitude(stability.SustainedMoment1) & "; " & _
        "M2l=" & FormatMomentMagnitude(stability.SustainedMoment2)

    AddValue "Плоскость 1: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch1) & "; " & StabilityPassDisplayName(stability.PlanePassed1)
    AddValue "Плоскость 1: eta и запас", _
        FormatNumberInvariant(stability.Eta1) & "; " & FormatNumberInvariant(stability.Reserve1)
    If LCase$(stability.PlaneBranch1) = "sp35-table" Then
        AddValue "Плоскость 1: Nult,stab", FormatCompressionForce(stability.Nultimate1)
        AddStep "В табличной ветви плоскости 1 Ncr и жесткость D не рассчитываются."
    ElseIf stability.PlaneApplicable1 Then
        AddValue "Плоскость 1: Ncr", FormatCompressionForce(stability.Ncr1)
    End If
    AddValue "Плоскость 1: h, i и гибкость l0/i", _
        "h=" & FormatLength(stability.Depth1) & "; " & _
        "i=" & FormatLength(stability.Radius1) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness1)
    If UCase$(stability.Code) = "SP35" Then AddValue "Плоскость 1: ядровое расстояние r", FormatLength(stability.CoreDistance1)
    AddValue "Плоскость 1: e, ea", _
        "e=" & FormatLength(stability.Eccentricity1) & "; " & _
        "ea=" & FormatLength(stability.AccidentalEcc1)
    If stability.PlaneApplicable1 And LCase$(stability.PlaneBranch1) <> "sp35-table" Then _
        AddValue "Плоскость 1: phi_l и D", "phi_l=" & FormatNumberInvariant(stability.PhiL1) & "; D=" & FormatStiffness(stability.StiffnessD1)

    AddValue "Плоскость 2: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch2) & "; " & StabilityPassDisplayName(stability.PlanePassed2)
    AddValue "Плоскость 2: eta и запас", _
        FormatNumberInvariant(stability.Eta2) & "; " & FormatNumberInvariant(stability.Reserve2)
    If LCase$(stability.PlaneBranch2) = "sp35-table" Then
        AddValue "Плоскость 2: Nult,stab", FormatCompressionForce(stability.Nultimate2)
        AddStep "В табличной ветви плоскости 2 Ncr и жесткость D не рассчитываются."
    ElseIf stability.PlaneApplicable2 Then
        AddValue "Плоскость 2: Ncr", FormatCompressionForce(stability.Ncr2)
    End If
    AddValue "Плоскость 2: h, i и гибкость l0/i", _
        "h=" & FormatLength(stability.Depth2) & "; " & _
        "i=" & FormatLength(stability.Radius2) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness2)
    If UCase$(stability.Code) = "SP35" Then AddValue "Плоскость 2: ядровое расстояние r", FormatLength(stability.CoreDistance2)
    AddValue "Плоскость 2: e, ea", _
        "e=" & FormatLength(stability.Eccentricity2) & "; " & _
        "ea=" & FormatLength(stability.AccidentalEcc2)
    If stability.PlaneApplicable2 And LCase$(stability.PlaneBranch2) <> "sp35-table" Then _
        AddValue "Плоскость 2: phi_l и D", "phi_l=" & FormatNumberInvariant(stability.PhiL2) & "; D=" & FormatStiffness(stability.StiffnessD2)
End Sub

' Поясняет в txt-отчете стартовую логику поиска образования нормальной трещины.
' Низкоуровневый DiagnosticLog ниже сохраняет точные попытки, а эта строка
' нужна пользователю как быстрый ответ, что именно программа собирается делать.
Public Sub ReportCrackFormationSearchStart(ByVal formationPath As String)
    If Not mEnabled Then Exit Sub
    If StrComp(formationPath, "Auto", vbTextCompare) = 0 Then
        AddStep "Трещинообразование: режим Auto. При одной составляющей нагрузки используется путь по этой составляющей: " & _
            "N, Mx, My или совместный Mxy. При N + M сначала масштабируется момент, затем при неудаче поиска " & _
            "пробуются пути по N и NMxy. Найденный предел при множителе меньше 1 не запускает другой путь; " & _
            "ошибки входа и настройки не маскируются повторным поиском."
    Else
        AddStep "Трещинообразование: используется фиксированный путь " & formationPath & _
            "; Auto-перебор других путей не выполняется."
    End If
End Sub

' Объясняет готовый итог образования трещины, различая отсутствие трещины
' и неудачу поиска. Применимость ширины, фактический psi_s
' и предупреждения берутся из отдельного результата расчёта ширины.
Public Sub ReportCrackFormationChoice(ByVal crack As CCrackFormationResult)
    If Not mEnabled Then Exit Sub
    If crack Is Nothing Then Exit Sub

    If crack.HasLimitPoint Then
        AddStep "Трещинообразование: принята точка по пути " & _
            crack.FormationMethod & ". По этой точке записаны Ncrc/Mxy,crc и состояния PreCrackState/PostCrackState, если они применимы."
    ElseIf crack.ConfirmedNotCracked Then
        AddStep "Трещинообразование: подтверждено отсутствие нормальной трещины при заданной нагрузке; " & _
            "проверка ширины не выполняется. " & crack.ResultMeta.ResultComment
    ElseIf crack.ResultMeta.InternalStatus = rsNotApplicable Then
        AddStep "Трещинообразование: поиск пороговой точки не требуется. " & crack.ResultMeta.ResultComment
    Else
        AddStep "Трещинообразование: пригодная точка образования не получена; причина сохранена в результате поиска и диагностике ниже. " & _
            "Применимость дальнейших проверок, использованный коэффициент и предупреждения приведены в отдельных результатах трещин."
    End If
End Sub

' Записывает готовый результат продольных трещин, полученный
' CLongitudinalCrackCalculator. Равновесие и напряжения заново не рассчитываются.
Public Sub ReportLongitudinalCrackResult(ByVal longitudinalResult As CLongitudinalCrackResult)
    If Not mEnabled Then Exit Sub
    If longitudinalResult Is Nothing Then
        AddStep "Проверка продольных трещин не сформировала результат."
        Exit Sub
    End If

    Dim meta As CResultMeta
    Set meta = longitudinalResult.ResultMeta
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Dim statusText As String
    statusText = policy.ExternalStatus(meta)

    AddValue "Статус продольных трещин", statusText
    If statusText <> policy.NotApplicable Then
        AddValue "sigma_c,max", FormatStress(longitudinalResult.MaxCompressionStress)
        AddValue "Rb,mc2", FormatStress(longitudinalResult.RbMc2)
        AddValue "Использование по продольным трещинам", _
            FormatNumberInvariant(longitudinalResult.Utilization)
    ElseIf Len(Trim$(meta.ResultComment)) > 0 Then
        AddStep meta.ResultComment
    Else
        AddStep "Проверка продольных трещин неприменима: в прямом НДС нет сжатого бетона."
    End If
End Sub

' Сохраняет перевод каждой строки без повторного копирования всего блока.
Private Sub AppendReportPart(ByRef parts() As String, ByRef count As Long, ByVal lineText As String)
    If Len(lineText) = 0 Then Exit Sub
    parts(count) = lineText
    count = count + 1
End Sub

' Собирает заполненную часть массива строк в исходном порядке.
' Пустые части отбрасываются раньше, при добавлении через AppendReportPart.
Private Function JoinReportParts(ByRef parts() As String, ByVal count As Long) As String
    If count = 0 Then Exit Function
    ReDim Preserve parts(0 To count - 1)
    JoinReportParts = Join(parts, vbCrLf)
End Function
