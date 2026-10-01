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
' работу с файловой системой по solver-ам.

Private Const REPORT_FILE_NAME As String = "RC_Section_NDM_execution_report.txt" ' Фиксированное имя последнего отчета рядом с книгой.

Private mEnabled As Boolean ' True, если пользователь включил запись отчета в Config.
Private mFilePath As String ' Полный путь к txt-файлу, который будет перезаписан в конце запуска.
Private mStartedAt As Date ' Время начала расчетного сценария.
Private mStartTimer As Double ' Значение Timer для вывода относительного времени шагов.
Private mUnits As CUnitSystem ' Выходные единицы и знаки физических величин отчета.
Private mLines As Collection ' Накопленные строки отчета в порядке выполнения.

' Готовит объект к работе по настройке General.ExecutionReportEnabled.
' Если флаг выключен, остальные методы остаются безопасными пустыми операциями.
Public Sub Initialize(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    mEnabled = False
    mFilePath = vbNullString
    Set mLines = New Collection
    If workbook Is Nothing Then Exit Sub
    If settings Is Nothing Then Exit Sub

    mEnabled = settings.GetBoolean("General.ExecutionReportEnabled", False)
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
    AddRawLine String$(72, "-")
    If Len(settings.MigrationWarning) > 0 Then AddStep settings.MigrationWarning
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

' Добавляет многострочный диагностический блок, например журнал итераций
' CSectionSolver или одномерного поиска несущей способности. Метод не
' анализирует текст журнала: он только аккуратно вкладывает готовую трассировку
' в общий отчет запуска.
Public Sub AddBlock(ByVal title As String, ByVal body As String)
    If Not mEnabled Then Exit Sub
    If Len(Trim$(body)) = 0 Then Exit Sub

    AddRawLine TimePrefix() & " " & Trim$(title)
    AddRawLine Trim$(body)
End Sub

' Фиксирует ошибку перед сохранением отчета. Ошибка не подавляется: вызывающий
' код по-прежнему показывает пользователю обычное сообщение Excel.
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
    AddRawLine "Длительность по таймеру: " & Format$(ElapsedSeconds(), "0.000") & " с"

    Dim fso As Object
    Dim stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(mFilePath, True, True)
    stream.Write JoinLines()
    stream.Close
End Sub

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
End Function

' Добавляет человекочитаемую диагностику, сохраняя сырой журнал в result.
' Только presentation-слой сжимает повторяющиеся строки и переводит подписи.
Public Sub AddDiagnosticBlock(ByVal title As String, ByVal body As String)
    If Not mEnabled Then Exit Sub
    AddBlock title, HumanReadableReportBlock(title, body)
End Sub

' Форматирует длину для execution_report в той же выходной единице, что и
' табличный Results. Внутри batch все еще хранит миллиметры.
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

' Форматирует обычную знаковую продольную силу N с учетом Output.SignConvention.N.
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

' Форматирует Mx с пользовательским знаком Output.SignConvention.Mx.
Public Function FormatMx(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatMx = FormatNumberInvariant(value) & " Н*мм"
    Else
        FormatMx = FormatNumberInvariant(mUnits.InternalMomentMxToOutput(value)) & " " & mUnits.OutputMomentUnit
    End If
End Function

' Форматирует My с пользовательским знаком Output.SignConvention.My.
Public Function FormatMy(ByVal value As Double) As String
    If mUnits Is Nothing Then
        FormatMy = FormatNumberInvariant(value) & " Н*мм"
    Else
        FormatMy = FormatNumberInvariant(mUnits.InternalMomentMyToOutput(value)) & " " & mUnits.OutputMomentUnit
    End If
End Function

' Форматирует скалярный момент без смены знака: используется для M1/M2,
' Mxy,crc и других величин, где в отчете нужен модуль, а не глобальный Mx/My.
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
            StabilityBranchDisplayName = "СП 35: ветвь ec > r через eta и Ncr"
        Case "sp35-table"
            StabilityBranchDisplayName = "СП 35: табличная ветвь ec <= r через Nult,stab"
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

' Преобразует низкоуровневую диагностику solver-ов в форму, пригодную для
' пользовательского txt-отчета. Внутренние DiagnosticLog остаются подробными
' и машинно-проверяемыми, но отчет не должен забиваться сотнями строк
' initialGuess/lambda_crc_iter без пояснений.
Private Function HumanReadableReportBlock(ByVal title As String, ByVal body As String) As String
    Dim text As String
    text = body
    If Len(text) = 0 Then Exit Function

    If InStr(1, title, "Диагностика расчета трещин", vbTextCompare) > 0 Then
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

' Приводит подробный журнал трещин к читаемому виду: смысловые строки и
' итерации lambda_crc оставляет по порядку, а повторяющиеся warm-start
' сообщения заменяет короткой сводкой.
Private Function CompactCrackDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim warmStartCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пустые строки внутри diagnostic-блока не несут смысла.
        ElseIf Left$(lineText, Len("lambda_crc_iter=")) = "lambda_crc_iter=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf Left$(lineText, Len("lambda_crc_path_iter=")) = "lambda_crc_path_iter=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportLine result, "Для повторных НДС использовалась стартовая плоскость из предыдущего близкого решения; повторов: " & CStr(warmStartCount) & "."

    CompactCrackDiagnosticText = result
End Function

' Сжимает diagnostic CStateSolutionRunner. Особенно важно убрать десятки
' одинаковых warm-start строк: они полезны разработчику, но не пользователю.
Private Function CompactStateSolutionDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim warmStartCount As Long
    Dim activeGuessCount As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пропускаем.
        ElseIf Left$(lineText, Len("step=")) = "step=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        ElseIf lineText = "initialGuess=warmStartSolver" Then
            warmStartCount = warmStartCount + 1
        ElseIf lineText = "initialGuess=directStateActiveSet" Then
            activeGuessCount = activeGuessCount + 1
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    If warmStartCount > 0 Then _
        AppendReportLine result, "Стартовая плоскость: использовано предыдущее близкое решение; повторов: " & CStr(warmStartCount) & "."
    If activeGuessCount > 0 Then _
        AppendReportLine result, "Стартовая плоскость: подобрана по активной зоне сечения; повторов: " & CStr(activeGuessCount) & "."

    CompactStateSolutionDiagnosticText = result
End Function

' Сжимает диагностический блок поиска несущей способности. Полный список
' lambda-точек может быть большим, но для инженерной диагностики полезно
' видеть каждую проверенную точку и возможный retry НДС.
Private Function CompactCapacityDiagnosticText(ByVal body As String) As String
    Dim lines() As String
    lines = Split(body, vbCrLf)

    Dim result As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineText As String
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then
            ' Пропускаем.
        ElseIf Left$(lineText, Len("probe=")) = "probe=" Then
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        Else
            AppendReportLine result, TranslateDiagnosticLine(lineText)
        End If
    Next i

    CompactCapacityDiagnosticText = result
End Function

' Переводит короткие машинные diagnostic-строки в более понятные фразы.
Private Function TranslateDiagnosticLine(ByVal lineText As String) As String
    Dim translated As String
    translated = lineText
    If translated = "ultimatePath pureAxial redirected=LoadMultiplier" Or _
            translated = "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier." Then
        TranslateDiagnosticLine = "Силовая траектория lambda*N перенаправлена на LoadMultiplier, потому что UltimateStrain для нее вырождается."
        Exit Function
    End If
    translated = Replace$(translated, "ultimatePath iter=", "итерация UltimateStrain по общей lambda-траектории=")
    translated = Replace$(translated, "ultimate iter=", "итерация UltimateStrain по моментной траектории=")
    translated = Replace$(translated, "lambda_crc=", "Найден lambda_crc=")
    translated = Replace$(translated, "lambda_crc_iter=", "итерация прямого поиска lambda_crc=")
    translated = Replace$(translated, "lambda_crc_path_iter=", "итерация UltimateStrain по lambda-траектории lambda_crc=")
    translated = Replace$(translated, "lambda_crc_load_probe=", "точка LoadMultiplier lambda_crc=")
    translated = Replace$(translated, "iterations=", "итераций=")
    translated = Replace$(translated, "method=", "метод=")
    translated = Replace$(translated, "formation=Ncrc", "Образование трещины найдено по центральному растяжению Ncrc")
    translated = Replace$(translated, "search=", "Поиск: ")
    translated = Replace$(translated, "searchMethod=", "метод поиска=")
    translated = Replace$(translated, "secantGuard=insideBracket", "защита Secant=точка возвращена внутрь скобки")
    translated = Replace$(translated, "rawLambda=", "исходная lambda=")
    translated = Replace$(translated, "guardedLambda=", "защищенная lambda=")
    translated = Replace$(translated, "bracketLow=", "нижняя граница=")
    translated = Replace$(translated, "bracketHigh=", "верхняя граница=")
    translated = Replace$(translated, "probe=", "точка lambda=")
    translated = Replace$(translated, "step=", "шаг=")
    translated = Replace$(translated, "iter=", "итерация=")
    translated = Replace$(translated, "directStateRetry=", "повтор прямого НДС=")
    translated = Replace$(translated, "extensionRetry=", "повтор после выхода за диаграмму=")
    translated = Replace$(translated, "directStateActiveSet", "активная зона прямого НДС")
    translated = Replace$(translated, "warmStartSolver", "предыдущее близкое решение")
    translated = Replace$(translated, "direct target without line search", "целевая нагрузка без line search")
    translated = Replace$(translated, "direct target", "целевая нагрузка")
    translated = Replace$(translated, "stepped", "ступенчатая нагрузка")
    translated = Replace$(translated, "attempt=", "повтор НДС=")
    translated = Replace$(translated, "converged=True", "сошлось=да")
    translated = Replace$(translated, "converged=False", "сошлось=нет")
    translated = Replace$(translated, "state=", "состояние=")
    translated = Replace$(translated, "stop=", "причина=")
    translated = Replace$(translated, "lineSearchUsed=", "line search использован=")
    translated = Replace$(translated, "lineSearchIterations=", "итераций line search=")
    translated = Replace$(translated, "internalForceEvaluationCount=", "оценок внутренних усилий=")
    translated = Replace$(translated, "internalNewtonCallCount=", "вызовов Newton=")
    translated = Replace$(translated, "result=converged", "результат=сошлось")
    translated = Replace$(translated, "result=failed", "результат=не сошлось")
    translated = Replace$(translated, "rPath1=", "невязка пути 1=")
    translated = Replace$(translated, "rPath2=", "невязка пути 2=")
    translated = Replace$(translated, "rLimit=", "невязка предельной деформации=")
    translated = Replace$(translated, "rDir=", "невязка направления=")
    translated = Replace$(translated, "residual=", "невязка образования трещины=")
    translated = Replace$(translated, "rN=", "невязка N=")
    translated = Replace$(translated, "rMomentDir=", "невязка направления момента=")
    translated = Replace$(translated, "eps_bt_ult=", "eps_bt,ult=")
    translated = Replace$(translated, "residualN=", "невязка N=")
    translated = Replace$(translated, "residualMx=", "невязка Mx=")
    translated = Replace$(translated, "residualMy=", "невязка My=")
    translated = Replace$(translated, "critical=", "критический элемент=")
    translated = Replace$(translated, "minConcreteStrain=", "мин. деформация бетона=")
    translated = Replace$(translated, "maxConcreteStrain=", "макс. деформация бетона=")
    translated = Replace$(translated, "maxAbsSteelStrain=", "макс. модуль деформации арматуры=")
    TranslateDiagnosticLine = translated
End Function

' Пишет в txt-отчет подробные данные фильтра устойчивости по двум главным
' плоскостям. Таблица Results остается основным машинным выводом, а этот блок
' нужен для быстрой трассировки того, какая ветвь СП сработала в сценарии LC.
Public Sub ReportStabilityDetails(ByVal stability As CStabilityResult)
    If Not mEnabled Then Exit Sub
    If stability Is Nothing Then Exit Sub

    If Len(stability.Meta.ResultComment) > 0 Then _
        AddValue "Причина остановки устойчивости", stability.Meta.ResultComment
    AddValue "Продольная сила для проверки", FormatForce(stability.DesignN)
    AddValue "Моменты после учета eta по главным плоскостям 1/2", _
        "M1=" & FormatMomentMagnitude(stability.DesignMoment1) & "; " & _
        "M2=" & FormatMomentMagnitude(stability.DesignMoment2)
    AddValue "Длительная часть продольной силы", FormatSignedCompressionForce(stability.SustainedN)
    AddValue "Длительные моменты по главным плоскостям 1/2", _
        "M1l=" & FormatMomentMagnitude(stability.SustainedMoment1) & "; " & _
        "M2l=" & FormatMomentMagnitude(stability.SustainedMoment2)

    AddValue "Плоскость 1: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch1) & "; " & StabilityPassDisplayName(stability.PlanePassed1)
    AddValue "Плоскость 1: eta и запас", _
        FormatNumberInvariant(stability.Eta1) & "; " & FormatNumberInvariant(stability.Reserve1)
    AddValue "Плоскость 1: Ncr/Nult", _
        "Ncr=" & FormatCompressionForce(stability.Ncr1) & "; Nult=" & FormatCompressionForce(stability.Nultimate1)
    AddValue "Плоскость 1: h, i, r и гибкость l0/i", _
        "h=" & FormatLength(stability.Depth1) & "; " & _
        "i=" & FormatLength(stability.Radius1) & "; " & _
        "r=" & FormatLength(stability.CoreDistance1) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness1)
    AddValue "Плоскость 1: e, ea, phi_l и D", _
        "e=" & FormatLength(stability.Eccentricity1) & "; " & _
        "ea=" & FormatLength(stability.AccidentalEcc1) & "; " & _
        "phi_l=" & FormatNumberInvariant(stability.PhiL1) & "; " & _
        "D=" & FormatStiffness(stability.StiffnessD1)

    AddValue "Плоскость 2: ветвь и итог", _
        StabilityBranchDisplayName(stability.PlaneBranch2) & "; " & StabilityPassDisplayName(stability.PlanePassed2)
    AddValue "Плоскость 2: eta и запас", _
        FormatNumberInvariant(stability.Eta2) & "; " & FormatNumberInvariant(stability.Reserve2)
    AddValue "Плоскость 2: Ncr/Nult", _
        "Ncr=" & FormatCompressionForce(stability.Ncr2) & "; Nult=" & FormatCompressionForce(stability.Nultimate2)
    AddValue "Плоскость 2: h, i, r и гибкость l0/i", _
        "h=" & FormatLength(stability.Depth2) & "; " & _
        "i=" & FormatLength(stability.Radius2) & "; " & _
        "r=" & FormatLength(stability.CoreDistance2) & "; " & _
        "l0/i=" & FormatNumberInvariant(stability.Slenderness2)
    AddValue "Плоскость 2: e, ea, phi_l и D", _
        "e=" & FormatLength(stability.Eccentricity2) & "; " & _
        "ea=" & FormatLength(stability.AccidentalEcc2) & "; " & _
        "phi_l=" & FormatNumberInvariant(stability.PhiL2) & "; " & _
        "D=" & FormatStiffness(stability.StiffnessD2)
End Sub

' Поясняет в txt-отчете стартовую логику поиска образования нормальной трещины.
' Низкоуровневый DiagnosticLog ниже сохраняет точные попытки, а эта строка
' нужна пользователю как быстрый ответ, что именно программа собирается делать.
Public Sub ReportCrackFormationSearchStart(ByVal formationPath As String)
    If Not mEnabled Then Exit Sub
    If StrComp(formationPath, "Auto", vbTextCompare) = 0 Then
        AddStep "Трещинообразование: режим Auto. Программа пробует пути по порядку: " & _
            ChrW$(&H3BB) & "*Mxy, затем " & ChrW$(&H3BB) & "*N, затем " & _
            ChrW$(&H3BB) & "*NMxy. Если путь дает только резервный результат с " & _
            ChrW$(&H3C8) & "s = 1 без найденной точки образования трещины, программа переходит к следующему пути."
    Else
        AddStep "Трещинообразование: используется фиксированный путь " & formationPath & _
            "; Auto-перебор других путей не выполняется."
    End If
End Sub

' Поясняет итог выбора пути образования трещины для текущего LC. Это не заменяет
' DiagnosticLog solver-а, а переводит ключевой результат Auto/резервного
' нормальный инженерный язык прямо в отчете.
Public Sub ReportCrackFormationChoice(ByVal crack As CCrackFormationResult)
    If Not mEnabled Then Exit Sub
    If crack Is Nothing Then Exit Sub

    If Len(crack.FormationMethod) > 0 Then
        AddStep "Трещинообразование: принята точка по пути " & _
            crack.FormationMethod & ". По этой точке записаны Ncrc/Mxy,crc и состояния PreCrackState/PostCrackState, если они применимы."
    ElseIf crack.CrackFormed And crack.Converged Then
        AddStep "Трещинообразование: нормальная трещина считается образованной, но выбранный путь не дал отдельную расчетную точку образования трещины. Для этой проверки принят резервный вариант " & _
            ChrW$(&H3C8) & "s = 1; причина указана в CrackFormationStopReason и диагностике ниже."
    ElseIf crack.Converged Then
        AddStep "Трещинообразование: проверка завершена, нормальная трещина при текущем сочетании не образуется."
    Else
        AddStep "Трещинообразование: точку образования трещины найти не удалось; дальнейшая причина приведена в диагностике ниже."
    End If
End Sub

' Записывает в txt-отчет результат проверки продольных трещин, уже полученный
' отдельным CLongitudinalCrackCalculator. Здесь не запускается solver и не
' пересчитываются напряжения: отчет только поясняет готовый result-объект.
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

' Аккуратно добавляет строку в собираемый многострочный блок отчета.
Private Sub AppendReportLine(ByRef target As String, ByVal lineText As String)
    If Len(lineText) = 0 Then Exit Sub
    If Len(target) > 0 Then target = target & vbCrLf
    target = target & lineText
End Sub
