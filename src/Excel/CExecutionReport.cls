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
' схемы. Это не журнал итераций Newton/Secant и не часть расчетного ядра.
'
' Отчет накапливается в памяти и записывается одним действием в txt-файл рядом
' с книгой. Такой подход не замедляет горячие расчетные циклы и не размазывает
' работу с файловой системой по solver-ам.

Private Const REPORT_FILE_NAME As String = "RC_Section_NDM_execution_report.txt" ' Фиксированное имя последнего отчета рядом с книгой.

Private mEnabled As Boolean ' True, если пользователь включил запись отчета в Config.
Private mFilePath As String ' Полный путь к txt-файлу, который будет перезаписан в конце запуска.
Private mStartedAt As Date ' Время начала расчетного сценария.
Private mStartTimer As Double ' Значение Timer для вывода относительного времени шагов.
Private mLines As Collection ' Накопленные строки отчета в порядке выполнения.

' Готовит объект к работе по настройке General.ExecutionReportEnabled.
' Если флаг выключен, остальные методы остаются безопасными пустыми операциями.
Public Sub Initialize(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
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

Private Sub AddRawLine(ByVal text As String)
    If mLines Is Nothing Then Set mLines = New Collection
    mLines.Add text
End Sub

Private Function JoinLines() As String
    Dim parts() As String
    ReDim parts(0 To mLines.Count - 1)

    Dim i As Long
    For i = 1 To mLines.Count
        parts(i - 1) = CStr(mLines.Item(i))
    Next i
    JoinLines = Join(parts, vbCrLf)
End Function

Private Function TimePrefix() As String
    TimePrefix = "[" & Format$(ElapsedSeconds(), "0.000") & " с]"
End Function

Private Function ElapsedSeconds() As Double
    ElapsedSeconds = Timer - mStartTimer
    If ElapsedSeconds < 0# Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

Private Function WorkbookFolder(ByVal workbook As Object) As String
    WorkbookFolder = Trim$(CStr(workbook.Path))
    If Len(WorkbookFolder) = 0 Then WorkbookFolder = CurDir$
End Function

Private Function SafeWorkbookName(ByVal workbook As Object) As String
    On Error GoTo Failed
    SafeWorkbookName = CStr(workbook.Name)
    Exit Function
Failed:
    SafeWorkbookName = "(имя книги недоступно)"
End Function
