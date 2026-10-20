Attribute VB_Name = "modWorkbookMessages"
Option Explicit

Public Const CONTOUR_MESH_GENERATION_NOTICE As String = "БЕТОННАЯ СЕТКА АВТОМАТИЧЕСКИ СГЕНЕРИРОВАНА ПО КОНТУРАМ AUTOCAD."

' ==========================================================================
' Unicode-сообщения пользовательских сценариев Excel
' ==========================================================================
' Модуль дополняет готовые пояснения адресами ввода и показывает их без ANSI.
' Это сохраняет греческие обозначения, математические знаки и адреса файлов
' независимо от системной кодовой страницы. Причины, статусы и коды результата
' определяют расчётные классы; оформление сообщений на их выбор не влияет.
' Все сообщения кнопок Excel и AutoCAD-export используют этот общий вывод.

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxUnicode Lib "user32" Alias "MessageBoxW" ( _
        ByVal ownerWindow As LongPtr, ByVal textPointer As LongPtr, _
        ByVal titlePointer As LongPtr, ByVal flags As Long) As Long
#Else
    Private Declare Function MessageBoxUnicode Lib "user32" Alias "MessageBoxW" ( _
        ByVal ownerWindow As Long, ByVal textPointer As Long, _
        ByVal titlePointer As Long, ByVal flags As Long) As Long
#End If

' Отдельная заметная строка сохраняет происхождение сетки при импорте
' и последующем экспорте, в том числе после открытия сохраненного Results.
Public Function WithConcreteMeshGenerationNotice(ByVal text As String, ByVal section As CSectionModel) As String
    WithConcreteMeshGenerationNotice = text
    If section.ConcreteMeshSource = "AutoCADContours" Then _
        WithConcreteMeshGenerationNotice = text & vbCrLf & vbCrLf & CONTOUR_MESH_GENERATION_NOTICE
End Function

' Показывает модальное сообщение, принадлежащее текущему Excel, и возвращает
' выбранную кнопку. StrPtr передает исходные UTF-16 строки непосредственно
' в Unicode API; текст, математические обозначения и флаги не подменяются.
' При ошибке самого окна возвращается явная техническая ошибка, а не успех.
Public Function ShowWorkbookMessage(ByVal message As String, _
        Optional ByVal style As VbMsgBoxStyle = vbInformation, _
        Optional ByVal title As String = "RC Section NDM") As VbMsgBoxResult
    Dim response As Long
#If VBA7 Then
    response = MessageBoxUnicode(CLngPtr(Application.hWnd), StrPtr(message), StrPtr(title), CLng(style))
#Else
    response = MessageBoxUnicode(Application.hWnd, StrPtr(message), StrPtr(title), CLng(style))
#End If
    If response = 0 Then Err.Raise vbObjectError + 4250, "ShowWorkbookMessage", _
        "Не удалось показать сообщение Excel. Код системной ошибки: " & CStr(Err.LastDllError) & "."
    ShowWorkbookMessage = response
End Function

' Дополняет готовое пояснение проверки номером сочетания и сохранённым адресом ввода.
' Статус и физическую причину определяют расчётные классы; здесь меняется только текст.
Public Function CalculationResultMessage(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal meta As CResultMeta) As String
    If meta Is Nothing Then Exit Function
    Dim reason As String, reference As String, parameterReferences As String
    reason = Trim$(meta.ResultComment)
    If Len(reason) = 0 Then Exit Function
    CalculationResultMessage = "Сочетание " & batch.CombinationID(index) & ": " & reason
    If meta.InternalStatus = rsNumericalFailure Then CalculationResultMessage = CalculationResultMessage & _
        " Эта проверка не получила подтверждённого численного результата; сам статус NumFail не доказывает физическое невыполнение. Проверьте исходные данные и диагностику расчёта."
    parameterReferences = batch.SettingReferences(meta.SourceKeys)
    If Len(parameterReferences) > 0 Then CalculationResultMessage = CalculationResultMessage & " " & parameterReferences
    reference = batch.SourceReference(index)
    If Len(reference) > 0 Then CalculationResultMessage = CalculationResultMessage & " Исходная строка сочетания: " & reference & "."
End Function

' Сохраняет исходную причину ошибки кнопки и технический источник для проверки разработчиком.
' При пустом системном описании пользователь всё равно получает код и место остановки.
Public Function WorkbookErrorMessage(ByVal operation As String, ByVal number As Long, _
        ByVal source As String, ByVal reason As String) As String
    WorkbookErrorMessage = operation & ":" & vbCrLf & vbCrLf
    If Len(Trim$(reason)) > 0 Then
        WorkbookErrorMessage = WorkbookErrorMessage & ErrorReasonParagraphs(reason)
    Else
        WorkbookErrorMessage = WorkbookErrorMessage & "Программа остановилась без описания причины."
    End If
    If Len(source) > 0 Then WorkbookErrorMessage = WorkbookErrorMessage & vbCrLf & vbCrLf & "Источник ошибки: " & source & "."
    WorkbookErrorMessage = WorkbookErrorMessage & vbCrLf & "Код ошибки: " & CStr(number) & "."
End Function

' Оформляет обычные сообщения reader-а так же, как геометрические ошибки.
' Адрес берётся из уже сохранённого места ввода; для утраченного ключа остаётся название раздела.
Private Function ErrorReasonParagraphs(ByVal reason As String) As String
    ErrorReasonParagraphs = reason
    If Left$(reason, Len("Настройка ")) <> "Настройка " Then Exit Function
    Dim first As Long, last As Long, location As String, details As String, position As Long
    first = InStr(1, reason, " (лист ", vbBinaryCompare)
    If first = 0 Then Exit Function
    last = InStr(first + 2, reason, ")", vbBinaryCompare)
    If last = 0 Then Exit Function
    location = Mid$(reason, first + 2, last - first - 2)
    position = InStr(1, location, ", ячейка ", vbBinaryCompare)
    If position > 0 Then location = Mid$(location, Len("лист ") + 1, position - Len("лист ") - 1) & "!" & Mid$(location, position + Len(", ячейка "))
    details = Trim$(Mid$(reason, last + 1))
    If Left$(details, 1) = ":" Then details = Trim$(Mid$(details, 2))
    ErrorReasonParagraphs = SectionRecipeErrorText(details, location, Left$(reason, first - 1))
End Function
