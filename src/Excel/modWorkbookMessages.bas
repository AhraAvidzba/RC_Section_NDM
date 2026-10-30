Attribute VB_Name = "modWorkbookMessages"
Option Explicit

' ==========================================================================
' Unicode-сообщения пользовательских сценариев Excel
' ==========================================================================
' Модуль дополняет готовые пояснения адресами ввода и показывает их без ANSI.
' Это сохраняет греческие обозначения, математические знаки и адреса файлов
' независимо от системной кодовой страницы. Причины, статусы и коды результата
' определяют расчётные классы; оформление сообщений на их выбор не влияет.
' Все сообщения кнопок Excel и AutoCAD-export используют этот общий вывод.

Public Const CONTOUR_MESH_GENERATION_NOTICE As String = "БЕТОННАЯ СЕТКА АВТОМАТИЧЕСКИ СГЕНЕРИРОВАНА ПО КОНТУРАМ AUTOCAD." ' Только сообщение формирования сечения по контурам AutoCAD.

Private mPendingDrawingBooks As Collection ' Готовые изменения диаграмм до завершения показа; расчётные данные здесь не кэшируются.

#If VBA7 Then
    Private Declare PtrSafe Function PaintPendingWindow Lib "user32" Alias "RedrawWindow" ( _
        ByVal window As LongPtr, ByVal rectangle As LongPtr, ByVal region As LongPtr, ByVal flags As Long) As Long
    Private Declare PtrSafe Function MessageBoxUnicode Lib "user32" Alias "MessageBoxW" ( _
        ByVal ownerWindow As LongPtr, ByVal textPointer As LongPtr, _
        ByVal titlePointer As LongPtr, ByVal flags As Long) As Long
#Else
    Private Declare Function PaintPendingWindow Lib "user32" Alias "RedrawWindow" ( _
        ByVal window As Long, ByVal rectangle As Long, ByVal region As Long, ByVal flags As Long) As Long
    Private Declare Function MessageBoxUnicode Lib "user32" Alias "MessageBoxW" ( _
        ByVal ownerWindow As Long, ByVal textPointer As Long, _
        ByVal titlePointer As Long, ByVal flags As Long) As Long
#End If

' Показывает модальное сообщение, принадлежащее текущему Excel, и возвращает
' выбранную кнопку. StrPtr передает исходные UTF-16 строки непосредственно
' в Unicode API; текст, математические обозначения и флаги не подменяются.
' При ошибке самого окна возвращается явная техническая ошибка, а не успех.
Public Function ShowWorkbookMessage(ByVal message As String, _
        Optional ByVal style As VbMsgBoxStyle = vbInformation, _
        Optional ByVal title As String = "RC Section NDM") As VbMsgBoxResult
    Dim response As Long, displayText As String
    Dim drawingWarning As String: drawingWarning = TryCompleteWorkbookDrawing(ThisWorkbook)
    If Len(drawingWarning) > 0 Then
        message = WorkbookMessageParagraph(message, drawingWarning)
        If style = vbInformation Then style = vbExclamation
    End If
    displayText = WorkbookDialogText(message)
#If VBA7 Then
    response = MessageBoxUnicode(CLngPtr(Application.hWnd), StrPtr(displayText), StrPtr(title), CLng(style))
#Else
    response = MessageBoxUnicode(Application.hWnd, StrPtr(displayText), StrPtr(title), CLng(style))
#End If
    If response = 0 Then Err.Raise vbObjectError + 4250, "ShowWorkbookMessage", _
        "Не удалось показать сообщение Excel. Код системной ошибки: " & CStr(Err.LastDllError) & "."
    ShowWorkbookMessage = response
End Function

' Разделяет готовый текст окна на смысловые блоки, сохраняя причины, числа,
' идентификаторы и фактические адреса. Пояснения ячеек Results здесь не меняются.
' Повторное оформление не добавляет пустые строки; точки в числах и путях не разбиваются.
Public Function WorkbookDialogText(ByVal message As String) As String
    Dim text As String, caption As Variant, action As Variant, lines As Variant, i As Long
    text = message
    For Each caption In Array("Прочность:", "Устойчивость:", "Трещинообразование:", _
            "Ширина трещин:", "Продольные трещины:", "НДС:", "Предупреждение:")
        text = ReplaceWorkbookMessageFragment(text, "; " & CStr(caption), vbCrLf & vbCrLf & CStr(caption))
        text = ReplaceWorkbookMessageFragment(text, " " & CStr(caption), vbCrLf & vbCrLf & CStr(caption))
        text = ReplaceWorkbookMessageFragment(text, vbCrLf & CStr(caption), vbCrLf & vbCrLf & CStr(caption))
        text = ReplaceWorkbookMessageFragment(text, CStr(caption) & " ", CStr(caption) & vbCrLf)
    Next caption
    text = ReplaceWorkbookMessageFragment(text, " Настройка ", vbCrLf & vbCrLf & "Настройка ")
    text = ReplaceWorkbookMessageFragment(text, " Исходная строка сочетания:", vbCrLf & vbCrLf & "Исходная строка сочетания:")
    text = ReplaceWorkbookMessageFragment(text, "Исходная строка сочетания: ", "Исходная строка сочетания:" & vbCrLf)
    text = ReplaceWorkbookMessageFragment(text, " (статус ", vbCrLf & "(статус ")
    text = Replace$(Replace$(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(CStr(lines(i)), Len("Настройка ")) = "Настройка " Then lines(i) = ErrorReasonParagraphs(CStr(lines(i)))
    Next i
    text = Join(lines, vbCrLf)
    text = ReplaceWorkbookMessageFragment(text, " Введённое значение:", vbCrLf & "Введённое значение:")
    For Each action In Array("Введите ", "Выберите ", "Исправьте ", "Восстановите ", _
            "Уменьшите ", "Увеличьте ", "Измените ", "Проверьте ", "Укажите ", "Задайте ", _
            "Удалите ", "Заполните ", "Повторите ", "Сохраните ", "Закройте ", "Откройте ", "Обновите ", _
            "Добавьте ", "Установите ", "Используйте ")
        text = ReplaceWorkbookMessageFragment(text, ". " & CStr(action), "." & vbCrLf & vbCrLf & CStr(action))
    Next action
    ' Геометрические предупреждения хранятся компактно; только окно разделяет счётчик, причину и техническую ссылку.
    For Each caption In Array("В каждой совпавшей позиции ", "Их положения по нормали ", _
            "В расчёте и подписях учтён ", "Площадь каждого из них ")
        text = ReplaceWorkbookMessageFragment(text, ". " & CStr(caption), "." & vbCrLf & vbCrLf & CStr(caption))
    Next caption
    text = ReplaceWorkbookMessageFragment(text, ". Подробности и адреса полей приведены в ", _
        "." & vbCrLf & vbCrLf & "Подробности и адреса полей:" & vbCrLf)
    text = ReplaceWorkbookMessageFragment(text, vbCrLf & "По настройке Mesh.MinElementArea ", _
        vbCrLf & vbCrLf & "По настройке Mesh.MinElementArea ")
    WorkbookDialogText = TrimWorkbookMessage(text)
End Function

' Разделяет только пояснения программы: фразы внутри кавычек остаются точными
' пользовательскими значениями, именами слоёв или цитатами, а не командами.
Private Function ReplaceWorkbookMessageFragment(ByVal text As String, ByVal fragment As String, ByVal replacement As String) As String
    If InStr(1, text, fragment, vbBinaryCompare) = 0 Then ReplaceWorkbookMessageFragment = text: Exit Function
    Dim position As Long, start As Long, quoted As Boolean, brackets As Long, character As String, result As String
    position = 1: start = 1
    Do While position <= Len(text)
        character = Mid$(text, position, 1)
        If character = Chr$(34) Then quoted = Not quoted
        If Not quoted Then
            If character = "«" Then brackets = brackets + 1
            If character = "»" And brackets > 0 Then brackets = brackets - 1
        End If
        If Not quoted And brackets = 0 And Mid$(text, position, Len(fragment)) = fragment Then
            result = result & Mid$(text, start, position - start) & replacement
            position = position + Len(fragment): start = position
        Else
            position = position + 1
        End If
    Loop
    ReplaceWorkbookMessageFragment = result & Mid$(text, start)
End Function

' Добавляет непустой смысловой блок через одну пустую строку; пустой отчёт
' или предупреждение не оставляют лишнего разделителя в итоговом окне.
Public Function WorkbookMessageParagraph(ByVal message As String, ByVal paragraph As String) As String
    message = TrimWorkbookMessage(message): paragraph = TrimWorkbookMessage(paragraph)
    WorkbookMessageParagraph = message
    If Len(paragraph) = 0 Then Exit Function
    If Len(message) = 0 Then
        WorkbookMessageParagraph = paragraph
    Else
        WorkbookMessageParagraph = message & vbCrLf & vbCrLf & paragraph
    End If
End Function

' Вставляет короткий источник сразу после итогового абзаца, сохраняя счётчики, причины и пути отдельными блоками.
Public Function WorkbookMessageAfterOutcome(ByVal message As String, ByVal sourceText As String) As String
    Dim position As Long: position = InStr(1, message, vbCrLf & vbCrLf, vbBinaryCompare)
    If position = 0 Then
        WorkbookMessageAfterOutcome = WorkbookMessageParagraph(message, sourceText)
    Else
        WorkbookMessageAfterOutcome = Left$(message, position - 1) & vbCrLf & vbCrLf & sourceText & Mid$(message, position)
    End If
End Function

' Называет только источник текущего действия. Расчёт читает геометрию; экспорт полного снимка — геометрию и НДС.
Public Function SnapshotActionDataOrigin(ByVal workbook As Object, ByVal action As String, Optional ByVal geometryOnly As Boolean = False) As String
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim origin As String
    If store.StorageMode(workbook) = "File" Then
        origin = "из файла «" & store.SelectedFileName(workbook) & "»"
    Else
        origin = "с листа Results текущей книги"
    End If
    Select Case action
        Case "Calculation": SnapshotActionDataOrigin = "Геометрия для расчёта загружена " & origin & "."
        Case "Export"
            If geometryOnly Then SnapshotActionDataOrigin = "Геометрия для экспорта загружена " & origin & "." Else _
                SnapshotActionDataOrigin = "Геометрия и НДС для экспорта загружены " & origin & "."
        Case "Update": SnapshotActionDataOrigin = "Данные для схемы и таблиц загружены " & origin & "."
        Case Else: Err.Raise 5, "SnapshotActionDataOrigin", "Неизвестное действие для сообщения об источнике данных."
    End Select
End Function

' Собирает итог один раз по готовым результатам: отдельно считает ошибки ввода,
' внутренние ошибки и численные отказы; берёт собственную причину каждой категории.
' Чтение результатов не запускает solver и не меняет статусы или таблицы.
Public Function CalculationBatchMessage(ByVal batch As CBatchSectionCalculator, _
        Optional ByRef hasInputErrors As Boolean = False, Optional ByRef requiresAttention As Boolean = False) As String
    Dim index As Long, count As Long, firstError As String, meta As CResultMeta
    Dim internalCount As Long, numericalCount As Long, firstInternal As String, firstNumerical As String
    For index = 1 To batch.Count
        Set meta = batch.ResultAt(index).InputErrorMeta
        If Not meta Is Nothing Then
            count = count + 1
            If Len(firstError) = 0 Then firstError = RequiredAttentionMessage(batch, index, meta)
        End If
        Set meta = batch.ResultAt(index).InternalErrorMeta
        If Not meta Is Nothing Then
            internalCount = internalCount + 1
            If Len(firstInternal) = 0 Then firstInternal = RequiredAttentionMessage(batch, index, meta)
        End If
        Set meta = batch.ResultAt(index).NumericalFailureMeta
        If Not meta Is Nothing Then
            numericalCount = numericalCount + 1
            If Len(firstNumerical) = 0 Then firstNumerical = RequiredAttentionMessage(batch, index, meta)
        End If
    Next index
    hasInputErrors = (count > 0)
    requiresAttention = hasInputErrors Or internalCount > 0 Or numericalCount > 0
    Dim text As String
    text = CalculationCompletionMessage(batch.Count, count, firstError, vbNullString)
    If internalCount > 0 Then
        text = WorkbookMessageParagraph(text, "Внутренних ошибок: " & CStr(internalCount) & ".")
        text = WorkbookMessageParagraph(text, "Первая внутренняя ошибка:" & vbCrLf & firstInternal)
    End If
    If numericalCount > 0 Then
        text = WorkbookMessageParagraph(text, "Сочетаний с неподтверждённым численным результатом: " & CStr(numericalCount) & ".")
        text = WorkbookMessageParagraph(text, "Первая численная неудача:" & vbCrLf & firstNumerical)
    End If
    CalculationBatchMessage = WorkbookDialogText(text)
End Function

' Сохраняет машинную причину и адрес даже при пустом пользовательском пояснении.
' Этот запасной текст относится только к уже классифицированному обязательному уведомлению.
Private Function RequiredAttentionMessage(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal meta As CResultMeta) As String
    RequiredAttentionMessage = CalculationResultMessage(batch, index, meta)
    If Len(RequiredAttentionMessage) > 0 Then Exit Function
    RequiredAttentionMessage = "Сочетание " & batch.CombinationID(index) & ": описание причины не задано. Код причины: " & ResultCodeToText(meta.ResultCode) & "."
    Dim reference As String
    reference = batch.SettingReferences(meta.SourceKeys)
    If Len(reference) > 0 Then RequiredAttentionMessage = RequiredAttentionMessage & " " & reference
    reference = batch.SourceReference(index)
    If Len(reference) > 0 Then RequiredAttentionMessage = RequiredAttentionMessage & " Исходная строка сочетания: " & reference & "."
End Function

' Строит сводку запуска из готовых счётчиков и собственной причины первой ошибки.
' Не определяет расчётный статус и не заменяет исходную диагностику кратким кодом.
Public Function CalculationCompletionMessage(ByVal combinationCount As Long, ByVal inputErrorCount As Long, _
        ByVal firstInputError As String, ByVal governingCombinationID As String) As String
    Dim text As String, summary As String, detail As String
    summary = "Обработано сочетаний: " & CStr(combinationCount) & "."
    If inputErrorCount > 0 Then
        text = "Расчет завершен с ошибками ввода."
        summary = summary & vbCrLf & "Ошибок ввода: " & CStr(inputErrorCount) & "."
        text = WorkbookMessageParagraph(text, summary)
        text = WorkbookMessageParagraph(text, "Результаты и пояснения всех проверок — на листе Results.")
        detail = firstInputError
        If Left$(detail, Len("CombinationID=")) = "CombinationID=" Then detail = "Сочетание " & Mid$(detail, Len("CombinationID=") + 1)
        If Len(Trim$(detail)) > 0 Then text = WorkbookMessageParagraph(text, "Показана только первая ошибка ввода:" & vbCrLf & WorkbookDialogText(detail))
    Else
        text = WorkbookMessageParagraph("Расчет завершен.", summary)
        text = WorkbookMessageParagraph(text, "Результаты и пояснения всех проверок — на листе Results.")
        If Len(governingCombinationID) > 0 Then text = WorkbookMessageParagraph(text, "Определяющее сочетание: " & governingCombinationID)
    End If
    CalculationCompletionMessage = WorkbookDialogText(text)
End Function

' Нормализует переводы строк и убирает только внешние разделители и лишние
' пустые строки. Содержимое строк, пути, единицы и точные числа сохраняются.
Private Function TrimWorkbookMessage(ByVal text As String) As String
    text = Trim$(Replace$(Replace$(text, vbCrLf, vbLf), vbCr, vbLf))
    Do While Len(text) > 0
        If Left$(text, 1) <> vbLf Then Exit Do
        text = Mid$(text, 2)
    Loop
    Do While Len(text) > 0
        If Right$(text, 1) <> vbLf Then Exit Do
        text = Left$(text, Len(text) - 1)
    Loop
    Do While InStr(1, text, vbLf & vbLf & vbLf, vbBinaryCompare) > 0
        text = Replace$(text, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    TrimWorkbookMessage = Replace$(Trim$(text), vbLf, vbCrLf)
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
    WorkbookErrorMessage = WorkbookDialogText(WorkbookErrorMessage)
End Function


' Дополняет причину отказа формирования понятными действиями. Причина, слои
' и альтернативный способ задания сечения отделены пустыми строками.
Public Function SectionFormationFailureText(ByVal reason As String, ByVal source As String, _
        ByVal settings As CSystemSettingsReader) As String
    SectionFormationFailureText = SectionRecipeErrorText(reason)
    If StrComp(source, "AutoCAD", vbTextCompare) = 0 Then
        Dim solutionTitle As String
        solutionTitle = "Возможное решение:"
        If InStr(1, SectionFormationFailureText, solutionTitle, vbTextCompare) > 0 Then solutionTitle = "Проверьте также:"
        If Not settings Is Nothing Then
            SectionFormationFailureText = SectionFormationFailureText & vbCrLf & vbCrLf & _
                "Слои, указанные на Config:" & vbCrLf & _
                "Бетон: «" & settings.GetRawString("AutoCAD.Common.ConcreteLayer", "не задан") & "»." & vbCrLf & _
                "Арматура: «" & settings.GetRawString("AutoCAD.Common.RebarLayer", "не задан") & "»." & vbCrLf & _
                "Наружный контур: «" & settings.GetRawString("AutoCAD.Common.SectionContourLayer", "не задан") & "»."
        End If
        SectionFormationFailureText = SectionFormationFailureText & vbCrLf & vbCrLf & _
            solutionTitle & vbCrLf & _
            "Проверьте активный чертёж и наличие подходящей геометрии на этих слоях." & vbCrLf & vbCrLf & _
            "Другой вариант — сформировать сечение автоматически:" & vbCrLf & _
            "1. На Config выберите Geometry.Source = Generated и нужный Geometry.Type." & vbCrLf & _
            "2. На листе «Сечения» задайте размеры и армирование выбранного типа." & vbCrLf & _
            "3. Нажмите «Сформировать сечение»."
    ElseIf InStr(1, SectionFormationFailureText, "Возможное решение:", vbTextCompare) = 0 Then
        SectionFormationFailureText = SectionFormationFailureText & vbCrLf & vbCrLf & _
            "Возможное решение:" & vbCrLf & _
            "Проверьте поля, указанные в сообщении." & vbCrLf & _
            "Размеры и армирование задаются на листе «Сечения»." & vbCrLf & _
            "Источник, тип сечения и сетка задаются на Config." & vbCrLf & _
            "После исправления нажмите «Сформировать сечение»."
    End If
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


' Запоминает только необходимость закончить показ уже изменённой диаграммы, без повторного построения схемы.
Public Sub MarkWorkbookDrawingPending(ByVal workbook As Object)
    If mPendingDrawingBooks Is Nothing Then Set mPendingDrawingBooks = New Collection
    Dim pending As Object
    For Each pending In mPendingDrawingBooks
        If pending Is workbook Then Exit Sub
    Next pending
    mPendingDrawingBooks.Add workbook
End Sub

' После восстановления ScreenUpdating завершает показ готовой схемы и WM_PAINT своей книги.
' DoEvents не вызывается: обработчик не передаёт управление повторным нажатиям кнопок и очереди OnTime.
Public Sub CompleteWorkbookDrawing(ByVal workbook As Object)
    If workbook Is Nothing Then Exit Sub
    If Not workbook.Application.ScreenUpdating Then Exit Sub
    Dim refresh As Boolean, pending As Object, i As Long
    If Not mPendingDrawingBooks Is Nothing Then
        For i = mPendingDrawingBooks.Count To 1 Step -1
            Set pending = mPendingDrawingBooks.Item(i)
            If pending Is workbook Then mPendingDrawingBooks.Remove i: refresh = True
        Next i
        If mPendingDrawingBooks.Count = 0 Then Set mPendingDrawingBooks = Nothing
    End If
    Dim window As Object, succeeded As Long, flags As Long: flags = &H180&
    If refresh Then flags = flags Or &H1& ' Перерисовывается готовое окно; Chart.Refresh повторно не перестраивает диаграмму.
    For Each window In workbook.Windows
        If window.hWnd <> 0 Then ' Нулевой HWND обозначал бы рабочий стол Windows; его эта операция не затрагивает.
#If VBA7 Then
            succeeded = PaintPendingWindow(CLngPtr(window.hWnd), 0, 0, flags)
#Else
            succeeded = PaintPendingWindow(window.hWnd, 0, 0, flags)
#End If
            If succeeded = 0 Then Err.Raise vbObjectError + 4251, "CompleteWorkbookDrawing", "Excel не смог завершить обновление окна книги. Код Windows: " & CStr(Err.LastDllError)
        End If
    Next window
End Sub

' Сохранённые данные и НДС не превращаются в неудачный расчёт из-за отдельной ошибки экранного показа.
Public Function TryCompleteWorkbookDrawing(ByVal workbook As Object) As String
    On Error GoTo Failed
    CompleteWorkbookDrawing workbook
    Exit Function
Failed:
    TryCompleteWorkbookDrawing = "Предупреждение:" & vbCrLf & "Excel не завершил отрисовку схемы или окна книги. Данные сохранены." & vbCrLf & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет завершение показа конкретной книги без требования, чтобы чужие API-действия отсутствовали.
Public Function TestWorkbookDrawingPending(ByVal workbook As Object) As Boolean
    If mPendingDrawingBooks Is Nothing Then Exit Function
    Dim pending As Object
    For Each pending In mPendingDrawingBooks
        If pending Is workbook Then TestWorkbookDrawingPending = True: Exit Function
    Next pending
End Function
