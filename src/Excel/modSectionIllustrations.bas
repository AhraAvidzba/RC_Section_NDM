Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая справочная картинка из именованных рисунков листа «Сечения»
' ==========================================================================
Private Const SOURCE_SHEET As String = "Сечения"
Private Const TARGET_SHEET As String = "Расчет"
Private Const SELECTOR_NAME As String = "rngSectionIllustrationSelector"
Private Const TARGET_PICTURE As String = "CurrentSection"

' Копирует выбранный рисунок, пропорционально вписывая его в текущую рамку.
' Источник заменяется вручную; список выбранной ячейки содержит простое перечисление имён.
' Расчётные таблицы и режим пересчёта не используются.
' Обновление экрана не переключается; исходный лист не активируется.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, selector As Range, original As Shape, frame As Shape, replacement As Shape
    Dim pictureName As String, fitRatio As Double, previousName As String
    Dim previousUpdating As Boolean, previousEvents As Boolean, previousSelection As Object, previousSheet As Object
    Dim number As Long, description As String, action As String, renamed As Boolean
    Dim attempt As Long, pasteNumber As Long, pasteDescription As String
    ' Сохраняем общие флаги Excel и место пользователя до перехода между листами.
    previousUpdating = Application.ScreenUpdating ' Флаг нужен только для обработки отложенной перерисовки в конце.
    previousEvents = Application.EnableEvents
    Set previousSheet = Application.ActiveSheet
    If TypeOf Application.Selection Is Range Then Set previousSelection = Application.Selection
    On Error GoTo Failed
    action = SELECTOR_NAME
    Set selector = SectionIllustrationSelector()
    action = selector.Worksheet.Name & "!" & selector.Address(False, False) & " (" & SELECTOR_NAME & ")"
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    pictureName = CStr(selector.Value2)
    action = SOURCE_SHEET & "!" & pictureName
    Set original = ThisWorkbook.Worksheets(SOURCE_SHEET).Shapes(pictureName)
    action = TARGET_SHEET & "!" & TARGET_PICTURE
    Set frame = sheet.Shapes(TARGET_PICTURE)
    If original.Type <> msoPicture Or frame.Type <> msoPicture Then _
        Err.Raise vbObjectError + 2801, , "Источник и CurrentSection должны быть рисунками."
    action = "Вписывание рисунка «" & pictureName & "»"
    ' Габариты CurrentSection постоянны; коэффициент выбирается по ограничивающей стороне.
    fitRatio = frame.Width / original.Width
    If original.Height * fitRatio > frame.Height Then fitRatio = frame.Height / original.Height
    Application.EnableEvents = False ' Собственные переходы и выделения не вызывают обработчики книги повторно.
    For attempt = 1 To 3
        original.Copy ' Берём именованный рисунок напрямую, не переключаясь на исходный лист.
        DoEvents ' Excel должен обработать отложенное заполнение буфера перед Paste.
        If Not Application.ActiveSheet Is sheet Then sheet.Activate ' При выборе на целевом листе он уже активен.
        On Error Resume Next
        Err.Clear
        sheet.Paste
        pasteNumber = Err.Number: pasteDescription = Err.Description
        On Error GoTo Failed
        If pasteNumber = 0 Then Exit For
        DoEvents ' При временной занятости буфера заново копируем исходный рисунок.
    Next attempt
    If pasteNumber <> 0 Then Err.Raise pasteNumber, , pasteDescription
    Set replacement = Application.Selection.ShapeRange.Item(1) ' Paste выделяет вставленный рисунок; индекс Shapes не равен порядку создания.
    With replacement.PictureFormat.Crop
        .ShapeWidth = frame.Width: .ShapeHeight = frame.Height
        .ShapeLeft = frame.Left: .ShapeTop = frame.Top
        .PictureWidth = original.Width * fitRatio: .PictureHeight = original.Height * fitRatio
        .PictureOffsetX = 0: .PictureOffsetY = 0 ' Центр картинки внутри постоянной рамки, без обрезки.
    End With
    replacement.Placement = frame.Placement
    replacement.LockAspectRatio = frame.LockAspectRatio
    replacement.AlternativeText = pictureName
    replacement.OnAction = vbNullString
    ' Пока новая картинка не готова, существующий CurrentSection сохраняется.
    previousName = replacement.Name
    frame.Name = previousName & "_previous": renamed = True
    replacement.Name = TARGET_PICTURE
    frame.Delete ' Удаляем исходный рисунок только после успешной подготовки замены.
    GoTo RestoreState
Failed:
    number = Err.Number: description = action & ": " & Err.Description
    On Error Resume Next
    If Not replacement Is Nothing Then replacement.Delete
    If renamed Then frame.Name = TARGET_PICTURE
RestoreState:
    On Error Resume Next
    ' Возвращаем выделение и состояние событий; режим обновления экрана остаётся у вызывающего кода.
    Application.CutCopyMode = False
    If Not Application.ActiveSheet Is previousSheet Then previousSheet.Activate
    If Not previousSelection Is Nothing Then previousSelection.Select
    Application.EnableEvents = previousEvents
    If previousUpdating Then DoEvents ' Даём Excel обработать отложенную перерисовку и стрелку списка.
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "SwitchSectionPicture", description
End Sub

' Возвращает единственную ячейку выбора в этой книге; её адрес может изменяться при вырезании и вставке.
Public Function SectionIllustrationSelector() As Range
    Dim cell As Range
    On Error GoTo InvalidName
    Set cell = ThisWorkbook.Names(SELECTOR_NAME).RefersToRange
    If cell.CountLarge <> 1 Then GoTo InvalidName
    If Not cell.Worksheet.Parent Is ThisWorkbook Then GoTo InvalidName
    Set SectionIllustrationSelector = cell
    Exit Function
InvalidName:
    Err.Raise vbObjectError + 2802, "SectionIllustrationSelector", _
        "Именованный диапазон «" & SELECTOR_NAME & "» должен ссылаться на одну ячейку этой книги." & vbCrLf & vbCrLf & _
        "Проверьте ссылку в диспетчере имён. Ячейку выбора можно перемещать вырезанием и вставкой."
End Function

' Отбирает только изменение текущей именованной ячейки; чужие листы и книги не затрагивает.
Public Function IsSectionIllustrationSelectorChange(ByVal changed As Range) As Boolean
    Dim selector As Range
    On Error GoTo Unrelated
    Set selector = SectionIllustrationSelector()
    If Not changed.Worksheet Is selector.Worksheet Then Exit Function
    IsSectionIllustrationSelectorChange = Not Application.Intersect(changed, selector) Is Nothing
Unrelated:
End Function
