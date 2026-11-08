Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая справочная картинка из именованных рисунков листа «Сечения»
' ==========================================================================
Private Const SOURCE_SHEET As String = "Сечения"
Private Const TARGET_SHEET As String = "Расчет"
Private Const SELECTOR_NAME As String = "rngSectionIllustrationSelector"
Private Const TARGET_PICTURE As String = "CurrentSection"
Private Const OFFICE_DRAWING_FORMAT As Long = 5 ' Формат Microsoft Office Drawing Object в Excel PasteSpecial.
Private Const CLIPBOARD_BUSY As Long = &H800401D0 ' Windows не смог открыть буфер обмена.

Private mSwitching As Boolean ' Вложенный вызов не должен перезаписывать буфер и рамку.

' Копирует именованный рисунок и вписывает его в рамку CurrentSection.
' Сохраняет текущий рисунок до готовности нового. Пересчёт и ScreenUpdating не переключаются.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, selector As Range, original As Shape, frame As Shape, replacement As Shape
    Dim previousSheet As Object, pumpSheet As Object, activeSheet As Object, previousSelection As Range
    Dim pictureName As String, action As String, previousName As String, failureDetails As String
    Dim value As Variant, fitRatio As Double, copyMode As Variant
    Dim pictureWidth As Double, pictureHeight As Double, frameWidth As Double, frameHeight As Double
    Dim previousUpdating As Boolean, previousEvents As Boolean, previousCancel As XlEnableCancelKey
    Dim restoring As Boolean, usedClipboard As Boolean, changedContext As Boolean, renamed As Boolean
    Dim number As Long, description As String, attempt As Long, copyNumber As Long, copyDescription As String
    Dim frameId As Long, previousWindow As Long, oldIds() As Long, pasteStarted As Boolean

    If mSwitching Then Exit Sub
    mSwitching = True
    On Error GoTo Failed
    action = "чтение ячейки выбора иллюстрации"
    Set selector = SectionIllustrationSelector()
    value = selector.Value2
    If IsError(value) Then Err.Raise vbObjectError + 2801, , _
        "В ячейке " & IllustrationAddress(selector) & " ошибка Excel. Выберите имя рисунка из списка."
    pictureName = CStr(value)
    If Len(Trim$(pictureName)) = 0 Then Err.Raise vbObjectError + 2801, , _
        "Выберите имя рисунка из списка в ячейке " & IllustrationAddress(selector) & "."

    action = "проверка исходного рисунка"
    Set original = FindIllustration(ThisWorkbook.Worksheets(SOURCE_SHEET), pictureName)
    action = "проверка рамки CurrentSection"
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    Set frame = FindIllustration(sheet, TARGET_PICTURE)
    If sheet.Visible <> xlSheetVisible Then Err.Raise vbObjectError + 2801, , _
        "Отобразите лист «" & TARGET_SHEET & "» и повторите выбор в ячейке " & IllustrationAddress(selector) & "."
    If sheet.ProtectDrawingObjects Then
        If Not sheet.ProtectionMode Then Err.Raise vbObjectError + 2801, , _
            "Рисунки листа «" & TARGET_SHEET & "» защищены. Разрешите изменение объектов и повторите выбор."
    End If
    If Application.CutCopyMode <> False Then Err.Raise vbObjectError + 2801, , _
        "Excel копирует или перемещает ячейки. Завершите вставку либо нажмите Esc и повторите выбор."

    action = "подготовка рамки"
    frameId = frame.ID
    previousUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    previousCancel = Application.EnableCancelKey
    Set previousSheet = Application.ActiveSheet
    If TypeOf Application.Selection Is Range Then Set previousSelection = Application.Selection
    restoring = True
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler ' Esc и Ctrl+Break проходят через восстановление рамки и флагов.

    For attempt = 1 To 3
        pasteStarted = False
        action = "копирование исходного рисунка"
        usedClipboard = True
        On Error Resume Next
        Err.Clear
        original.Copy
        copyNumber = Err.Number: copyDescription = Err.Description
        On Error GoTo Failed
        If copyNumber = 0 Then
            previousWindow = ActiveIllustrationWindow()
            Set pumpSheet = Application.ActiveSheet
            copyMode = Application.CutCopyMode
            DoEvents ' Excel завершает подготовку буфера перед Paste.
            Set activeSheet = Application.ActiveSheet
            changedContext = ActiveIllustrationWindow() <> previousWindow Or Not activeSheet Is pumpSheet Or _
                Application.CutCopyMode <> copyMode
            If changedContext Then Err.Raise vbObjectError + 2801, , _
                "Во время копирования изменилось окно или состояние буфера. Повторите выбор в ячейке " & IllustrationAddress(selector) & "."
            value = selector.Value2
            If IsError(value) Then Err.Raise vbObjectError + 2801, , _
                "Выберите имя рисунка из списка в ячейке " & IllustrationAddress(selector) & "."
            If CStr(value) <> pictureName Then Err.Raise vbObjectError + 2801, , _
                "Значение выбора изменилось. Повторите выбор в ячейке " & IllustrationAddress(selector) & "."
            If Not activeSheet Is sheet Then sheet.Activate
            If Not TypeOf Application.Selection Is Range Then sheet.Cells(1, 1).Select ' PasteSpecial требует выделенную ячейку.
            action = "вставка рисунка на лист «" & TARGET_SHEET & "»"
            ReadIllustrationShapeIds sheet, oldIds
            pasteStarted = True
            On Error Resume Next
            Err.Clear
            sheet.PasteSpecial Format:=OFFICE_DRAWING_FORMAT
            copyNumber = Err.Number: copyDescription = Err.Description
            On Error GoTo Failed
            If copyNumber = 0 Then
                action = "проверка результата вставки"
                Set replacement = PastedIllustration(sheet, oldIds)
                If Not replacement Is Nothing Then Exit For
                copyNumber = 1004 ' Excel не создал рисунок либо создал неоднозначный набор фигур.
                copyDescription = "Excel не подтвердил вставку одного рисунка на лист «" & TARGET_SHEET & "»."
            End If
            If RemoveIncompletePaste(sheet, oldIds) Then Err.Raise copyNumber, , copyDescription
            pasteStarted = False
        End If
        If (copyNumber <> 1004 And copyNumber <> CLIPBOARD_BUSY) Or attempt = 3 Then Err.Raise copyNumber, , copyDescription

        previousWindow = ActiveIllustrationWindow()
        Set pumpSheet = Application.ActiveSheet
        copyMode = Application.CutCopyMode
        DoEvents ' Повтор нужен только после временного отказа буфера.
        changedContext = ActiveIllustrationWindow() <> previousWindow Or Not Application.ActiveSheet Is pumpSheet Or _
                Application.CutCopyMode <> copyMode
        If changedContext Then Err.Raise vbObjectError + 2801, , _
            "Во время копирования изменилось окно или состояние буфера. Повторите выбор в ячейке " & IllustrationAddress(selector) & "."
    Next attempt

    action = "вписывание рисунка"
    pictureWidth = original.Width: pictureHeight = original.Height
    frameWidth = frame.Width: frameHeight = frame.Height
    fitRatio = frameWidth / pictureWidth
    If pictureHeight * fitRatio > frameHeight Then fitRatio = frameHeight / pictureHeight
    With replacement.PictureFormat.Crop
        .ShapeWidth = frameWidth: .ShapeHeight = frameHeight
        .ShapeLeft = frame.Left: .ShapeTop = frame.Top
        .PictureWidth = pictureWidth * fitRatio: .PictureHeight = pictureHeight * fitRatio
        .PictureOffsetX = 0: .PictureOffsetY = 0 ' Центр картинки внутри рамки.
    End With
    replacement.Placement = frame.Placement
    replacement.LockAspectRatio = frame.LockAspectRatio
    replacement.AlternativeText = pictureName
    replacement.OnAction = vbNullString
    action = "замена рисунка CurrentSection"
    previousName = replacement.Name
    frame.Name = previousName & "_previous": renamed = True
    replacement.Name = TARGET_PICTURE
    frame.Delete
    GoTo RestoreState

Failed:
    number = Err.Number: failureDetails = Err.Description
    Select Case number
        Case vbObjectError + 2801, vbObjectError + 2802
            description = failureDetails
        Case CLIPBOARD_BUSY
            description = "Буфер обмена занят. Завершите копирование или вставку и повторите выбор в ячейке " & IllustrationAddress(selector) & "."
        Case 18
            description = "Переключение прервано. Повторите выбор в ячейке " & IllustrationAddress(selector) & "."
        Case Else
            description = "Excel не завершил действие «" & action & "» для иллюстрации «" & pictureName & "»." & _
                vbCrLf & "Повторите выбор в ячейке " & IllustrationAddress(selector) & "." & vbCrLf & vbCrLf & _
                "Код Excel " & CStr(number) & ". " & failureDetails
    End Select
    On Error Resume Next
    If pasteStarted Then
        If renamed Then
            Set frame = Nothing
            Set frame = IllustrationById(sheet, frameId)
            If Not frame Is Nothing Then frame.Name = TARGET_PICTURE
        End If
        If Not frame Is Nothing Then RemoveIncompletePaste sheet, oldIds
    End If
RestoreState:
    On Error Resume Next
    If usedClipboard And Not changedContext Then Application.CutCopyMode = False
    If restoring Then
        If Not changedContext Then
            If Not Application.ActiveSheet Is previousSheet Then previousSheet.Activate
            If Not previousSelection Is Nothing Then previousSelection.Select
        End If
        Application.EnableCancelKey = previousCancel
        Application.EnableEvents = previousEvents
        If previousUpdating Then DoEvents ' Excel обновляет рисунок и стрелку списка.
    End If
    mSwitching = False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "SwitchSectionPicture", description
End Sub

' Адрес ячейки нужен для диагностики. При недоступной ячейке возвращается имя диапазона.
Private Function IllustrationAddress(ByVal selector As Range) As String
    IllustrationAddress = SELECTOR_NAME
    If selector Is Nothing Then Exit Function
    On Error Resume Next
    IllustrationAddress = selector.Worksheet.Name & "!" & selector.Address(False, False)
End Function

' Возвращает дескриптор текущего окна для проверки контекста после DoEvents.
Private Function ActiveIllustrationWindow() As Long
    Dim window As Window
    Set window = Application.ActiveWindow
    If Not window Is Nothing Then ActiveIllustrationWindow = window.Hwnd
End Function

' Находит обычный рисунок по имени на заданном листе и объясняет, как восстановить отсутствующий объект.
Private Function FindIllustration(ByVal sheet As Worksheet, ByVal pictureName As String) As Shape
    Dim item As Shape
    On Error Resume Next
    Set item = sheet.Shapes(pictureName)
    On Error GoTo 0
    If item Is Nothing Then Err.Raise vbObjectError + 2801, , _
        "Рисунок «" & pictureName & "» на листе «" & sheet.Name & "» отсутствует. Вставьте рисунок и присвойте ему это имя."
    If item.Type <> msoPicture Then Err.Raise vbObjectError + 2801, , _
        "Объект «" & pictureName & "» на листе «" & sheet.Name & "» должен быть обычным рисунком Excel."
    Set FindIllustration = item
End Function

' Запоминает идентификаторы фигур перед Paste для проверки вставки и отката.
Private Sub ReadIllustrationShapeIds(ByVal sheet As Worksheet, ByRef ids() As Long)
    Dim item As Shape, index As Long
    ReDim ids(1 To sheet.Shapes.Count)
    For Each item In sheet.Shapes
        index = index + 1
        ids(index) = item.ID
    Next item
End Sub

' Отличает вставленную фигуру от объектов, находившихся на листе до Paste.
Private Function IsExistingIllustration(ByVal id As Long, ByRef ids() As Long) As Boolean
    Dim index As Long
    For index = LBound(ids) To UBound(ids)
        If ids(index) = id Then IsExistingIllustration = True: Exit Function
    Next index
End Function

' Обычно Paste выделяет новый рисунок. Если выделение осталось на ячейке, ищем ровно один новый рисунок по ID.
Private Function PastedIllustration(ByVal sheet As Worksheet, ByRef ids() As Long) As Shape
    Dim selected As Shape, item As Shape, found As Shape
    On Error Resume Next
    Set selected = Application.Selection.ShapeRange.Item(1)
    On Error GoTo 0
    If Not selected Is Nothing Then
        If selected.Parent Is sheet And selected.Type = msoPicture And Not IsExistingIllustration(selected.ID, ids) Then
            Set PastedIllustration = selected
            Exit Function
        End If
    End If
    For Each item In sheet.Shapes
        If item.Type = msoPicture And Not IsExistingIllustration(item.ID, ids) Then
            If Not found Is Nothing Then Exit Function ' Несколько новых рисунков не дают однозначной замены.
            Set found = item
        End If
    Next item
    Set PastedIllustration = found
End Function

' После отказа вставки удаляет только добавленные фигуры. Стрелка списка принадлежит Excel.
Private Function RemoveIncompletePaste(ByVal sheet As Worksheet, ByRef ids() As Long) As Boolean
    Dim item As Shape, index As Long
    For index = sheet.Shapes.Count To 1 Step -1
        Set item = sheet.Shapes(index)
        If item.Type <> msoFormControl And Not IsExistingIllustration(item.ID, ids) Then
            RemoveIncompletePaste = True
            item.Delete
        End If
    Next index
End Function

' Находит исходную рамку даже после переименования во время замены.
Private Function IllustrationById(ByVal sheet As Worksheet, ByVal id As Long) As Shape
    Dim item As Shape
    For Each item In sheet.Shapes
        If item.ID = id Then Set IllustrationById = item: Exit Function
    Next item
End Function

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
