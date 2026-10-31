Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая справочная картинка: без расчётных таблиц и пересчёта книги
' ==========================================================================
Private Const SOURCE_SHEET As String = "Сечения"
Private Const TARGET_SHEET As String = "Расчет"
Private Const SELECTOR_CELL As String = "K2"
Private Const TARGET_PICTURE As String = "CurrentSection"

' Копирует выбранный рисунок, пропорционально вписывая его в текущую рамку.
' Исходные картинки, ячейки и расчётные настройки не меняет.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, original As Shape, frame As Shape, replacement As Shape
    Dim pictureName As String, fitRatio As Double, previousName As String
    Dim previousUpdating As Boolean, previousEvents As Boolean, previousSelection As Object, previousSheet As Object
    Dim number As Long, description As String, action As String, renamed As Boolean
    previousUpdating = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Set previousSheet = Application.ActiveSheet
    If TypeOf Application.Selection Is Range Then Set previousSelection = Application.Selection
    On Error GoTo Failed
    action = TARGET_SHEET & "!" & SELECTOR_CELL
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    pictureName = CStr(sheet.Range(SELECTOR_CELL).Value2)
    action = SOURCE_SHEET & "!" & pictureName
    Set original = ThisWorkbook.Worksheets(SOURCE_SHEET).Shapes(pictureName)
    action = TARGET_SHEET & "!" & TARGET_PICTURE
    Set frame = sheet.Shapes(TARGET_PICTURE)
    If original.Type <> msoPicture Or frame.Type <> msoPicture Then _
        Err.Raise vbObjectError + 2801, , "Источник и CurrentSection должны быть рисунками."
    action = "Вписывание рисунка «" & pictureName & "»"
    fitRatio = frame.Width / original.Width
    If original.Height * fitRatio > frame.Height Then fitRatio = frame.Height / original.Height
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    original.Parent.Activate
    original.Copy
    sheet.Activate ' Вставка выполняется после копирования и на целевом активном листе.
    sheet.Paste
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
    Application.CutCopyMode = False
    If Not Application.ActiveSheet Is previousSheet Then previousSheet.Activate
    If Not previousSelection Is Nothing Then previousSelection.Select
    Application.ScreenUpdating = previousUpdating
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "SwitchSectionPicture", description
End Sub
