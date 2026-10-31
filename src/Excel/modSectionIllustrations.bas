Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая иллюстрация из именованных рисунков листа «Сечения»
' ==========================================================================
' Источники заменяются вручную; K2 содержит обычное перечисление их имён.
' Связанный рисунок показывает выбранный источник без Copy/Paste, файлов и пересчёта.
Private Const SOURCE_SHEET As String = "Сечения" ' Единственное место хранения редактируемых иллюстраций.
Private Const TARGET_SHEET As String = "Расчет"
Private Const SELECTOR_CELL As String = "K2" ' Тот же адрес используется событием книги.
Private Const TARGET_PICTURE As String = "CurrentSection" ' Пользователь перемещает и растягивает всю рамку.
Private Const FRAME_NAME As String = "CurrentSection_Frame"
Private Const VIEW_NAME As String = "CurrentSection_View"

' Показывает текущее содержимое выбранного рисунка, включая несохранённую ручную замену.
' Положение и габариты CurrentSection сохраняются; другие книги и состояния Excel не меняются.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, source As Shape, frame As Shape, view As Shape, group As Shape, area As Range
    Dim pictureName As String
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    pictureName = CStr(sheet.Range(SELECTOR_CELL).Value2)
    On Error Resume Next
    Set source = ThisWorkbook.Worksheets(SOURCE_SHEET).Shapes(pictureName)
    On Error GoTo 0
    If source Is Nothing Then Err.Raise vbObjectError + 2801, "SwitchSectionPicture", _
        "На листе «" & SOURCE_SHEET & "» не найден рисунок «" & pictureName & "»." & vbCrLf & vbCrLf & _
        "Вставьте рисунок, присвойте ему это имя или выберите другое имя в " & TARGET_SHEET & "!" & SELECTOR_CELL & "."
    If Not IsSectionPicture(source) Or source.Visible = msoFalse Then Err.Raise vbObjectError + 2801, _
        "SwitchSectionPicture", "Объект «" & pictureName & "» на листе «" & SOURCE_SHEET & "» должен быть видимым рисунком или группой."
    Set group = PictureFrame(sheet) ' Обычный рисунок CurrentSection сам становится рамкой со связанным изображением.
    Set frame = group.GroupItems(FRAME_NAME): Set view = group.GroupItems(VIEW_NAME)
    Set area = source.Parent.Range(source.TopLeftCell, source.BottomRightCell)
    FitSectionPicture source, area, frame, view
    group.AlternativeText = pictureName
End Sub

' Принимает растровые/связанные рисунки, SVG и сгруппированные иллюстрации; кнопки и диаграммы в список не входят.
Private Function IsSectionPicture(ByVal source As Shape) As Boolean
    IsSectionPicture = (source.Type = msoPicture Or source.Type = msoLinkedPicture Or source.Type = msoGraphic Or source.Type = msoGroup)
End Function

' Подготавливает прозрачную рамку вокруг обычного рисунка CurrentSection без его копирования.
' Группа из рамки и вида позволяет вписывать разные пропорции, сохраняя пользовательские габариты.
Private Function PictureFrame(ByVal sheet As Worksheet) As Shape
    Dim view As Shape, border As Shape, group As Shape, placement As Long, locked As Long
    Set view = sheet.Shapes(TARGET_PICTURE)
    If view.Type = msoGroup Then
        Set PictureFrame = view
        Exit Function
    End If
    If view.Type <> msoPicture And view.Type <> msoLinkedPicture Then Err.Raise vbObjectError + 2801, _
        "SwitchSectionPicture", "На листе «" & TARGET_SHEET & "» нужен рисунок с именем «" & TARGET_PICTURE & "»."
    placement = view.Placement: locked = view.LockAspectRatio
    Set border = sheet.Shapes.AddShape(msoShapeRectangle, view.Left, view.Top, view.Width, view.Height)
    border.Name = FRAME_NAME: border.Fill.Transparency = 1#: border.Line.Visible = msoFalse
    view.Name = VIEW_NAME: view.LockAspectRatio = msoFalse ' Размеры изображения меняются независимо от рамки.
    Set group = sheet.Shapes.Range(Array(border.Name, view.Name)).group
    group.Name = TARGET_PICTURE: group.Placement = placement: group.LockAspectRatio = locked
    Set PictureFrame = group
End Function

' Связывает вид с диапазоном под источником и обрезает захваченные ячейки за его видимыми границами.
' Масштаб обеих осей одинаковый; поля вокруг картинки остаются свободными.
Private Sub FitSectionPicture(ByVal source As Shape, ByVal area As Range, ByVal frame As Shape, ByVal view As Shape)
    Dim x As Double, y As Double, visibleWidth As Double, visibleHeight As Double, rightEdge As Double, bottomEdge As Double
    Dim leftPos As Double, topPos As Double, frameWidth As Double, frameHeight As Double, factor As Double, crop As Object
    leftPos = frame.Left: topPos = frame.Top: frameWidth = frame.Width: frameHeight = frame.Height
    x = source.Left: y = source.Top: visibleWidth = source.Width: visibleHeight = source.Height
    If source.Type = msoPicture Or source.Type = msoLinkedPicture Then
        Set crop = source.PictureFormat.Crop
        rightEdge = WorksheetFunction.Min(x + visibleWidth, x + (visibleWidth + crop.PictureWidth) / 2# + crop.PictureOffsetX)
        bottomEdge = WorksheetFunction.Min(y + visibleHeight, y + (visibleHeight + crop.PictureHeight) / 2# + crop.PictureOffsetY)
        x = WorksheetFunction.Max(x, x + (visibleWidth - crop.PictureWidth) / 2# + crop.PictureOffsetX)
        y = WorksheetFunction.Max(y, y + (visibleHeight - crop.PictureHeight) / 2# + crop.PictureOffsetY)
        visibleWidth = rightEdge - x: visibleHeight = bottomEdge - y ' Пустые поля исходной рамки не уменьшают полезный рисунок.
    End If
    If visibleWidth <= 0# Or visibleHeight <= 0# Then Err.Raise vbObjectError + 2801, "SwitchSectionPicture", "Видимая область рисунка «" & source.Name & "» пуста."
    factor = WorksheetFunction.Min(frameWidth / visibleWidth, frameHeight / visibleHeight)
    view.DrawingObject.Formula = "=" & area.Address(True, True, xlA1, True) ' Ссылка внутри ThisWorkbook, без внешних файлов.
    Set crop = view.PictureFormat.Crop
    crop.ShapeWidth = visibleWidth * factor: crop.ShapeHeight = visibleHeight * factor
    crop.ShapeLeft = leftPos + (frameWidth - visibleWidth * factor) / 2#: crop.ShapeTop = topPos + (frameHeight - visibleHeight * factor) / 2#
    crop.PictureWidth = area.width * factor: crop.PictureHeight = area.height * factor
    crop.PictureOffsetX = (area.left + area.width / 2# - x - visibleWidth / 2#) * factor
    crop.PictureOffsetY = (area.top + area.height / 2# - y - visibleHeight / 2#) * factor
End Sub
