Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая справочная иллюстрация без буфера обмена и пересчёта книги
' ==========================================================================
' Сборщик встраивает четыре PNG из tools/build_workbook/SectionIllustrations
' в группу CurrentSection. Процедура меняет только масштаб и видимость её рисунков.
Private Const TARGET_SHEET As String = "Расчет" ' Лист выбора иллюстрации; расчётный тип сечения здесь не задаётся.
Private Const SELECTOR_CELL As String = "K2" ' Адрес совпадает с адресом в событии Workbook_SheetChange.
Private Const TARGET_PICTURE As String = "CurrentSection" ' Всю группу можно вручную перемещать и менять её размеры.

' Пропорционально вписывает выбранный встроенный рисунок в текущую рамку группы.
' Не копирует объекты, не переключает листы, не меняет выделение и состояния Excel.
' Для обновления самих изображений используются PNG сборщика и пересборка книги.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, frame As Shape, pictures(0 To 3) As Shape, crop As Object
    Dim names As Variant, pictureName As String, dimensions As Variant
    Dim index As Long, selected As Long, fitHeight As Double, aspect As Double
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    pictureName = CStr(sheet.Range(SELECTOR_CELL).Value2) ' Список содержит точные идентификаторы четырёх сечений.
    names = Array("Circle", "Rectangles", "PointedSection", "HollowRectangle")
    selected = -1 ' Неизвестный пункт не должен скрыть уже показанный рисунок.
    For index = 0 To 3
        If pictureName = CStr(names(index)) Then selected = index
    Next index
    If selected < 0 Then Err.Raise vbObjectError + 2801, "SwitchSectionPicture", _
        "Выберите название иллюстрации из списка " & TARGET_SHEET & "!" & SELECTOR_CELL & "."
    Set frame = sheet.Shapes(TARGET_PICTURE) ' Прозрачная рамка внутри группы сохраняет её полные габариты.
    If frame.Type <> msoGroup Then Err.Raise vbObjectError + 2801, "SwitchSectionPicture", _
        "Не найдена группа иллюстраций «" & TARGET_PICTURE & "». Восстановите её из актуальной сборки."

    ' Сначала проверяем весь комплект: отсутствующий рисунок не должен погасить текущую иллюстрацию.
    For index = 0 To 3
        Set pictures(index) = frame.GroupItems.Item(TARGET_PICTURE & "_" & CStr(names(index)))
        If pictures(index).Type <> msoPicture Then Err.Raise vbObjectError + 2801, "SwitchSectionPicture", _
            "В группе «" & TARGET_PICTURE & "» требуется рисунок «" & CStr(names(index)) & "»."
    Next index
    dimensions = Split(pictures(selected).AlternativeText, "|") ' Исходные размеры PNG в пикселях: ширина|высота.
    aspect = CDbl(dimensions(0)) / CDbl(dimensions(1)) ' Пропорции исходника не зависят от растяжения группы.
    Set crop = pictures(selected).PictureFormat.Crop
    fitHeight = pictures(selected).Height ' Сначала ограничиваем масштаб высотой рамки рисунка.
    If fitHeight * aspect > pictures(selected).Width Then fitHeight = pictures(selected).Width / aspect ' При узкой рамке ограничиваем шириной.
    crop.PictureWidth = fitHeight * aspect: crop.PictureHeight = fitHeight ' Один масштаб обеих сторон; свободное место остаётся полями.
    crop.PictureOffsetX = 0: crop.PictureOffsetY = 0 ' Центр изображения совпадает с центром пользовательской рамки.
    For index = 0 To 3
        pictures(index).Visible = (index = selected) ' Меняем видимость готовых объектов, не создавая и не удаляя их.
    Next index
    frame.AlternativeText = pictureName ' Свойства группы показывают выбранную иллюстрацию.
End Sub
