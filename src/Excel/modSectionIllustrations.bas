Attribute VB_Name = "modSectionIllustrations"
Option Explicit

' ==========================================================================
' Независимая справочная картинка: без расчётных таблиц и пересчёта книги
' ==========================================================================
Private Const SOURCE_SHEET As String = "Сечения" ' Лист с четырьмя исходными рисунками, названными по типам сечения.
Private Const TARGET_SHEET As String = "Расчет" ' Лист, на котором пользователь выбирает и видит справочную картинку.
Private Const SELECTOR_CELL As String = "K2" ' Ячейка выбора; событие Workbook_SheetChange обращается к этой же ячейке.
Private Const TARGET_PICTURE As String = "CurrentSection" ' Постоянное имя пользовательской рамки независимо от выбранной картинки.

' Копирует выбранный рисунок, пропорционально вписывая его в текущую рамку.
' Исходные картинки, ячейки и расчётные настройки не меняет.
' Вызывается на листе «Расчет»: листы и ScreenUpdating не переключает, выделение восстанавливает.
Public Sub SwitchSectionPicture()
    Dim sheet As Worksheet, original As Shape, frame As Shape, replacement As Shape, previousSelection As Object
    Dim pictureName As String, fitRatio As Double, previousName As String
    Dim previousEvents As Boolean, started As Double, elapsed As Double, pasteSucceeded As Boolean
    Dim number As Long, description As String, action As String, renamed As Boolean
    previousEvents = Application.EnableEvents ' Запоминаем состояние Excel, чтобы вернуть его и при успешной вставке, и при ошибке.
    If TypeOf Application.Selection Is Range Then Set previousSelection = Application.Selection ' Запоминаем ячейку выбора до временного выделения вставленного рисунка.
    On Error GoTo Failed ' Любая ошибка до удаления старой картинки должна оставить её в книге.
    action = TARGET_SHEET & "!" & SELECTOR_CELL
    Set sheet = ThisWorkbook.Worksheets(TARGET_SHEET)
    If Not Application.ActiveSheet Is sheet Then Err.Raise vbObjectError + 2801, , "Для выбора картинки откройте лист «" & TARGET_SHEET & "»."
    pictureName = CStr(sheet.Range(SELECTOR_CELL).Value2) ' Название пункта списка совпадает с именем исходного рисунка.
    action = SOURCE_SHEET & "!" & pictureName
    Set original = ThisWorkbook.Worksheets(SOURCE_SHEET).Shapes(pictureName) ' Получаем объект без перехода на лист с исходниками.
    action = TARGET_SHEET & "!" & TARGET_PICTURE
    Set frame = sheet.Shapes(TARGET_PICTURE) ' Текущие габариты берём из объекта: пользователь может передвигать и растягивать рамку.
    If original.Type <> msoPicture Or frame.Type <> msoPicture Then _
        Err.Raise vbObjectError + 2801, , "Источник и CurrentSection должны быть рисунками."
    action = "Вписывание рисунка «" & pictureName & "»"
    fitRatio = frame.Width / original.Width ' Сначала пробуем занять всю ширину рамки, сохранив пропорции исходника.
    If original.Height * fitRatio > frame.Height Then fitRatio = frame.Height / original.Height ' Если высота не помещается, ограничиваем масштаб по высоте.
    Application.EnableEvents = False ' Временная вставка не должна повторно запускать обработчик выбора картинки.
    Application.CutCopyMode = False ' Завершаем предыдущий режим копирования Excel, прежде чем подготовить этот рисунок.
    original.Copy ' Копируем сам встроенный рисунок; расчётные листы и файлы сечения не загружаются.

    ' Excel передаёт рисунок из неактивного листа в буфер асинхронно.
    ' Коротко обрабатываем сообщения и повторяем только вставку; при готовом буфере цикл заканчивается сразу.
    started = Timer
    Do
        DoEvents
        On Error Resume Next ' Неготовый буфер даёт временную ошибку Paste; остальные операции остаются под общим обработчиком.
        Err.Clear
        sheet.Paste ' Обычная вставка сохраняет исходные данные PNG, не превращая рисунок в снимок экрана.
        pasteSucceeded = (Err.Number = 0)
        On Error GoTo Failed
        If pasteSucceeded Then Exit Do
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400# ' Timer обнуляется в полночь; ожидание всё равно ограничено половиной секунды.
    Loop While elapsed < 0.5
    If Not pasteSucceeded Then Err.Raise vbObjectError + 2801, , "Excel не успел подготовить рисунок к вставке. Повторите выбор картинки."
    Set replacement = Application.Selection.ShapeRange.Item(1) ' Paste выделяет вставленный объект; порядок Shapes не используется как признак новой копии.
    replacement.Visible = msoFalse ' Не показываем промежуточный размер нового объекта, пока не настроены его рамка и масштаб.

    ' Внешняя рамка остаётся постоянной, а изображение внутри неё получает нужный масштаб и пустые поля.
    With replacement.PictureFormat.Crop
        .ShapeWidth = frame.Width: .ShapeHeight = frame.Height ' Сохраняем пользовательские ширину и высоту области CurrentSection.
        .ShapeLeft = frame.Left: .ShapeTop = frame.Top ' Сохраняем её положение на листе.
        .PictureWidth = original.Width * fitRatio: .PictureHeight = original.Height * fitRatio ' Масштабируем обе стороны одним коэффициентом.
        .PictureOffsetX = 0: .PictureOffsetY = 0 ' Центр картинки внутри постоянной рамки, без обрезки.
    End With
    replacement.Placement = frame.Placement ' Сохраняем поведение рамки при изменении строк и столбцов.
    replacement.LockAspectRatio = frame.LockAspectRatio ' Сохраняем выбранный пользователем режим ручного изменения размеров.
    replacement.AlternativeText = pictureName ' Оставляем у результата понятное имя показанной иллюстрации.
    replacement.OnAction = vbNullString ' Справочная картинка не должна становиться кнопкой запуска макроса.
    previousName = replacement.Name ' Автоматическое имя новой копии используем как основу временного имени старой.
    frame.Name = previousName & "_previous": renamed = True ' Освобождаем постоянное имя, сохраняя возможность отката.
    replacement.Name = TARGET_PICTURE ' Следующее переключение найдёт ту же пользовательскую область по постоянному имени.
    replacement.Visible = frame.Visible ' Показываем готовую замену, если исходная рамка была видимой.
    frame.Delete ' Удаляем исходный рисунок только после успешной подготовки замены.
    GoTo RestoreState
Failed:
    number = Err.Number: description = action & ": " & Err.Description ' Сохраняем причину до попытки убрать незавершённую копию.
    On Error Resume Next ' Служебная очистка не должна затереть исходное сообщение об ошибке.
    If Not replacement Is Nothing Then replacement.Delete ' Удаляем только новую копию, а не исходную пользовательскую картинку.
    If renamed Then frame.Name = TARGET_PICTURE ' При откате возвращаем старому объекту постоянное имя.
RestoreState:
    On Error Resume Next
    Application.CutCopyMode = False ' После завершения не оставляем активный режим копирования Excel.
    If Not previousSelection Is Nothing Then previousSelection.Select ' Возвращаем выделение в ячейку, пока события Excel ещё отключены.
    Application.EnableEvents = previousEvents ' Возвращаем именно исходное состояние событий; ScreenUpdating вообще не меняли.
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "SwitchSectionPicture", description ' Событие книги покажет пользователю причину без потери текущего рисунка.
End Sub
