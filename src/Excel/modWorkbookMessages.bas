Attribute VB_Name = "modWorkbookMessages"
Option Explicit

Public Const CONTOUR_MESH_GENERATION_NOTICE As String = "БЕТОННАЯ СЕТКА АВТОМАТИЧЕСКИ СГЕНЕРИРОВАНА ПО КОНТУРАМ AUTOCAD."

' ==========================================================================
' Unicode-сообщения пользовательских сценариев Excel
' ==========================================================================
' Модуль показывает уже подготовленный текст без преобразования через ANSI.
' Это сохраняет греческие обозначения, математические знаки и адреса файлов
' независимо от системной кодовой страницы. Расчеты, статусы и комментарии
' результатов формируются ответственными слоями и здесь не изменяются.
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
