Attribute VB_Name = "modSectionOptionalData"
Option Explicit

' ==========================================================================
' Необязательные сведения об источниках элементов и оформлении сечения
' ==========================================================================
' Сохраняет плоский JSON-объект со строковыми значениями внутри книги.
' Расчётные величины остаются в числовых ячейках; ядро JSON не читает.

' Записывает пары ключ/значение. Строковые значения сохраняют имена,
' комментарии и обозначения без зависимости от десятичного разделителя Excel.
Public Function SectionOptionalJson(ParamArray pairs() As Variant) As String
    Dim i As Long, text As String
    text = "{""schema"":""1"""
    For i = LBound(pairs) To UBound(pairs) Step 2
        If i = UBound(pairs) Then Err.Raise 5, "SectionOptionalJson", "Не задано значение необязательного поля."
        text = text & "," & JsonQuoted(CStr(pairs(i))) & ":" & JsonQuoted(CStr(pairs(i + 1)))
    Next i
    SectionOptionalJson = text & "}"
End Function

' Читает только необязательное дополнение. Отсутствующая строка допустима;
' повреждённый JSON сообщает адрес и не должен вызываться слоем расчёта.
Public Function ReadSectionOptionalJson(ByVal text As String, Optional ByVal address As String = "") As Object
    Dim result As Object, position As Long, key As String, value As String
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbBinaryCompare
    If Len(Trim$(text)) = 0 Then Set ReadSectionOptionalJson = result: Exit Function
    On Error GoTo InvalidJson
    position = 1: JsonSpace text, position
    If Mid$(text, position, 1) <> "{" Then GoTo InvalidJson
    position = position + 1: JsonSpace text, position
    If Mid$(text, position, 1) <> "}" Then
        Do
            key = JsonReadString(text, position): JsonSpace text, position
            If Mid$(text, position, 1) <> ":" Then GoTo InvalidJson
            position = position + 1: JsonSpace text, position
            value = JsonReadString(text, position)
            If result.Exists(key) Then GoTo InvalidJson
            result.Add key, value
            JsonSpace text, position
            If Mid$(text, position, 1) = "}" Then Exit Do
            If Mid$(text, position, 1) <> "," Then GoTo InvalidJson
            position = position + 1: JsonSpace text, position
        Loop
    End If
    position = position + 1: JsonSpace text, position
    If position <= Len(text) Then GoTo InvalidJson
    If Not result.Exists("schema") Then GoTo InvalidJson
    If result.Item("schema") <> "1" Then GoTo InvalidJson
    Set ReadSectionOptionalJson = result
    Exit Function
InvalidJson:
    Err.Raise vbObjectError + 4382, "ReadSectionOptionalJson", _
        "Не удалось прочитать необязательные сведения сечения" & IIf(Len(address) > 0, ": " & address, "") & _
        ". Расчётные данные хранятся отдельно. Для восстановления этих сведений заново сформируйте сечение."
End Function

' Возвращает отдельное необязательное поле; его отсутствие не заменяет
' никакую расчётную величину и допускается при полном восстановлении.
Public Function SectionOptionalText(ByVal values As Object, ByVal key As String) As String
    If values.Exists(key) Then SectionOptionalText = CStr(values.Item(key))
End Function

' Сохраняет числовое значение оформления с запасом значащих цифр.
' Формат JSON остаётся одинаковым при русской и английской локали Windows.
Public Function SectionOptionalNumber(ByVal value As Double) As String
    SectionOptionalNumber = Replace$(Format$(value, "0.0000000000000000E+00"), ",", ".")
End Function

' Проверяет число оформления перед обратным преобразованием. Повреждённая
' строка не превращается в ноль; расчёт этот метод не вызывает.
Public Function ReadSectionOptionalNumber(ByVal value As String, ByVal address As String) As Double
    Dim decimalSeparator As String
    decimalSeparator = Mid$(CStr(0.5), 2, 1)
    If Len(Trim$(value)) = 0 Or InStr(value, ",") > 0 Then GoTo InvalidNumber
    If Not IsNumeric(Replace$(value, ".", decimalSeparator)) Then GoTo InvalidNumber
    ReadSectionOptionalNumber = Val(value)
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4382, "ReadSectionOptionalNumber", _
        "В необязательных сведениях подписи записано неверное число: " & address & ". Заново сформируйте сечение."
End Function

' Экранирует кавычки, обратную косую черту и управляющие символы JSON.
' Русские буквы остаются обычным Unicode-текстом.
Private Function JsonQuoted(ByVal value As String) As String
    Dim i As Long, start As Long, character As String, code As Long, escaped As String, result As String
    start = 1
    For i = 1 To Len(value)
        character = Mid$(value, i, 1): code = AscW(character)
        escaped = vbNullString
        Select Case character
            Case """", "\": escaped = "\" & character
            Case Else
                If code >= 0 And code < 32 Then
                    escaped = "\u" & Right$("0000" & Hex$(code), 4)
                End If
        End Select
        If Len(escaped) > 0 Then
            result = result & Mid$(value, start, i - start) & escaped
            start = i + 1
        End If
    Next i
    JsonQuoted = """" & result & Mid$(value, start) & """"
End Function

' Читает одну строку JSON, включая стандартные escape-последовательности.
' Незавершённые строки и неэкранированные управляющие символы отклоняются.
Private Function JsonReadString(ByVal text As String, ByRef position As Long) As String
    Dim character As String, result As String, code As Long, digits As String, i As Long
    If Mid$(text, position, 1) <> """" Then Err.Raise 5
    position = position + 1
    Do While position <= Len(text)
        character = Mid$(text, position, 1): position = position + 1
        If character = """" Then JsonReadString = result: Exit Function
        If character = "\" Then
            character = Mid$(text, position, 1): position = position + 1
            Select Case character
                Case """", "\", "/": result = result & character
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    digits = Mid$(text, position, 4)
                    If Len(digits) <> 4 Then Err.Raise 5
                    For i = 1 To 4
                        If InStr(1, "0123456789abcdef", Mid$(digits, i, 1), vbTextCompare) = 0 Then Err.Raise 5
                    Next i
                    code = CLng("&H" & digits)
                    If code > 32767 Then code = code - 65536 ' ChrW принимает знаковое 16-битное число; кодовая единица Unicode сохраняется без изменения.
                    result = result & ChrW$(code): position = position + 4
                Case Else: Err.Raise 5
            End Select
        Else
            If AscW(character) >= 0 And AscW(character) < 32 Then Err.Raise 5
            result = result & character
        End If
    Loop
    Err.Raise 5
End Function

' Пропускает только четыре пробельных символа, разрешённых JSON.
Private Sub JsonSpace(ByVal text As String, ByRef position As Long)
    Do While position <= Len(text)
        Select Case Mid$(text, position, 1)
            Case " ", vbTab, vbCr, vbLf: position = position + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub
