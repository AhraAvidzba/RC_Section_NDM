VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadCombinationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение таблицы сочетаний нагрузок
' ==========================================================================
' Reader переносит строки rngLoadCombinations в CBatchSectionCalculator. Пустой
' Mx или My считается нулем, полностью нулевая строка нагрузок пропускается,
' а неполные/некорректные строки сохраняются как InvalidInput, чтобы пользователь
' увидел их в отчете.
'
' Если передан CUnitSystem, усилия сразу переводятся из INPUT-единиц и
' пользовательской системы знаков во внутреннюю систему расчетного ядра.

' Загружает сочетания из именованного диапазона книги.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3950, "CLoadCombinationReader", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3951, "CLoadCombinationReader", "Пакетный расчетчик не передан."
    LoadFromRange workbook.Names.Item("rngLoadCombinations").RefersToRange, batch, units
End Sub

' Загружает сочетания из уже переданного диапазона.
' Это используется и рабочей книгой, и тестами, чтобы проверять reader без
' ручного запуска Excel-интерфейса.
Public Sub LoadFromRange(ByVal source As Object, ByVal batch As CBatchSectionCalculator, Optional ByVal units As CUnitSystem = Nothing)
    If source Is Nothing Then Err.Raise vbObjectError + 3952, "CLoadCombinationReader", "Диапазон сочетаний нагрузок не найден."
    If batch Is Nothing Then Err.Raise vbObjectError + 3953, "CLoadCombinationReader", "Пакетный расчетчик не передан."

    Dim data As Variant
    data = source.Value2

    Dim rowIndex As Long
    ' Пустые строки оставляем пользователю для ввода, а не превращаем в ошибки.
    For rowIndex = 2 To UBound(data, 1)
        If IsEmptyCombinationRow(data, rowIndex) Then
            ' Строка свободна для будущего ввода LC.
        Else
            AddRow data, batch, rowIndex, units
        End If
    Next rowIndex
End Sub

' Читает одну строку LC и добавляет ее в batch.
' Ошибки ввода не прерывают весь расчет: строка попадает в отчет как InvalidInput,
' остальные сочетания продолжают считаться.
Private Sub AddRow(ByRef data As Variant, ByVal batch As CBatchSectionCalculator, ByVal rowIndex As Long, ByVal units As CUnitSystem)
    Dim combinationID As String
    Dim profileId As String
    Dim capacityLoadPath As String
    Dim comment As String
    combinationID = Trim$(CStr(data(rowIndex, 1)))
    profileId = Trim$(CStr(data(rowIndex, 5)))
    If UBound(data, 2) >= 7 Then
        capacityLoadPath = Trim$(CStr(data(rowIndex, 6)))
        comment = Trim$(CStr(data(rowIndex, 7)))
    Else
        capacityLoadPath = vbNullString
        comment = Trim$(CStr(data(rowIndex, 6)))
    End If

    Dim errorText As String
    Dim nValue As Double
    Dim mxValue As Double
    Dim myValue As Double
    nValue = ReadOptionalLoadDouble(data(rowIndex, 2), "N", errorText)
    mxValue = ReadOptionalLoadDouble(data(rowIndex, 3), "Mx", errorText)
    myValue = ReadOptionalLoadDouble(data(rowIndex, 4), "My", errorText)

    If Abs(nValue) + Abs(mxValue) + Abs(myValue) <= 0.000000001 Then Exit Sub

    If Len(Trim$(combinationID)) = 0 Then AppendError errorText, "CombinationID не заполнен"
    If Len(profileId) = 0 Then AppendError errorText, "ProfileId не заполнен"

    If Len(errorText) > 0 Then
        batch.AddInvalidCombination combinationID, profileId, comment, errorText, capacityLoadPath
    Else
        If Not units Is Nothing Then
            nValue = units.InputForceToInternal(nValue)
            mxValue = units.InputMomentMxToInternal(mxValue)
            myValue = units.InputMomentMyToInternal(myValue)
        End If
        batch.AddCombination combinationID, nValue, mxValue, myValue, profileId, comment, capacityLoadPath
    End If
End Sub

' Проверяет, что пользователь действительно ничего не ввел в строке LC.
Private Function IsEmptyCombinationRow(ByRef data As Variant, ByVal rowIndex As Long) As Boolean
    Dim col As Long
    For col = 1 To UBound(data, 2)
        If Len(Trim$(CStr(data(rowIndex, col)))) > 0 Then Exit Function
    Next col
    IsEmptyCombinationRow = True
End Function

' Читает нагрузку, где пустая ячейка означает 0.
' Это позволяет задавать одноосный изгиб как N + Mx при пустом My.
Private Function ReadOptionalLoadDouble(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then
        ReadOptionalLoadDouble = 0#
    ElseIf Not IsNumeric(value) Then
        AppendError errorText, fieldName & " must be numeric"
    Else
        ReadOptionalLoadDouble = CDbl(value)
    End If
End Function

' Читает обязательное число.
' Сейчас оставлено для совместимости с тестами и будущими обязательными полями.
Private Function ReadRequiredDouble(ByVal value As Variant, ByVal fieldName As String, ByRef errorText As String) As Double
    If Len(Trim$(CStr(value))) = 0 Then
        AppendError errorText, fieldName & " не заполнен"
    ElseIf Not IsNumeric(value) Then
        AppendError errorText, fieldName & " должен быть числом"
    Else
        ReadRequiredDouble = CDbl(value)
    End If
End Function

' Добавляет текст ошибки в одну строку через точку с запятой.
Private Sub AppendError(ByRef errorText As String, ByVal item As String)
    If Len(errorText) > 0 Then errorText = errorText & "; "
    errorText = errorText & item
End Sub

