Attribute VB_Name = "CExcelAppStateGuard"
Option Explicit

' ==========================================================================
' Временные настройки Excel на время макроса
' ==========================================================================
' Сохраняет настройки приложения, отключает обновления на время работы
' и восстанавливает исходные значения при выходе, в том числе после ошибки.
' Настройки Application общие для всех книг одного экземпляра Excel.
' Esc/Ctrl+Break передаются обработчику как ошибка 18; команда End, сброс
' проекта в VBE и завершение процесса обходят обычное восстановление VBA.

Private mApplication As Object ' Экземпляр Excel, настройки которого временно изменены.
Private mValues(0 To 4) As Variant ' Исходные EnableCancelKey, ScreenUpdating, EnableEvents, DisplayAlerts и Calculation.
Private mPending(0 To 4) As Boolean ' Какие из пяти настроек ещё нужно восстановить после работы макроса.
Private mActive As Boolean ' Признак сохранённых настроек, восстановление которых ещё не завершено.

' Сначала сохраняет полный снимок. При частичном отказе Enter возвращает уже
' изменённые свойства и сохраняет исходную ошибку, включая ошибку прерывания.
Public Sub Enter(ByVal excelApplication As Object)
    If excelApplication Is Nothing Then Err.Raise vbObjectError + 4180, "CExcelAppStateGuard.Enter", "Excel.Application не передан."
    If mActive Then Restore True

    Dim i As Long, fastValues As Variant
    On Error GoTo Failed
    Set mApplication = excelApplication
    For i = 0 To 4
        mValues(i) = ReadProperty(i)
    Next i
    fastValues = Array(2, False, False, False, -4135) ' xlErrorHandler = 2; xlCalculationManual = -4135.
    mActive = True
    For i = 0 To 4
        mPending(i) = Not IsError(mValues(i)) ' Без открытых книг Excel может вернуть ошибку вместо режима пересчёта. Такое значение нельзя восстановить присваиванием; сохраняем фактическую причину ошибки записи свойства.
        WriteProperty i, fastValues(i)
    Next i
    Exit Sub
Failed:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    Restore
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

' Пробует вернуть каждое свойство независимо от ошибок остальных. При
' временном COM-отказе повторяет попытку; невосстановленный снимок сохраняет
' для следующего вызова или деструктора. На штатном выходе raiseOnFailure=True
' не позволяет сообщить об успешном завершении с выключенным интерфейсом.
Public Sub Restore(Optional ByVal raiseOnFailure As Boolean = False)
    If Not mActive Then Exit Sub
    Dim attempt As Long, i As Long, failure As String
    For attempt = 1 To 2
        failure = vbNullString
        On Error Resume Next
        For i = 4 To 0 Step -1 ' Пересчёт возвращаем до событий, обновление экрана — после событий, исходный режим прерывания — последним.
            If mPending(i) Then
                Err.Clear
                WriteProperty i, mValues(i)
                If Err.Number = 0 Then
                    mPending(i) = False
                Else
                    failure = failure & PropertyName(i) & ": " & Err.Description & vbCrLf
                End If
            End If
        Next i
        On Error GoTo 0
        If Len(failure) = 0 Then
            mActive = False
            Set mApplication = Nothing
            Exit Sub
        End If
    Next attempt
    If raiseOnFailure Then Err.Raise vbObjectError + 4181, "CExcelAppStateGuard.Restore", _
        "Excel не смог восстановить настройки после макроса:" & vbCrLf & failure
End Sub

' Возвращает имя настройки Excel по индексу для сообщения об ошибке восстановления.
Private Function PropertyName(ByVal index As Long) As String
    Select Case index
        Case 0: PropertyName = "EnableCancelKey"
        Case 1: PropertyName = "ScreenUpdating"
        Case 2: PropertyName = "EnableEvents"
        Case 3: PropertyName = "DisplayAlerts"
        Case 4: PropertyName = "Calculation"
    End Select
End Function

' Читает исходное значение настройки Excel, сохраняя обычное поведение и ошибки COM.
Private Function ReadProperty(ByVal index As Long) As Variant
    Select Case index
        Case 0: ReadProperty = mApplication.EnableCancelKey
        Case 1: ReadProperty = mApplication.ScreenUpdating
        Case 2: ReadProperty = mApplication.EnableEvents
        Case 3: ReadProperty = mApplication.DisplayAlerts
        Case 4: ReadProperty = mApplication.Calculation
    End Select
End Function

' Записывает значение выбранной настройки Excel; ошибку передаёт вызывающему обработчику.
Private Sub WriteProperty(ByVal index As Long, ByVal value As Variant)
    Select Case index
        Case 0: mApplication.EnableCancelKey = value
        Case 1: mApplication.ScreenUpdating = value
        Case 2: mApplication.EnableEvents = value
        Case 3: mApplication.DisplayAlerts = value
        Case 4: mApplication.Calculation = value
    End Select
End Sub

' Повторяет восстановление оставшихся настроек при освобождении объекта.
Private Sub Class_Terminate()
    Restore
End Sub
