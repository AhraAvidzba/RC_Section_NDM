Attribute VB_Name = "CTestExcelApplicationState"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: имитация настроек приложения Excel
' ==========================================================================
' Хранит состояния обновления экрана, событий, предупреждений и пересчёта.
' Позволяет вызвать заданный отказ чтения или записи свойства
' для проверки восстановления настроек классом CExcelAppStateGuard.

Public FaultProperty As String
Public FaultsRemaining As Long
Public FaultNumber As Long
Public GetterFault As String
Private mCancel As Long
Private mScreen As Boolean
Private mEvents As Boolean
Private mAlerts As Boolean
Private mCalculation As Long

' ДЛЯ ТЕСТОВ: задаёт исходный автоматический пересчёт и включённые настройки Excel.
Private Sub Class_Initialize()
    mCancel = 1: mScreen = True: mEvents = True: mAlerts = True
    mCalculation = -4105: FaultNumber = 1004
End Sub

' ДЛЯ ТЕСТОВ: вызывает заданную ошибку чтения выбранного свойства.
Private Sub BeforeGet(ByVal name As String)
    If GetterFault = name Then Err.Raise 1004, "CTestExcelApplicationState", "get " & name
End Sub

' ДЛЯ ТЕСТОВ: вызывает ошибку записи и уменьшает число оставшихся отказов, если оно положительно.
Private Sub BeforeSet(ByVal name As String)
    If FaultProperty = name And FaultsRemaining <> 0 Then
        If FaultsRemaining > 0 Then FaultsRemaining = FaultsRemaining - 1
        Err.Raise FaultNumber, "CTestExcelApplicationState", "set " & name
    End If
End Sub

' ДЛЯ ТЕСТОВ: возвращает режим прерывания макроса или заданную ошибку чтения.
Public Property Get EnableCancelKey() As Long
    BeforeGet "EnableCancelKey"
    EnableCancelKey = mCancel
End Property

' ДЛЯ ТЕСТОВ: сохраняет режим прерывания макроса после проверки отказа записи.
Public Property Let EnableCancelKey(ByVal value As Long)
    BeforeSet "EnableCancelKey"
    mCancel = value
End Property

' ДЛЯ ТЕСТОВ: возвращает состояние обновления экрана или заданную ошибку чтения.
Public Property Get ScreenUpdating() As Boolean
    BeforeGet "ScreenUpdating"
    ScreenUpdating = mScreen
End Property

' ДЛЯ ТЕСТОВ: сохраняет состояние обновления экрана после проверки отказа записи.
Public Property Let ScreenUpdating(ByVal value As Boolean)
    BeforeSet "ScreenUpdating"
    mScreen = value
End Property

' ДЛЯ ТЕСТОВ: возвращает состояние обработки событий или заданную ошибку чтения.
Public Property Get EnableEvents() As Boolean
    BeforeGet "EnableEvents"
    EnableEvents = mEvents
End Property

' ДЛЯ ТЕСТОВ: сохраняет состояние обработки событий после проверки отказа записи.
Public Property Let EnableEvents(ByVal value As Boolean)
    BeforeSet "EnableEvents"
    mEvents = value
End Property

' ДЛЯ ТЕСТОВ: возвращает состояние предупреждений Excel или заданную ошибку чтения.
Public Property Get DisplayAlerts() As Boolean
    BeforeGet "DisplayAlerts"
    DisplayAlerts = mAlerts
End Property

' ДЛЯ ТЕСТОВ: сохраняет состояние предупреждений Excel после проверки отказа записи.
Public Property Let DisplayAlerts(ByVal value As Boolean)
    BeforeSet "DisplayAlerts"
    mAlerts = value
End Property

' ДЛЯ ТЕСТОВ: возвращает режим пересчёта или заданную ошибку чтения.
Public Property Get Calculation() As Long
    BeforeGet "Calculation"
    Calculation = mCalculation
End Property

' ДЛЯ ТЕСТОВ: сохраняет режим пересчёта после проверки отказа записи.
Public Property Let Calculation(ByVal value As Long)
    BeforeSet "Calculation"
    mCalculation = value
End Property
