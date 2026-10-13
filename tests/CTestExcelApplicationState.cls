Attribute VB_Name = "CTestExcelApplicationState"
Option Explicit

' ДЛЯ ТЕСТОВ: имитирует свойства Excel и воспроизводимые отказы COM.
Public FaultProperty As String
Public FaultsRemaining As Long
Public FaultNumber As Long
Public GetterFault As String
Private mCancel As Long
Private mScreen As Boolean
Private mEvents As Boolean
Private mAlerts As Boolean
Private mCalculation As Long

Private Sub Class_Initialize()
    mCancel = 1: mScreen = True: mEvents = True: mAlerts = True
    mCalculation = -4105: FaultNumber = 1004
End Sub

Private Sub BeforeGet(ByVal name As String)
    If GetterFault = name Then Err.Raise 1004, "CTestExcelApplicationState", "get " & name
End Sub

Private Sub BeforeSet(ByVal name As String)
    If FaultProperty = name And FaultsRemaining <> 0 Then
        If FaultsRemaining > 0 Then FaultsRemaining = FaultsRemaining - 1
        Err.Raise FaultNumber, "CTestExcelApplicationState", "set " & name
    End If
End Sub

Public Property Get EnableCancelKey() As Long
    BeforeGet "EnableCancelKey"
    EnableCancelKey = mCancel
End Property

Public Property Let EnableCancelKey(ByVal value As Long)
    BeforeSet "EnableCancelKey"
    mCancel = value
End Property

Public Property Get ScreenUpdating() As Boolean
    BeforeGet "ScreenUpdating"
    ScreenUpdating = mScreen
End Property

Public Property Let ScreenUpdating(ByVal value As Boolean)
    BeforeSet "ScreenUpdating"
    mScreen = value
End Property

Public Property Get EnableEvents() As Boolean
    BeforeGet "EnableEvents"
    EnableEvents = mEvents
End Property

Public Property Let EnableEvents(ByVal value As Boolean)
    BeforeSet "EnableEvents"
    mEvents = value
End Property

Public Property Get DisplayAlerts() As Boolean
    BeforeGet "DisplayAlerts"
    DisplayAlerts = mAlerts
End Property

Public Property Let DisplayAlerts(ByVal value As Boolean)
    BeforeSet "DisplayAlerts"
    mAlerts = value
End Property

Public Property Get Calculation() As Long
    BeforeGet "Calculation"
    Calculation = mCalculation
End Property

Public Property Let Calculation(ByVal value As Long)
    BeforeSet "Calculation"
    mCalculation = value
End Property
