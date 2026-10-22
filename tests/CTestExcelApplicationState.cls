Attribute VB_Name = "CTestExcelApplicationState"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: имитация настроек приложения Excel
' ==========================================================================
' Хранит состояния обновления экрана, событий, предупреждений и пересчёта.
' Позволяет вызвать заданный отказ чтения или записи свойства
' для проверки восстановления настроек классом CExcelAppStateGuard.

Public FaultProperty As String ' Имя свойства, при записи которого нужно имитировать отказ.
Public FaultsRemaining As Long ' Число оставшихся имитируемых отказов.
Public FaultNumber As Long ' Номер имитируемой ошибки.
Public TestNativeApplication As Object ' Необязательный реальный Excel: свойства и Calculate делегируются только в приватной книге native-теста.
Public TestCalculateCalls As Long ' Число явных Application.Calculate; пересчёт листа Results сюда не входит.
Public TestCalculateFault As Long ' Искусственная ошибка самого пересчёта до первого чтения входных таблиц.
Public TestCalculateGuarded As Boolean ' Все вызовы Calculate произошли при Manual и отключённых событиях/экране/окнах.
Public GetterFault As String ' Имя свойства, при чтении которого нужно имитировать отказ.
Private mCancel As Long ' Текущее подставное значение режима обработки прерывания.
Private mScreen As Boolean ' Текущее подставное значение обновления экрана.
Private mEvents As Boolean ' Текущее подставное значение обработки событий.
Private mAlerts As Boolean ' Текущее подставное значение показа предупреждений.
Private mCalculation As Long ' Текущий подставной режим пересчёта Excel.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задаёт исходный автоматический пересчёт и включённые настройки Excel.
Private Sub Class_Initialize()
    mCancel = 1: mScreen = True: mEvents = True: mAlerts = True
    mCalculation = -4105: FaultNumber = 1004: TestCalculateGuarded = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' вызывает заданную ошибку чтения выбранного свойства.
Private Sub TestBeforeGet(ByVal name As String)
    If GetterFault = name Then Err.Raise 1004, "CTestExcelApplicationState", "get " & name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' вызывает ошибку записи и уменьшает число оставшихся отказов, если оно положительно.
Private Sub TestBeforeSet(ByVal name As String)
    If FaultProperty = name And FaultsRemaining <> 0 Then
        If FaultsRemaining > 0 Then FaultsRemaining = FaultsRemaining - 1
        Err.Raise FaultNumber, "CTestExcelApplicationState", "set " & name
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает режим прерывания макроса или заданную ошибку чтения.
Public Property Get EnableCancelKey() As Long
    TestBeforeGet "EnableCancelKey"
    If TestNativeApplication Is Nothing Then EnableCancelKey = mCancel Else EnableCancelKey = TestNativeApplication.EnableCancelKey
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет режим прерывания макроса после проверки отказа записи.
Public Property Let EnableCancelKey(ByVal value As Long)
    TestBeforeSet "EnableCancelKey"
    If TestNativeApplication Is Nothing Then mCancel = value Else TestNativeApplication.EnableCancelKey = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает состояние обновления экрана или заданную ошибку чтения.
Public Property Get ScreenUpdating() As Boolean
    TestBeforeGet "ScreenUpdating"
    If TestNativeApplication Is Nothing Then ScreenUpdating = mScreen Else ScreenUpdating = TestNativeApplication.ScreenUpdating
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет состояние обновления экрана после проверки отказа записи.
Public Property Let ScreenUpdating(ByVal value As Boolean)
    TestBeforeSet "ScreenUpdating"
    If TestNativeApplication Is Nothing Then mScreen = value Else TestNativeApplication.ScreenUpdating = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает состояние обработки событий или заданную ошибку чтения.
Public Property Get EnableEvents() As Boolean
    TestBeforeGet "EnableEvents"
    If TestNativeApplication Is Nothing Then EnableEvents = mEvents Else EnableEvents = TestNativeApplication.EnableEvents
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет состояние обработки событий после проверки отказа записи.
Public Property Let EnableEvents(ByVal value As Boolean)
    TestBeforeSet "EnableEvents"
    If TestNativeApplication Is Nothing Then mEvents = value Else TestNativeApplication.EnableEvents = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает состояние предупреждений Excel или заданную ошибку чтения.
Public Property Get DisplayAlerts() As Boolean
    TestBeforeGet "DisplayAlerts"
    If TestNativeApplication Is Nothing Then DisplayAlerts = mAlerts Else DisplayAlerts = TestNativeApplication.DisplayAlerts
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет состояние предупреждений Excel после проверки отказа записи.
Public Property Let DisplayAlerts(ByVal value As Boolean)
    TestBeforeSet "DisplayAlerts"
    If TestNativeApplication Is Nothing Then mAlerts = value Else TestNativeApplication.DisplayAlerts = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает режим пересчёта или заданную ошибку чтения.
Public Property Get Calculation() As Long
    TestBeforeGet "Calculation"
    If TestNativeApplication Is Nothing Then Calculation = mCalculation Else Calculation = TestNativeApplication.Calculation
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет режим пересчёта после проверки отказа записи.
Public Property Let Calculation(ByVal value As Long)
    TestBeforeSet "Calculation"
    If TestNativeApplication Is Nothing Then mCalculation = value Else TestNativeApplication.Calculation = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Имитирует обязательный контракт Excel.Application.Calculate и считает явные
' актуализации. В native-сценарии пересчитывает настоящую приватную книгу;
' искусственный отказ позволяет проверить cleanup до чтения настроек.
Public Sub Calculate()
    TestCalculateCalls = TestCalculateCalls + 1
    TestCalculateGuarded = TestCalculateGuarded And (Calculation = -4135 And Not EnableEvents And Not ScreenUpdating And Not DisplayAlerts)
    If TestCalculateFault <> 0 Then Err.Raise TestCalculateFault, "CTestExcelApplicationState.Calculate", "Непредвиденный отказ пересчёта для проверки восстановления Excel."
    If Not TestNativeApplication Is Nothing Then TestNativeApplication.Calculate
End Sub
