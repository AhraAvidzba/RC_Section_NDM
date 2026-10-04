Attribute VB_Name = "CExcelAppStateGuard"
Option Explicit

' ==========================================================================
' Временный быстрый режим Excel
' ==========================================================================
' Класс на время длинного пользовательского сценария переводит текущий
' экземпляр Excel в режим, который меньше зависит от других открытых книг:
' отключает события, экранное обновление, предупреждения и автоматический
' пересчет. Это не меняет расчетную модель NDM, а только защищает макрос от
' лишней работы Excel-оболочки.
'
' Важно: настройки Excel.Application общие для всех книг в одном процессе
' Excel. Поэтому guard должен жить только внутри одного макроса и обязан
' восстановить исходные значения при любом завершении, включая ошибку.

Private Const XL_CALCULATION_MANUAL As Long = -4135 ' xlCalculationManual; используем число, чтобы класс работал через late binding.

Private mApplication As Object ' Экземпляр Excel, настройки которого временно изменены.
Private mActive As Boolean ' True, пока быстрый режим включен и исходные настройки нужно вернуть.
Private mScreenUpdating As Boolean ' Исходное состояние обновления экрана.
Private mEnableEvents As Boolean ' Исходное состояние событий Excel.
Private mDisplayAlerts As Boolean ' Исходное состояние предупреждений Excel.
Private mCalculation As Variant ' Исходный режим пересчета: Automatic, Manual или Semiautomatic.

' Включает быстрый режим для переданного Excel.Application.
' Метод сохраняет исходные настройки до первого изменения, чтобы Restore
' мог вернуть пользователю ровно то состояние Excel, которое было до запуска
' расчета NDM. Отказ любого setter-а сразу возвращает уже измененные поля
' и передает исходную COM-ошибку вызывающему коду без подмены.
Public Sub Enter(ByVal excelApplication As Object)
    If excelApplication Is Nothing Then Err.Raise vbObjectError + 4180, "CExcelAppStateGuard.Enter", "Excel.Application не передан."

    If mActive Then Restore

    On Error GoTo Failed
    Set mApplication = excelApplication

    mScreenUpdating = mApplication.ScreenUpdating
    mEnableEvents = mApplication.EnableEvents
    mDisplayAlerts = mApplication.DisplayAlerts
    mCalculation = mApplication.Calculation

    ' Снимок полон: с этого момента восстановление нужно даже при partial Enter.
    mActive = True
    mApplication.ScreenUpdating = False
    mApplication.EnableEvents = False
    mApplication.DisplayAlerts = False
    mApplication.Calculation = XL_CALCULATION_MANUAL
    Exit Sub
Failed:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    Restore
    Set mApplication = Nothing
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

' Возвращает Excel.Application в состояние, которое было перед Enter.
' Восстановление сделано максимально терпимым к ошибкам COM: если Excel уже
' закрывается или недоступен, guard не должен маскировать исходную ошибку
' расчета новой ошибкой восстановления интерфейса.
Public Sub Restore()
    If Not mActive Then Exit Sub
    If mApplication Is Nothing Then
        mActive = False
        Exit Sub
    End If

    On Error Resume Next
    mApplication.Calculation = mCalculation
    mApplication.DisplayAlerts = mDisplayAlerts
    mApplication.EnableEvents = mEnableEvents
    mApplication.ScreenUpdating = mScreenUpdating
    On Error GoTo 0

    mActive = False
    Set mApplication = Nothing
End Sub

' Подстраховка на случай неожиданного выхода из процедуры.
' Основной код все равно вызывает Restore явно, но деструктор помогает не
' оставить Excel в ручном пересчете или с выключенными событиями при
' нестандартном завершении объекта.
Private Sub Class_Terminate()
    Restore
End Sub
