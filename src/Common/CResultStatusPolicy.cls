VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResultStatusPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ПОЛИТИКА ВНЕШНИХ СТАТУСОВ РЕЗУЛЬТАТОВ
' ==========================================================================
' CResultStatusPolicy является единственным местом, где типизированные
' InternalStatus + ResultCode + ResultKind превращаются в пользовательский
' display-статус. Класс не читает Excel, не запускает расчеты и не анализирует
' численные диаграммы: он только применяет согласованный словарь вывода.

Private Const STATUS_OK As String = "OK" ' Проверка выполнена и прошла.
Private Const STATUS_FAIL As String = "FAIL" ' Результат получен, но инженерная проверка не проходит.
Private Const STATUS_BASEFAIL As String = "BaseFail" ' Capacity: базовая точка траектории уже за пределом.
Private Const STATUS_NUMFAIL As String = "NumFail" ' Численная несходимость solver/search.
Private Const STATUS_INPUTERR As String = "InputErr" ' Ошибка пользовательских данных или Config.
Private Const STATUS_CALCERR As String = "CalcErr" ' Внутренняя ошибка программы или несформированный result.
Private Const STATUS_NA As String = "N/A" ' Расчет не запрошен, неприменим или заблокирован зависимостью.

Public Property Get OK() As String
    OK = STATUS_OK
End Property

Public Property Get Fail() As String
    Fail = STATUS_FAIL
End Property

Public Property Get BaseFail() As String
    BaseFail = STATUS_BASEFAIL
End Property

Public Property Get NumFail() As String
    NumFail = STATUS_NUMFAIL
End Property

Public Property Get InputErr() As String
    InputErr = STATUS_INPUTERR
End Property

Public Property Get CalcErr() As String
    CalcErr = STATUS_CALCERR
End Property

Public Property Get NotApplicable() As String
    NotApplicable = STATUS_NA
End Property

' Возвращает внешний статус по полной meta-структуре результата.
Public Function ExternalStatus(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then
        ExternalStatus = STATUS_CALCERR
        Exit Function
    End If

    If meta.ResultKind = rkCapacity And meta.InternalStatus = rsCheckFailed And _
            meta.ResultCode = rcInitialStateBeyondLimit Then
        ExternalStatus = STATUS_BASEFAIL
        Exit Function
    End If

    Select Case meta.InternalStatus
        Case rsUnset
            ExternalStatus = STATUS_CALCERR
        Case rsNotRequested, rsNotApplicable, rsBlockedByDependency
            ExternalStatus = STATUS_NA
        Case rsSuccess, rsSuccessWithWarning
            ExternalStatus = STATUS_OK
        Case rsCheckFailed
            ExternalStatus = STATUS_FAIL
        Case rsNumericalFailure
            ExternalStatus = STATUS_NUMFAIL
        Case rsInvalidInput, rsInvalidConfiguration
            ExternalStatus = STATUS_INPUTERR
        Case rsInternalError
            ExternalStatus = STATUS_CALCERR
        Case Else
            ExternalStatus = STATUS_CALCERR
    End Select
End Function

' Совместимый перевод старой строковой причины в meta. Метод нужен только для
' переходного этапа: новые расчетные блоки должны формировать CResultMeta сами.
Public Function MetaFromLegacyStatus(ByVal statusText As String, ByVal resultKind As EResultKind, _
        Optional ByVal resultComment As String = vbNullString) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    Select Case ToUserStatus(statusText)
        Case STATUS_OK
            meta.SetResult rsSuccess, rcCheckPassed, resultKind, resultComment
        Case STATUS_FAIL
            meta.SetResult rsCheckFailed, rcCheckFailed, resultKind, resultComment
        Case STATUS_BASEFAIL
            meta.SetResult rsCheckFailed, rcInitialStateBeyondLimit, resultKind, resultComment
        Case STATUS_NUMFAIL
            meta.SetResult rsNumericalFailure, rcNumericalFailure, resultKind, resultComment
        Case STATUS_INPUTERR
            meta.SetResult rsInvalidInput, rcInvalidInput, resultKind, resultComment
        Case STATUS_CALCERR
            meta.SetResult rsInternalError, rcInternalError, resultKind, resultComment
        Case Else
            meta.SetNotApplicable resultKind, resultComment
    End Select

    Set MetaFromLegacyStatus = meta
End Function

' Переводит исторический текстовый статус в утвержденный внешний словарь.
Public Function ToUserStatus(ByVal statusText As String) As String
    Dim value As String
    value = Trim$(statusText)
    Select Case UCase$(value)
        Case "OK"
            ToUserStatus = STATUS_OK
        Case "FAIL", "STRENGTHFAILED", "STRAINLIMITEXCEEDED"
            ToUserStatus = STATUS_FAIL
        Case "BASEFAIL"
            ToUserStatus = STATUS_BASEFAIL
        Case "NUMFAIL", "NUMERICALFAILURE", "SINGULARTANGENT"
            ToUserStatus = STATUS_NUMFAIL
        Case "INPUTERR", "INPUTERROR", "INVALIDINPUT", "INVALIDCONFIGURATION"
            ToUserStatus = STATUS_INPUTERR
        Case "CALCERR", "INTERNALERROR"
            ToUserStatus = STATUS_CALCERR
        Case "N/A", "NA", "NOTCALCULATED", "NOTCALCUL.", "SKIPPEDZEROMOMENT", "SKIPPEDZEROBASEMOMENT", ""
            ToUserStatus = STATUS_NA
        Case Else
            If IsInvalid(statusText) Then
                ToUserStatus = STATUS_INPUTERR
            ElseIf IsNumerical(statusText) Then
                ToUserStatus = STATUS_NUMFAIL
            Else
                ToUserStatus = STATUS_FAIL
            End If
    End Select
End Function

' Агрегирует несколько meta-результатов в худший внешний статус LC.
' N/A не ухудшает итог и не перебивает реальные результаты.
Public Function AggregateMeta(ByVal directMeta As CResultMeta, ByVal capacityMeta As CResultMeta, _
        ByVal crackMeta As CResultMeta, Optional ByVal longitudinalMeta As CResultMeta = Nothing, _
        Optional ByVal stabilityMeta As CResultMeta = Nothing) As String
    Dim resultStatus As String
    resultStatus = STATUS_NA
    If Not directMeta Is Nothing Then resultStatus = Merge(resultStatus, ExternalStatus(directMeta))
    If Not capacityMeta Is Nothing Then resultStatus = Merge(resultStatus, ExternalStatus(capacityMeta))
    If Not crackMeta Is Nothing Then resultStatus = Merge(resultStatus, ExternalStatus(crackMeta))
    If Not longitudinalMeta Is Nothing Then resultStatus = Merge(resultStatus, ExternalStatus(longitudinalMeta))
    If Not stabilityMeta Is Nothing Then resultStatus = Merge(resultStatus, ExternalStatus(stabilityMeta))
    AggregateMeta = resultStatus
End Function

' Историческое агрегирование строк оставлено для переходных участков и тестов.
Public Function AggregateText(ByVal directStateStatus As String, ByVal capacityStatus As String, _
        ByVal crackStatus As String, Optional ByVal longitudinalCrackStatus As String = vbNullString, _
        Optional ByVal stabilityStatus As String = vbNullString) As String
    Dim resultStatus As String
    resultStatus = STATUS_NA
    resultStatus = Merge(resultStatus, directStateStatus)
    resultStatus = Merge(resultStatus, capacityStatus)
    resultStatus = Merge(resultStatus, crackStatus)
    resultStatus = Merge(resultStatus, longitudinalCrackStatus)
    resultStatus = Merge(resultStatus, stabilityStatus)
    AggregateText = resultStatus
End Function

' True для расчетов, которые дали физически интерпретируемый результат.
Public Function IsFinished(ByVal statusText As String) As Boolean
    statusText = ToUserStatus(statusText)
    IsFinished = (statusText = STATUS_OK Or statusText = STATUS_FAIL Or statusText = STATUS_BASEFAIL)
End Function

' True для итогового статуса, который должен считаться проблемным.
Public Function IsFailedOverall(ByVal statusText As String) As Boolean
    statusText = ToUserStatus(statusText)
    IsFailedOverall = (statusText = STATUS_FAIL Or statusText = STATUS_BASEFAIL Or _
        statusText = STATUS_NUMFAIL Or statusText = STATUS_INPUTERR Or statusText = STATUS_CALCERR)
End Function

' Распознает численную несходимость в старых StopReason/status-строках.
Public Function IsNumerical(ByVal statusText As String) As Boolean
    IsNumerical = (InStr(1, statusText, STATUS_NUMFAIL, vbTextCompare) > 0 Or _
        InStr(1, statusText, "NumericalFailure", vbTextCompare) > 0 Or _
        InStr(1, statusText, "SingularTangent", vbTextCompare) > 0)
End Function

' Распознает ошибку исходных данных или конфигурации в старых строках.
Public Function IsInvalid(ByVal statusText As String) As Boolean
    IsInvalid = (InStr(1, statusText, STATUS_INPUTERR, vbTextCompare) > 0 Or _
        InStr(1, statusText, "InputError", vbTextCompare) > 0 Or _
        InStr(1, statusText, "InvalidInput", vbTextCompare) > 0 Or _
        InStr(1, statusText, "InvalidConfiguration", vbTextCompare) > 0)
End Function

' Объединяет два внешних статуса по единому приоритету.
Private Function Merge(ByVal currentStatus As String, ByVal candidateStatus As String) As String
    candidateStatus = ToUserStatus(candidateStatus)
    currentStatus = ToUserStatus(currentStatus)
    If candidateStatus = STATUS_NA Then
        Merge = currentStatus
    ElseIf currentStatus = STATUS_NA Then
        Merge = candidateStatus
    ElseIf StatusPriority(candidateStatus) > StatusPriority(currentStatus) Then
        Merge = candidateStatus
    Else
        Merge = currentStatus
    End If
End Function

' Численный приоритет нужен только для выбора худшего OverallStatus.
Private Function StatusPriority(ByVal statusText As String) As Long
    Select Case ToUserStatus(statusText)
        Case STATUS_CALCERR
            StatusPriority = 6
        Case STATUS_INPUTERR
            StatusPriority = 5
        Case STATUS_NUMFAIL
            StatusPriority = 4
        Case STATUS_BASEFAIL
            StatusPriority = 3
        Case STATUS_FAIL
            StatusPriority = 2
        Case STATUS_OK
            StatusPriority = 1
        Case Else
            StatusPriority = 0
    End Select
End Function
