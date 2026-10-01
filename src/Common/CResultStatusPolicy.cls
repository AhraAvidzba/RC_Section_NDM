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
Private Const STATUS_BASEFAIL As String = "BaseFail" ' Capacity: исходная часть выбранного пути уже не проходит проверку.
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

' Нормализует уже внешний display-статус. Метод не создает CResultMeta и не
' анализирует StopReason/ResultComment: смысл результата должен приходить
' через InternalStatus + ResultCode, а неизвестная строка считается CalcErr.
Public Function ToUserStatus(ByVal statusText As String) As String
    Dim value As String
    value = Trim$(statusText)
    Select Case UCase$(value)
        Case "OK"
            ToUserStatus = STATUS_OK
        Case "FAIL"
            ToUserStatus = STATUS_FAIL
        Case "BASEFAIL"
            ToUserStatus = STATUS_BASEFAIL
        Case "NUMFAIL"
            ToUserStatus = STATUS_NUMFAIL
        Case "INPUTERR"
            ToUserStatus = STATUS_INPUTERR
        Case "CALCERR"
            ToUserStatus = STATUS_CALCERR
        Case "N/A", ""
            ToUserStatus = STATUS_NA
        Case Else
            ToUserStatus = STATUS_CALCERR
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

' Выбирает худший готовый typed результат тем же единственным приоритетом,
' что и display-агрегация. Сохраняет точный код/диагностику выбранной причины
' и независимый snapshot; комментарий уже подготовлен владельцем поддерева.
Public Function WorstResultMeta(ByVal fallbackKind As EResultKind, _
        ParamArray metas() As Variant) As CResultMeta
    Dim selected As CResultMeta, candidate As CResultMeta
    Dim i As Long
    For i = LBound(metas) To UBound(metas)
        Set candidate = Nothing
        If IsObject(metas(i)) Then Set candidate = metas(i)
        If Not candidate Is Nothing Then
            If selected Is Nothing Then
                Set selected = candidate
            ElseIf StatusPriority(ExternalStatus(candidate)) > StatusPriority(ExternalStatus(selected)) Then
                Set selected = candidate
            ElseIf candidate.InternalStatus = rsSuccessWithWarning And selected.InternalStatus = rsSuccess Then
                ' Оба исхода отображаются OK, но значимое предупреждение и его
                ' машинный код не должны теряться из-за порядка дочерних ветвей.
                Set selected = candidate
            End If
        End If
    Next i
    Dim result As CResultMeta
    If selected Is Nothing Then
        Set result = New CResultMeta
        result.SetNotApplicable fallbackKind
    Else
        Set result = selected.Clone
    End If
    Set WorstResultMeta = result
End Function

' Агрегирует уже внешние display-статусы там, где формат Results хранит их
' плоскими строковыми полями. Этот метод не выводит внутренний смысл из текста,
' а только выбирает худший из утвержденных внешних статусов.
Public Function AggregateExternalStatus(ByVal directStateStatus As String, ByVal capacityStatus As String, _
        ByVal crackStatus As String, Optional ByVal longitudinalCrackStatus As String = vbNullString, _
        Optional ByVal stabilityStatus As String = vbNullString) As String
    Dim resultStatus As String
    resultStatus = STATUS_NA
    resultStatus = Merge(resultStatus, directStateStatus)
    resultStatus = Merge(resultStatus, capacityStatus)
    resultStatus = Merge(resultStatus, crackStatus)
    resultStatus = Merge(resultStatus, longitudinalCrackStatus)
    resultStatus = Merge(resultStatus, stabilityStatus)
    AggregateExternalStatus = resultStatus
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

' Распознает только утвержденные внешние статусы численной несходимости.
' Текстовые причины solver-а не анализируются, чтобы статус не зависел от
' формулировки комментария или языка diagnostic-сообщения.
Public Function IsNumerical(ByVal statusText As String) As Boolean
    IsNumerical = (ToUserStatus(statusText) = STATUS_NUMFAIL)
End Function

' Распознает только утвержденные внешние статусы ошибки ввода/Config.
' Подстроки StopReason вроде InputError/InvalidConfiguration не являются
' источником статуса; для этого используется CResultMeta.
Public Function IsInvalid(ByVal statusText As String) As Boolean
    IsInvalid = (ToUserStatus(statusText) = STATUS_INPUTERR)
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
Public Function StatusPriority(ByVal statusText As String) As Long
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
