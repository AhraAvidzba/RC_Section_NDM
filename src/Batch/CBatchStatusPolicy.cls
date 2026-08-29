VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBatchStatusPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Пользовательские статусы batch-расчета
' ==========================================================================
' CBatchStatusPolicy хранит единый короткий словарь статусов, которые видит
' пользователь в Summary и Results, и правила перехода от внутренних причин
' solver-ов к этому словарю.
'
' Внутренний слой может хранить подробные StopReason, LimitState и другие
' технические строки. Наружный слой, который выводится пользователю, всегда
' сводится только к OK, FAIL, NumFail, InputErr или N/A. Класс не запускает
' расчет и не знает о сочетаниях нагрузок: его граница ответственности -
' только перевод и агрегирование статусов.

Private Const STATUS_OK As String = "OK" ' Расчетная ветка завершилась без ошибки.
Private Const STATUS_FAIL As String = "FAIL" ' Равновесие найдено, но физический предел или проверка не проходят.
Private Const STATUS_NUMFAIL As String = "NumFail" ' Solver не нашел равновесие в заданных численных ограничениях.
Private Const STATUS_INPUTERR As String = "InputErr" ' Ошибка исходных данных, а не численная несходимость.
Private Const STATUS_NA As String = "N/A" ' Расчетная проверка не относится к этому LC/режиму.
Private Const STATUS_NOT_CALCULATED_SHORT As String = "NotCalcul." ' Короткая подпись для служебных нерассчитанных полей.

' Возвращает статус успешной расчетной ветки.
Public Property Get OK() As String
    OK = STATUS_OK
End Property

' Возвращает статус физически непрошедшей проверки при найденном равновесии.
Public Property Get Fail() As String
    Fail = STATUS_FAIL
End Property

' Возвращает статус настоящей численной несходимости.
Public Property Get NumFail() As String
    NumFail = STATUS_NUMFAIL
End Property

' Возвращает статус ошибки пользовательских исходных данных.
Public Property Get InputErr() As String
    InputErr = STATUS_INPUTERR
End Property

' Возвращает статус неприменимой или не запущенной проверки.
Public Property Get NotApplicable() As String
    NotApplicable = STATUS_NA
End Property

' Собирает OverallStatus из статусов применимых расчетных разделов.
' N/A не ухудшает итог: это означает, что ветка не относится к LC или режиму.
Public Function Aggregate(ByVal directStateStatus As String, ByVal capacityStatus As String, _
        ByVal crackStatus As String) As String
    Dim resultStatus As String
    resultStatus = STATUS_NA
    resultStatus = Merge(resultStatus, directStateStatus)
    resultStatus = Merge(resultStatus, capacityStatus)
    resultStatus = Merge(resultStatus, crackStatus)
    Aggregate = resultStatus
End Function

' Переводит внутренний текст расчетной ветки в один из пяти пользовательских
' статусов. Этот метод используется там, где нужна именно оценка результата:
' агрегирование OverallStatus, подсветка строк и проверка завершенности ветки.
Public Function ToUserStatus(ByVal statusText As String) As String
    Dim value As String
    value = Trim$(statusText)
    Select Case UCase$(value)
        Case "OK"
            ToUserStatus = STATUS_OK
        Case "FAIL", "STRENGTHFAILED", "STRAINLIMITEXCEEDED"
            ToUserStatus = STATUS_FAIL
        Case "NUMFAIL", "NUMERICALFAILURE", "SINGULARTANGENT"
            ToUserStatus = STATUS_NUMFAIL
        Case "INPUTERR", "INPUTERROR", "INVALIDINPUT"
            ToUserStatus = STATUS_INPUTERR
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

' Совместимое имя для старых вызовов внутри проекта. Новая расчетная логика
' должна использовать ToUserStatus, чтобы было видно различие внутреннего
' технического текста и пользовательского статуса.
Public Function Normalize(ByVal statusText As String) As String
    Normalize = ToUserStatus(statusText)
End Function

' Готовит текст для ячейки Results/Summary. В отличие от ToUserStatus, метод
' не превращает неизвестную строку в FAIL: например, CapacityLimitState
' "ConcreteStrainLimit" остается инженерной причиной, а не статусом.
Public Function DisplayStatus(ByVal statusText As String) As String
    Select Case UCase$(Trim$(statusText))
        Case "OK"
            DisplayStatus = STATUS_OK
        Case "FAIL", "STRENGTHFAILED", "STRAINLIMITEXCEEDED"
            DisplayStatus = STATUS_FAIL
        Case "NUMFAIL", "NUMERICALFAILURE", "SINGULARTANGENT"
            DisplayStatus = STATUS_NUMFAIL
        Case "INPUTERR", "INPUTERROR", "INVALIDINPUT"
            DisplayStatus = STATUS_INPUTERR
        Case "N/A", "NA", "NOTCALCULATED", "NOTCALCUL.", "SKIPPEDZEROMOMENT", "SKIPPEDZEROBASEMOMENT", ""
            DisplayStatus = STATUS_NA
        Case Else
            DisplayStatus = Replace$(statusText, "NotCalculated", STATUS_NOT_CALCULATED_SHORT, 1, -1, vbTextCompare)
    End Select
End Function

' True для расчетов, которые дали физически интерпретируемый результат:
' проверка либо прошла, либо показала недостаточную несущую способность.
Public Function IsFinished(ByVal statusText As String) As Boolean
    statusText = ToUserStatus(statusText)
    IsFinished = (statusText = STATUS_OK Or statusText = STATUS_FAIL)
End Function

' True для OverallStatus, при котором строку LC нужно выделить как проблемную.
' N/A сюда не входит: это нейтральный пропуск неприменимой проверки.
Public Function IsFailedOverall(ByVal statusText As String) As Boolean
    statusText = ToUserStatus(statusText)
    IsFailedOverall = (statusText = STATUS_FAIL Or statusText = STATUS_NUMFAIL Or statusText = STATUS_INPUTERR)
End Function

' Распознает статусы, означающие именно численную несходимость.
Public Function IsNumerical(ByVal statusText As String) As Boolean
    IsNumerical = (InStr(1, statusText, STATUS_NUMFAIL, vbTextCompare) > 0 Or _
        InStr(1, statusText, "NumericalFailure", vbTextCompare) > 0 Or _
        InStr(1, statusText, "SingularTangent", vbTextCompare) > 0)
End Function

' Распознает статусы ошибок исходных данных.
Public Function IsInvalid(ByVal statusText As String) As Boolean
    IsInvalid = (InStr(1, statusText, STATUS_INPUTERR, vbTextCompare) > 0 Or _
        InStr(1, statusText, "InputError", vbTextCompare) > 0 Or _
        InStr(1, statusText, "InvalidInput", vbTextCompare) > 0)
End Function

' Переводит результат capacity solver-а в короткий пользовательский статус.
' Lambda меньше единицы считается FAIL, потому что равновесие найдено, но
' заданный уровень нагрузки выше рассчитанной несущей способности.
Public Function CapacityStatusFromSolver(ByVal cap As CCapacitySolver) As String
    If cap Is Nothing Then
        CapacityStatusFromSolver = STATUS_NUMFAIL
    ElseIf Not cap.Converged Then
        If IsInvalid(cap.StopReason) Or IsInvalid(cap.LimitState) Then
            CapacityStatusFromSolver = STATUS_INPUTERR
        Else
            CapacityStatusFromSolver = STATUS_NUMFAIL
        End If
    ElseIf cap.LambdaUltimate > 0# And cap.LambdaUltimate < 1# Then
        CapacityStatusFromSolver = STATUS_FAIL
    Else
        CapacityStatusFromSolver = STATUS_OK
    End If
End Function

' Переводит результат расчета трещин в тот же пользовательский словарь.
' Для трещин FAIL означает, что ширина раскрытия превышает пользовательский
' a_crc,ult при корректно найденном НДС.
Public Function CrackStatusFromCalculator(ByVal crack As CCrackWidthCalculator) As String
    If crack Is Nothing Then
        CrackStatusFromCalculator = STATUS_NUMFAIL
    ElseIf crack.Converged Then
        If crack.CrackFormed And crack.AllowableCrackWidth > 0# And _
                crack.CrackWidth > crack.AllowableCrackWidth Then
            CrackStatusFromCalculator = STATUS_FAIL
        Else
            CrackStatusFromCalculator = STATUS_OK
        End If
    ElseIf IsInvalid(crack.StopReason) Then
        CrackStatusFromCalculator = STATUS_INPUTERR
    Else
        CrackStatusFromCalculator = STATUS_NUMFAIL
    End If
End Function

' Объединяет два статуса по утвержденному приоритету:
' InputErr -> NumFail -> FAIL -> OK, при этом N/A пропускается.
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
        Case STATUS_INPUTERR
            StatusPriority = 4
        Case STATUS_NUMFAIL
            StatusPriority = 3
        Case STATUS_FAIL
            StatusPriority = 2
        Case STATUS_OK
            StatusPriority = 1
        Case Else
            StatusPriority = 0
    End Select
End Function
