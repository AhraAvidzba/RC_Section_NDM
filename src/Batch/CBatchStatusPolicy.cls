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
' сводится только к OK, FAIL, BaseFail, NumFail, InputErr, CalcErr или N/A. Класс не запускает
' расчет и не знает о сочетаниях нагрузок: его граница ответственности -
' только перевод и агрегирование статусов.

Private Const STATUS_OK As String = "OK" ' Расчетная ветка завершилась без ошибки.
Private Const STATUS_FAIL As String = "FAIL" ' Равновесие найдено, но физический предел или проверка не проходят.
Private Const STATUS_BASEFAIL As String = "BaseFail" ' Capacity: базовая точка траектории уже за предельным состоянием.
Private Const STATUS_NUMFAIL As String = "NumFail" ' Solver не нашел равновесие в заданных численных ограничениях.
Private Const STATUS_INPUTERR As String = "InputErr" ' Ошибка исходных данных, а не численная несходимость.
Private Const STATUS_CALCERR As String = "CalcErr" ' Внутренняя ошибка программы или несформированный result.
Private Const STATUS_NA As String = "N/A" ' Расчетная проверка не относится к этому LC/режиму.
Private Const STATUS_NOT_CALCULATED_SHORT As String = "NotCalcul." ' Короткая подпись для служебных нерассчитанных полей.
Private Const STRESS_ZERO_TOLERANCE As Double = 0.000000001 ' Малые численные шумы напряжений не должны запускать проверку сжатого бетона.

Private mResultPolicy As CResultStatusPolicy ' Новый типизированный слой status policy; этот класс остается совместимым фасадом.

' Создает внутреннюю policy, которая знает про InternalStatus/ResultCode.
Private Sub Class_Initialize()
    Set mResultPolicy = New CResultStatusPolicy
End Sub

' Возвращает статус успешной расчетной ветки.
Public Property Get OK() As String
    OK = STATUS_OK
End Property

' Возвращает статус физически непрошедшей проверки при найденном равновесии.
Public Property Get Fail() As String
    Fail = STATUS_FAIL
End Property

' Возвращает специальный статус capacity, когда базовая точка траектории уже за пределом.
Public Property Get BaseFail() As String
    BaseFail = STATUS_BASEFAIL
End Property

' Возвращает статус настоящей численной несходимости.
Public Property Get NumFail() As String
    NumFail = STATUS_NUMFAIL
End Property

' Возвращает статус ошибки пользовательских исходных данных.
Public Property Get InputErr() As String
    InputErr = STATUS_INPUTERR
End Property

' Возвращает статус внутренней ошибки программы.
Public Property Get CalcErr() As String
    CalcErr = STATUS_CALCERR
End Property

' Возвращает статус неприменимой или не запущенной проверки.
Public Property Get NotApplicable() As String
    NotApplicable = STATUS_NA
End Property

' Собирает OverallStatus из статусов применимых расчетных разделов.
' N/A не ухудшает итог: это означает, что ветка не относится к LC или режиму.
Public Function Aggregate(ByVal directStateStatus As String, ByVal capacityStatus As String, _
        ByVal crackStatus As String, Optional ByVal longitudinalCrackStatus As String = vbNullString, _
        Optional ByVal stabilityStatus As String = vbNullString) As String
    Aggregate = Policy.AggregateText(directStateStatus, capacityStatus, crackStatus, _
        longitudinalCrackStatus, stabilityStatus)
End Function

' Собирает OverallStatus из типизированных ResultMeta. Это основной путь для
' новой архитектуры, где writer-ы не должны парсить StopReason.
Public Function AggregateMeta(ByVal directMeta As CResultMeta, ByVal capacityMeta As CResultMeta, _
        ByVal crackMeta As CResultMeta, Optional ByVal longitudinalMeta As CResultMeta = Nothing, _
        Optional ByVal stabilityMeta As CResultMeta = Nothing) As String
    AggregateMeta = Policy.AggregateMeta(directMeta, capacityMeta, crackMeta, longitudinalMeta, stabilityMeta)
End Function

' Переводит внутренний текст расчетной ветки в один из пяти пользовательских
' статусов. Этот метод используется там, где нужна именно оценка результата:
' агрегирование OverallStatus, подсветка строк и проверка завершенности ветки.
Public Function ToUserStatus(ByVal statusText As String) As String
    ToUserStatus = Policy.ToUserStatus(statusText)
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
        Case "BASEFAIL"
            DisplayStatus = STATUS_BASEFAIL
        Case "NUMFAIL", "NUMERICALFAILURE", "SINGULARTANGENT"
            DisplayStatus = STATUS_NUMFAIL
        Case "INPUTERR", "INPUTERROR", "INVALIDINPUT"
            DisplayStatus = STATUS_INPUTERR
        Case "CALCERR", "INTERNALERROR"
            DisplayStatus = STATUS_CALCERR
        Case "N/A", "NA", "NOTCALCULATED", "NOTCALCUL.", "SKIPPEDZEROMOMENT", "SKIPPEDZEROBASEMOMENT", ""
            DisplayStatus = STATUS_NA
        Case Else
            DisplayStatus = Replace$(statusText, "NotCalculated", STATUS_NOT_CALCULATED_SHORT, 1, -1, vbTextCompare)
    End Select
End Function

' True для расчетов, которые дали физически интерпретируемый результат:
' проверка либо прошла, либо показала недостаточную несущую способность.
Public Function IsFinished(ByVal statusText As String) As Boolean
    IsFinished = Policy.IsFinished(statusText)
End Function

' True для OverallStatus, при котором строку LC нужно выделить как проблемную.
' N/A сюда не входит: это нейтральный пропуск неприменимой проверки.
Public Function IsFailedOverall(ByVal statusText As String) As Boolean
    IsFailedOverall = Policy.IsFailedOverall(statusText)
End Function

' Распознает статусы, означающие именно численную несходимость.
Public Function IsNumerical(ByVal statusText As String) As Boolean
    IsNumerical = Policy.IsNumerical(statusText)
End Function

' Переводит типизированный InternalStatus/ResultCode в внешний статус Results.
Public Function StatusFromMeta(ByVal meta As CResultMeta) As String
    StatusFromMeta = Policy.ExternalStatus(meta)
End Function

' Создает meta из уже принятого внешнего статуса для переходных мест, где
' строковый статус пока формируется раньше result-объекта.
Public Function MetaFromStatus(ByVal statusText As String, ByVal resultKind As EResultKind, _
        Optional ByVal resultComment As String = vbNullString) As CResultMeta
    Set MetaFromStatus = Policy.MetaFromLegacyStatus(statusText, resultKind, resultComment)
End Function

' Формирует meta для этапа, который не запускался из-за отсутствия обязательного
' предыдущего результата. Внешний N/A появится позже только через CResultStatusPolicy.
Public Function BlockedByDependencyMeta(ByVal resultKind As EResultKind, _
        ByVal dependencyName As String, Optional ByVal reason As String = vbNullString) As CResultMeta
    Dim comment As String
    comment = "Расчет не выполнен: нет обязательного результата " & dependencyName & "."
    If Len(Trim$(reason)) > 0 Then comment = comment & " " & reason

    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsBlockedByDependency, rcBlockedByDependency, resultKind, comment
    Set BlockedByDependencyMeta = meta
End Function

' Распознает статусы ошибок исходных данных.
Public Function IsInvalid(ByVal statusText As String) As Boolean
    IsInvalid = Policy.IsInvalid(statusText)
End Function

' Переводит результат расчета трещин в тот же пользовательский словарь.
' Для трещин FAIL означает, что ширина раскрытия превышает пользовательский
' a_crc,ult при корректно найденном НДС.
Public Function CrackStatusFromCalculator(ByVal crack As CCrackWidthCalculator) As String
    CrackStatusFromCalculator = Policy.ExternalStatus(CrackMetaFromCalculator(crack))
End Function

' Формирует типизированный результат этапа образования нормальной трещины.
' CRITERION_NOT_REACHED считается корректным физическим исходом search: трещина
' не образовалась, но численной ошибки при этом нет.
Public Function CrackFormationMetaFromCalculator(ByVal crack As CCrackWidthCalculator) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If crack Is Nothing Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackFormation, "Crack calculator не создан."
    ElseIf crack.Converged And crack.CrackFormed Then
        meta.SetResult rsSuccess, rcCheckPassed, rkCrackFormation, crack.StopReason
    ElseIf crack.Converged Then
        meta.SetResult rsSuccess, rcCriterionNotReached, rkCrackFormation, crack.StopReason
    ElseIf IsInvalid(crack.StopReason) Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackFormation, crack.StopReason
    ElseIf IsNumerical(crack.StopReason) Then
        meta.SetResult rsNumericalFailure, rcNumericalFailure, rkCrackFormation, crack.StopReason
    Else
        meta.SetResult rsInternalError, rcInternalError, rkCrackFormation, crack.StopReason
    End If

    Set CrackFormationMetaFromCalculator = meta
End Function

' Формирует типизированный результат расчета нормальных трещин.
' Численная несходимость допускается только как отражение уже выполненного
' state/search-поиска внутри калькулятора; чистые формулы ширины сами NumFail не создают.
Public Function CrackMetaFromCalculator(ByVal crack As CCrackWidthCalculator) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If crack Is Nothing Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackWidth, "Crack calculator не создан."
    ElseIf crack.Converged And Not crack.CrackFormed Then
        meta.SetResult rsNotApplicable, rcCrackNotFormed, rkCrackWidth, crack.StopReason, _
            vbNullString, False, True
    ElseIf crack.Converged Then
        If crack.CrackFormed And crack.AllowableCrackWidth > 0# And _
                crack.CrackWidth > crack.AllowableCrackWidth Then
            meta.SetResult rsCheckFailed, rcCrackWidthExceeded, rkCrackWidth, crack.StopReason
        Else
            meta.SetResult rsSuccess, rcCheckPassed, rkCrackWidth, crack.StopReason
        End If
    ElseIf IsInvalid(crack.StopReason) Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackWidth, crack.StopReason
    ElseIf IsNumerical(crack.StopReason) Then
        meta.SetResult rsNumericalFailure, rcNumericalFailure, rkCrackWidth, crack.StopReason
    Else
        meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackWidth, crack.StopReason
    End If

    Set CrackMetaFromCalculator = meta
End Function

' Формирует meta формульного этапа из уже подготовленного result-объекта.
' Метод нужен, чтобы writer-ы и batch не читали внутренности calculator-а после
' того, как формула раскрытия вынесена в отдельный слой.
Public Function CrackMetaFromResult(ByVal widthResult As CCrackWidthResult) As CResultMeta
    If widthResult Is Nothing Then
        Dim meta As CResultMeta
        Set meta = New CResultMeta
        meta.SetResult rsBlockedByDependency, rcBlockedByDependency, rkCrackWidth, _
            "Результат ширины нормальной трещины не сформирован."
        Set CrackMetaFromResult = meta
    Else
        Set CrackMetaFromResult = widthResult.ResultMeta
    End If
End Function

' Возвращает статус проверки продольных трещин по сжатию бетона.
' Если сжатого бетона в найденном НДС нет, проверка неприменима и не ухудшает OverallStatus.
Public Function LongitudinalCrackStatus(ByVal maxCompressionStress As Double, ByVal rbMc2 As Double) As String
    LongitudinalCrackStatus = Policy.ExternalStatus(LongitudinalCrackMeta(maxCompressionStress, rbMc2))
End Function

' Формирует типизированный результат проверки продольных трещин.
' Эта проверка не решает равновесие и поэтому не может сама породить NumFail.
Public Function LongitudinalCrackMeta(ByVal maxCompressionStress As Double, ByVal rbMc2 As Double) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If maxCompressionStress <= STRESS_ZERO_TOLERANCE Then
        meta.SetNotApplicable rkLongitudinalCrack, "Сжатого бетона для проверки продольных трещин нет."
    ElseIf rbMc2 <= 0# Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkLongitudinalCrack, "Rb,mc2 должен быть больше нуля."
    ElseIf maxCompressionStress <= rbMc2 Then
        meta.SetResult rsSuccess, rcCheckPassed, rkLongitudinalCrack, vbNullString
    Else
        meta.SetResult rsCheckFailed, rcLongitudinalCrackExceeded, rkLongitudinalCrack, vbNullString
    End If

    Set LongitudinalCrackMeta = meta
End Function

' Возвращает пользовательский статус из результата фильтра устойчивости.
' Сам CStabilityCalculator хранит подробные расчетные причины, а наружу
' отдается только общий словарь Summary.
Public Function StabilityStatusFromCalculator(ByVal stability As CStabilityCalculator) As String
    StabilityStatusFromCalculator = Policy.ExternalStatus(StabilityMetaFromCalculator(stability))
End Function

' Формирует типизированный результат устойчивости.
' Фильтр устойчивости является инженерской проверкой по найденным усилиям и
' геометрическим характеристикам; численную несходимость он не создает.
Public Function StabilityMetaFromCalculator(ByVal stability As CStabilityCalculator) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta

    If stability Is Nothing Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkStability, "Stability calculator не создан."
    ElseIf Not stability.InputValid Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkStability, stability.StopReason
    ElseIf Not stability.Applicable Then
        meta.SetNotApplicable rkStability, stability.StopReason
    ElseIf stability.CheckPassed Then
        meta.SetResult rsSuccess, rcCheckPassed, rkStability, stability.StopReason
    Else
        meta.SetResult rsCheckFailed, rcCheckFailed, rkStability, stability.StopReason
    End If

    Set StabilityMetaFromCalculator = meta
End Function

' Возвращает новый типизированный слой policy и лениво создает его, если VBA
' восстановил объект без вызова Class_Initialize.
Private Function Policy() As CResultStatusPolicy
    If mResultPolicy Is Nothing Then Set mResultPolicy = New CResultStatusPolicy
    Set Policy = mResultPolicy
End Function
