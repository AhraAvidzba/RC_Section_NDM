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
' Единые правила метаданных и внешних статусов результатов
' ==========================================================================
' Преобразует InternalStatus, ResultCode и ResultKind в обозначение статуса
' для пользователя. Применяет общий словарь вывода к готовому результату;
' численные расчёты и чтение Excel выполняются другими классами.

Private Const STATUS_OK As String = "OK" ' Проверка выполнена и прошла.
Private Const STATUS_FAIL As String = "FAIL" ' Результат получен, но инженерная проверка не проходит.
Private Const STATUS_BASEFAIL As String = "BaseFail" ' Постоянная часть пути не проходит критерий либо фиксированный путь физически его не достигает.
Private Const STATUS_NUMFAIL As String = "NumFail" ' Равновесие или предельная точка не найдены численным методом.
Private Const STATUS_INPUTERR As String = "InputErr" ' Ошибка пользовательских данных или Config.
Private Const STATUS_CALCERR As String = "CalcErr" ' Внутренняя ошибка программы или отсутствие обязательного результата.
Private Const STATUS_NA As String = "N/A" ' Расчет не запрошен, неприменим или заблокирован зависимостью.

' Код успешной применимой проверки по общей политике статусов.
Public Property Get OK() As String
    OK = STATUS_OK
End Property

' Код физического невыполнения применимой проверки.
Public Property Get Fail() As String
    Fail = STATUS_FAIL
End Property

' Код непригодности базового состояния для зависимой проверки.
Public Property Get BaseFail() As String
    BaseFail = STATUS_BASEFAIL
End Property

' Код численной неудачи; не означает доказанное физическое разрушение.
Public Property Get NumFail() As String
    NumFail = STATUS_NUMFAIL
End Property

' Код ошибки исходных данных, требующей исправления ввода.
Public Property Get InputErr() As String
    InputErr = STATUS_INPUTERR
End Property

' Код вычислительной ошибки с сохранением её причины.
Public Property Get CalcErr() As String
    CalcErr = STATUS_CALCERR
End Property

' Код неприменимой или не запрошенной проверки.
Public Property Get NotApplicable() As String
    NotApplicable = STATUS_NA
End Property

' Возвращает обозначение статуса для пользователя по сохранённым коду и признакам результата.
Public Function ExternalStatus(ByVal meta As CResultMeta) As String
    If meta Is Nothing Then
        ExternalStatus = STATUS_CALCERR
        Exit Function
    End If

    If IsBaseFailure(meta) Then
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


' Определяет особый физический BaseFail по виду проверки и машинной причине.
' Пользовательский текст и признаки локализации не участвуют в решении.
Private Function IsBaseFailure(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    IsBaseFailure = (meta.ResultKind = rkCapacity Or meta.ResultKind = rkCrackFormation) And _
        meta.InternalStatus = rsCheckFailed And _
        (meta.ResultCode = rcInitialStateBeyondLimit Or meta.ResultCode = rcCriterionNotReached)
End Function

' Возвращает прежний приоритет непосредственно по метаданным. Отсутствующий,
' не сформированный и неизвестный исход сохраняет приоритет внутренней ошибки.
Public Function MetaPriority(ByVal meta As CResultMeta) As Long
    MetaPriority = 6
    If meta Is Nothing Then Exit Function
    Select Case meta.InternalStatus
        Case rsNotRequested, rsNotApplicable, rsBlockedByDependency
            MetaPriority = 0
        Case rsSuccess, rsSuccessWithWarning
            MetaPriority = 1
        Case rsCheckFailed
            If IsBaseFailure(meta) Then MetaPriority = 3 Else MetaPriority = 2
        Case rsNumericalFailure
            MetaPriority = 4
        Case rsInvalidInput, rsInvalidConfiguration
            MetaPriority = 5
    End Select
End Function

' Определяет физически интерпретируемое завершение без преобразования в строку.
' Признаки Applies/Calculated остаются исходными и здесь не переписываются.
Public Function IsMetaFinished(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    Select Case meta.InternalStatus
        Case rsSuccess, rsSuccessWithWarning, rsCheckFailed
            IsMetaFinished = True
    End Select
End Function

' Распознаёт ошибку ввода/Config по внутреннему статусу без внешнего кода.
Public Function IsMetaInvalid(ByVal meta As CResultMeta) As Boolean
    If meta Is Nothing Then Exit Function
    IsMetaInvalid = (meta.InternalStatus = rsInvalidInput Or meta.InternalStatus = rsInvalidConfiguration)
End Function

' Определяет проблемный итог по тому же приоритету, что строковый внешний API.
' Физический FAIL, численный отказ и ошибка данных сохраняют разные метаданные.
Public Function IsMetaFailedOverall(ByVal meta As CResultMeta) As Boolean
    IsMetaFailedOverall = (MetaPriority(meta) > 1)
End Function

' Отделяет обязательное внимание к неподтверждённому расчёту или ошибке
' от обычного физического невыполнения. NumFail остаётся численным отказом.
Public Function MetaRequiresAttention(ByVal meta As CResultMeta) As Boolean
    MetaRequiresAttention = (MetaPriority(meta) >= 4)
End Function

' Проверяет уже заданное внешнее обозначение статуса. Неизвестная строка
' считается CalcErr. Причина результата должна быть передана отдельными
' InternalStatus и ResultCode, а не определяться из текста комментария.
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

' Выбирает результат с наиболее серьёзным статусом по общему приоритету.
' Сохраняет его точную причину, диагностику и готовое пояснение
' в независимой копии для дальнейшего использования.
Public Function WorstResultMeta(ByVal fallbackKind As EResultKind, _
        ParamArray metas() As Variant) As CResultMeta
    Dim selected As CResultMeta, candidate As CResultMeta
    Dim i As Long, sourceKeys As String, key As Variant
    For i = LBound(metas) To UBound(metas)
        Set candidate = Nothing
        If IsObject(metas(i)) Then Set candidate = metas(i)
        If Not candidate Is Nothing Then
            If Len(candidate.SourceKeys) > 0 Then
                For Each key In Split(candidate.SourceKeys, ";")
                    If Len(Trim$(CStr(key))) > 0 Then
                        If InStr(1, ";" & sourceKeys & ";", ";" & Trim$(CStr(key)) & ";", vbTextCompare) = 0 Then
                            If Len(sourceKeys) > 0 Then sourceKeys = sourceKeys & ";"
                            sourceKeys = sourceKeys & Trim$(CStr(key))
                        End If
                    End If
                Next key
            End If
            If selected Is Nothing Then
                Set selected = candidate
            ElseIf MetaPriority(candidate) > MetaPriority(selected) Then
                Set selected = candidate
            ElseIf candidate.InternalStatus = rsSuccessWithWarning And selected.InternalStatus = rsSuccess Then
                Set selected = candidate ' Оба исхода отображаются OK, но значимое предупреждение и его машинный код не должны теряться из-за порядка дочерних ветвей.
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
    result.SourceKeys = sourceKeys
    Set WorstResultMeta = result
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

' Распознаёт предусмотренные внешние обозначения численной несходимости.
' Текст причины решателя не анализируется, поэтому формулировка
' и язык пояснения на классификацию не влияют.
Public Function IsNumerical(ByVal statusText As String) As Boolean
    IsNumerical = (ToUserStatus(statusText) = STATUS_NUMFAIL)
End Function

' Распознает только утвержденные внешние статусы ошибки ввода/Config.
' Подстроки StopReason вроде InputError/InvalidConfiguration не являются
' источником статуса; для этого используется CResultMeta.
Public Function IsInvalid(ByVal statusText As String) As Boolean
    IsInvalid = (ToUserStatus(statusText) = STATUS_INPUTERR)
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
