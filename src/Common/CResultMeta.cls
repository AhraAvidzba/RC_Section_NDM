VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResultMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' МЕТАДАННЫЕ РАСЧЕТНОГО РЕЗУЛЬТАТА
' ==========================================================================
' CResultMeta хранит машинную оценку результата: внутренний статус, код
' причины, пользовательский комментарий и техническую диагностику. Объект не
' решает, как это будет выглядеть на листе Results: внешний статус OK/FAIL/
' NumFail/N/A формирует CResultStatusPolicy.
'
' ResultMeta является постоянным машинным контрактом между расчетными слоями,
' batch-orchestration и writer-ами. Внешний текстовый статус получается только
' через CResultStatusPolicy.

Private mInternalStatus As EResultInternalStatus ' Внутренний статус, на котором работает расчетная логика.
Private mResultCode As EResultCode ' Машинный код точной причины результата.
Private mResultKind As EResultKind ' Контекст результата: direct state, capacity, cracks, stability.
Private mResultComment As String ' Короткое русское пояснение для пользователя и report.
Private mDiagnosticDetails As String ' Технические подробности solver/search, если они нужны.
Private mApplies As Boolean ' Ветка относится к LC; blocked сохраняет True, хотя нужных данных нет.
Private mCalculated As Boolean ' Выполнялся собственный solve/search/формульный расчет, не только проверка применимости.

' Инициализирует объект нейтральным "не сформировано" состоянием.
Private Sub Class_Initialize()
    Clear
End Sub

' Сбрасывает meta к безопасному начальному состоянию.
Public Sub Clear()
    mInternalStatus = rsUnset
    mResultCode = rcNone
    mResultKind = rkGeneric
    mResultComment = vbNullString
    mDiagnosticDetails = vbNullString
    mApplies = True
    mCalculated = False
End Sub

' Заполняет поля результата. Не запрошенная, неприменимая, заблокированная
' ветка и ошибка валидации не объявляют собственный расчет выполненным.
' Для внутренней ошибки после запуска вызывающий слой передает фактический
' calculated; для отсутствующего результата он обязан передать False.
Public Sub SetResult(ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal resultKind As EResultKind, _
        Optional ByVal resultComment As String = vbNullString, _
        Optional ByVal diagnosticDetails As String = vbNullString, _
        Optional ByVal applies As Boolean = True, Optional ByVal calculated As Boolean = True)
    mInternalStatus = internalStatus
    mResultCode = resultCode
    mResultKind = resultKind
    mResultComment = resultComment
    mDiagnosticDetails = diagnosticDetails
    mApplies = applies
    mCalculated = calculated
    Select Case internalStatus
        Case rsNotRequested, rsNotApplicable
            mApplies = False
            mCalculated = False
        Case rsBlockedByDependency, rsInvalidInput, rsInvalidConfiguration
            mCalculated = False
    End Select
End Sub

' Заполняет meta как отключенный профилем расчет.
Public Sub SetNotRequested(ByVal resultKind As EResultKind, Optional ByVal resultComment As String = vbNullString)
    SetResult rsNotRequested, rcNotRequested, resultKind, resultComment, vbNullString, False, False
End Sub

' Заполняет meta как неприменимую расчетную ветку.
Public Sub SetNotApplicable(ByVal resultKind As EResultKind, Optional ByVal resultComment As String = vbNullString)
    SetResult rsNotApplicable, rcNotApplicable, resultKind, resultComment, vbNullString, False, False
End Sub

' Отмечает применимый этап, который не запускался из-за отсутствующего
' обязательного результата. Причина зависимости дополняет сообщение этапа.
Public Sub SetBlockedByDependency(ByVal resultKind As EResultKind, _
        ByVal dependencyName As String, Optional ByVal reason As String = vbNullString)
    Dim comment As String
    comment = "Расчет не выполнен: нет обязательного результата " & dependencyName & "."
    If Len(Trim$(reason)) > 0 Then comment = comment & " " & Trim$(reason)
    SetResult rsBlockedByDependency, rcBlockedByDependency, resultKind, comment, vbNullString, True, False
End Sub

' Сохраняет машинную причину неуспешного solve без разбора текста сообщения.
' Вызывающий слой передает факт численной попытки отдельно от причины отказа;
' отсутствие обязательного solver-а является внутренней ошибкой контракта.
Public Sub SetSolverFailure(ByVal failureCode As ESolverFailureCode, _
        ByVal resultKind As EResultKind, ByVal reason As String, ByVal attempted As Boolean)
    Dim internalStatus As EResultInternalStatus
    Dim resultCode As EResultCode
    Select Case failureCode
        Case sfcInvalidInput
            internalStatus = rsInvalidInput: resultCode = rcInvalidInput
        Case sfcInvalidConfiguration
            internalStatus = rsInvalidConfiguration: resultCode = rcInvalidConfiguration
        Case sfcSingularTangent
            internalStatus = rsNumericalFailure: resultCode = rcSingularTangent
        Case sfcNumericalFailure
            internalStatus = rsNumericalFailure: resultCode = rcNumericalFailure
        Case Else
            internalStatus = rsInternalError: resultCode = rcInternalError
            If Len(Trim$(reason)) = 0 Then reason = "Решатель не сформировал результат равновесия."
    End Select
    SetResult internalStatus, resultCode, resultKind, reason, vbNullString, True, attempted
End Sub

' Создает независимую копию meta для хранения внутри result-объекта.
Public Function Clone() As CResultMeta
    Dim result As CResultMeta
    Set result = New CResultMeta
    result.SetResult mInternalStatus, mResultCode, mResultKind, mResultComment, _
        mDiagnosticDetails, mApplies, mCalculated
    Set Clone = result
End Function

Public Property Get InternalStatus() As EResultInternalStatus
    InternalStatus = mInternalStatus
End Property

Public Property Let InternalStatus(ByVal value As EResultInternalStatus)
    mInternalStatus = value
End Property

Public Property Get ResultCode() As EResultCode
    ResultCode = mResultCode
End Property

Public Property Let ResultCode(ByVal value As EResultCode)
    mResultCode = value
End Property

Public Property Get ResultKind() As EResultKind
    ResultKind = mResultKind
End Property

Public Property Let ResultKind(ByVal value As EResultKind)
    mResultKind = value
End Property

Public Property Get ResultComment() As String
    ResultComment = mResultComment
End Property

Public Property Let ResultComment(ByVal value As String)
    mResultComment = value
End Property

Public Property Get DiagnosticDetails() As String
    DiagnosticDetails = mDiagnosticDetails
End Property

Public Property Let DiagnosticDetails(ByVal value As String)
    mDiagnosticDetails = value
End Property

Public Property Get Applies() As Boolean
    Applies = mApplies
End Property

Public Property Let Applies(ByVal value As Boolean)
    mApplies = value
End Property

Public Property Get Calculated() As Boolean
    Calculated = mCalculated
End Property

Public Property Let Calculated(ByVal value As Boolean)
    mCalculated = value
End Property

' Возвращает компактную строку для diagnostics без влияния на пользовательский статус.
Public Property Get DiagnosticText() As String
    DiagnosticText = ResultKindToText(mResultKind) & ": " & _
        ResultInternalStatusToText(mInternalStatus) & " / " & ResultCodeToText(mResultCode)
    If Len(Trim$(mResultComment)) > 0 Then DiagnosticText = DiagnosticText & "; " & mResultComment
End Property
