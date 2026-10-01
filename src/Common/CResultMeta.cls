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
    comment = "Расчет не выполнен: нет допустимого обязательного результата " & dependencyName & "."
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

' Добавляет собственную причину в комментарий родительского result без дублей.
' Обычные успешные и неприменимые ветви не загромождают сводку; blocked и
' warning сохраняются. Для уже собранного subtree передается prepared=True:
' его текст не размечается повторно и не теряется при итоговом внешнем OK.
Public Sub AppendCommentTo(ByRef targetText As String, Optional ByVal prepared As Boolean = False)
    If Not prepared Then
        Select Case mInternalStatus
            Case rsCheckFailed, rsNumericalFailure, rsInvalidInput, rsInvalidConfiguration, _
                    rsInternalError, rsBlockedByDependency, rsSuccessWithWarning
            Case Else
                Exit Sub
        End Select
    End If
    Dim commentText As String
    commentText = Trim$(mResultComment)
    If Len(commentText) = 0 Then Exit Sub
    If Not prepared Then commentText = ResultKindLabel() & ": " & commentText
    If InStr(1, "; " & targetText & "; ", "; " & commentText & "; ", vbTextCompare) > 0 Then Exit Sub
    If Len(targetText) > 0 Then targetText = targetText & "; "
    targetText = targetText & commentText
End Sub

' Дает короткую инженерную метку владельца причины при сборке result-subtree.
' Это подпись готовых данных, а не правило выбора расчетной ветви или статуса.
Private Function ResultKindLabel() As String
    Select Case mResultKind
        Case rkDirectState: ResultKindLabel = "НДС"
        Case rkCapacity: ResultKindLabel = "Несущая"
        Case rkCrackFormation: ResultKindLabel = "Трещинообразование"
        Case rkCrackWidth: ResultKindLabel = "Ширина трещин"
        Case rkLongitudinalCrack: ResultKindLabel = "Продольные трещины"
        Case rkStability: ResultKindLabel = "Устойчивость"
        Case Else: ResultKindLabel = "Расчет"
    End Select
End Function

Public Property Get InternalStatus() As EResultInternalStatus
    InternalStatus = mInternalStatus
End Property

Public Property Get ResultCode() As EResultCode
    ResultCode = mResultCode
End Property

Public Property Get ResultKind() As EResultKind
    ResultKind = mResultKind
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

Public Property Get Calculated() As Boolean
    Calculated = mCalculated
End Property

' Возвращает компактную строку для diagnostics без влияния на пользовательский статус.
Public Property Get DiagnosticText() As String
    DiagnosticText = ResultKindToText(mResultKind) & ": " & _
        ResultInternalStatusToText(mInternalStatus) & " / " & ResultCodeToText(mResultCode)
    If Len(Trim$(mResultComment)) > 0 Then DiagnosticText = DiagnosticText & "; " & mResultComment
End Property
