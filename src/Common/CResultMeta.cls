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
' Статус результата и пояснение причины
' ==========================================================================
' Хранит внутренний статус, код причины, пользовательское пояснение
' и техническую диагностику. Внешние обозначения OK/FAIL/NumFail/N/A
' определяет CResultStatusPolicy по этим данным, без разбора текста ошибки.

Private mInternalStatus As EResultInternalStatus ' Внутренний статус, на котором работает расчетная логика.
Private mResultCode As EResultCode ' Машинный код точной причины результата.
Private mResultKind As EResultKind ' Вид проверки: НДС при заданной нагрузке, несущая способность, трещины или устойчивость.
Private mResultComment As String ' Короткое русское пояснение для пользователя и report.
Private mDiagnosticDetails As String ' Подробная диагностика решения равновесия или поиска предела, если она сохранена.
Private mSourceKeys As String ' Ключи исходных параметров через точку с запятой; адреса сопоставляются при выводе, без обращения ядра к Excel.
Private mApplies As Boolean ' Ветка относится к LC; blocked сохраняет True, хотя нужных данных нет.
Private mCalculated As Boolean ' Показывает, выполнялся ли сам расчёт, а не только проверка его применимости.

' Инициализирует объект нейтральным "не сформировано" состоянием.
Private Sub Class_Initialize()
    Clear
End Sub

' Очищает статус, пояснение и признаки выполнения перед новым заполнением.
Public Sub Clear()
    mInternalStatus = rsUnset
    mResultCode = rcNone
    mResultKind = rkGeneric
    mResultComment = vbNullString
    mDiagnosticDetails = vbNullString
    mSourceKeys = vbNullString
    mApplies = True
    mCalculated = False
End Sub

' Сохраняет статус, причину и признаки выполнения проверки. Незапрошенная,
' неприменимая или заблокированная проверка не считается выполненной.
' При внутренней ошибке вызывающий код передаёт фактический признак calculated;
' для отсутствующего результата он должен быть False.
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
    mSourceKeys = vbNullString
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

' Отмечает проверку как отключённую выбранным профилем.
Public Sub SetNotRequested(ByVal resultKind As EResultKind, Optional ByVal resultComment As String = vbNullString)
    SetResult rsNotRequested, rcNotRequested, resultKind, resultComment, vbNullString, False, False
End Sub

' Отмечает проверку как неприменимую к данному расчётному случаю.
Public Sub SetNotApplicable(ByVal resultKind As EResultKind, Optional ByVal resultComment As String = vbNullString)
    SetResult rsNotApplicable, rcNotApplicable, resultKind, resultComment, vbNullString, False, False
End Sub

' Отмечает применимый этап, который не запускался из-за отсутствующего
' обязательного результата. Причина зависимости дополняет сообщение этапа.
Public Sub SetBlockedByDependency(ByVal resultKind As EResultKind, _
        ByVal dependencyName As String, Optional ByVal reason As String = vbNullString)
    Dim comment As String
    comment = "Расчет не выполнен: недоступны обязательные данные — " & DependencyCaption(dependencyName) & "."
    If Len(Trim$(reason)) > 0 Then comment = comment & " " & Trim$(reason)
    SetResult rsBlockedByDependency, rcBlockedByDependency, resultKind, comment, vbNullString, True, False
End Sub

' Переводит названия обязательных данных в понятное пояснение; выбор зависимостей
' и решение о запуске проверки остаются у вызывающего расчётного класса.
Private Function DependencyCaption(ByVal dependencyName As String) As String
    Select Case dependencyName
        Case "CurrentCrackedState": DependencyCaption = "текущее НДС с трещинами от заданного сочетания"
        Case "исходных данных сочетания": DependencyCaption = "исходные данные сочетания"
        Case "текущего НДС от сочетания": DependencyCaption = "текущее НДС от сочетания"
        Case Else: DependencyCaption = dependencyName
    End Select
End Function

' Сохраняет код причины неудачного решения без разбора текста сообщения.
' Факт запуска численных итераций передаётся отдельно. Отсутствие обязательного
' объекта решателя считается внутренней ошибкой программы.
Public Sub SetSolverFailure(ByVal failureCode As ESolverFailureCode, _
        ByVal resultKind As EResultKind, ByVal reason As String, ByVal attempted As Boolean, _
        Optional ByVal sourceKeys As String = vbNullString)
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
    mSourceKeys = sourceKeys
End Sub

' Хранит структурированную связь с настройками; ключи не извлекаются из русского пояснения.
Public Property Get SourceKeys() As String
    SourceKeys = mSourceKeys
End Property

' Получает ключи от класса, проверившего условие, отдельно от пользовательского текста.
Public Property Let SourceKeys(ByVal value As String)
    mSourceKeys = value
End Property

' Создаёт независимую копию статуса и пояснения для хранения в объекте результата.
Public Function Clone() As CResultMeta
    Dim result As CResultMeta
    Set result = New CResultMeta
    result.SetResult mInternalStatus, mResultCode, mResultKind, mResultComment, _
        mDiagnosticDetails, mApplies, mCalculated
    result.SourceKeys = mSourceKeys
    Set Clone = result
End Function

' Добавляет пояснение проверки в общий комментарий без повторов. Успешные
' и неприменимые проверки не загромождают сводку; предупреждения и причины
' недоступности сохраняются. prepared = True означает, что текст уже собран
' и его не нужно повторно размечать даже при общем статусе OK.
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

' Возвращает короткое название проверки, к которой относится пояснение.
' Используется при сборке общего комментария и не выбирает вид расчёта или статус.
Private Function ResultKindLabel() As String
    Select Case mResultKind
        Case rkDirectState: ResultKindLabel = "НДС"
        Case rkCapacity: ResultKindLabel = "Несущая способность"
        Case rkCrackFormation: ResultKindLabel = "Трещинообразование"
        Case rkCrackWidth: ResultKindLabel = "Ширина трещин"
        Case rkLongitudinalCrack: ResultKindLabel = "Продольные трещины"
        Case rkStability: ResultKindLabel = "Устойчивость"
        Case Else: ResultKindLabel = "Расчет"
    End Select
End Function

' Внутренний статус, на котором работает расчетная логика.
Public Property Get InternalStatus() As EResultInternalStatus
    InternalStatus = mInternalStatus
End Property

' Машинный код точной причины результата.
Public Property Get ResultCode() As EResultCode
    ResultCode = mResultCode
End Property

' Возвращает вид проверки: НДС, несущая способность, трещины или устойчивость.
Public Property Get ResultKind() As EResultKind
    ResultKind = mResultKind
End Property

' Короткое русское пояснение для пользователя и report.
Public Property Get ResultComment() As String
    ResultComment = mResultComment
End Property

' Короткое русское пояснение для пользователя и report.
Public Property Let ResultComment(ByVal value As String)
    mResultComment = value
End Property

' Возвращает сохранённые технические подробности решения равновесия или поиска предела.
Public Property Get DiagnosticDetails() As String
    DiagnosticDetails = mDiagnosticDetails
End Property

' Сохраняет технические подробности решения равновесия или поиска предела.
Public Property Let DiagnosticDetails(ByVal value As String)
    mDiagnosticDetails = value
End Property

' Ветка относится к LC; blocked сохраняет True, хотя нужных данных нет.
Public Property Get Applies() As Boolean
    Applies = mApplies
End Property

' Показывает, выполнялся ли сам расчёт, а не только проверка его применимости.
Public Property Get Calculated() As Boolean
    Calculated = mCalculated
End Property

' Возвращает компактную строку для diagnostics без влияния на пользовательский статус.
Public Property Get DiagnosticText() As String
    DiagnosticText = ResultKindToText(mResultKind) & ": " & _
        ResultInternalStatusToText(mInternalStatus) & " / " & ResultCodeToText(mResultCode)
    If Len(Trim$(mResultComment)) > 0 Then DiagnosticText = DiagnosticText & "; " & mResultComment
End Property
