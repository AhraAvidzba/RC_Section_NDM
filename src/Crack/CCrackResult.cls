VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общий результат расчёта трещин одного сочетания нагрузок
' ==========================================================================
' Объединяет образование нормальной трещины, ширину её раскрытия
' и проверку продольных трещин по сжатому бетону. Получает готовые результаты
' для последующей записи в Results.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mFormation As CCrackFormationResult ' Результат поиска Ncrc/Mxy,crc.
Private mCurrentStateMeta As CResultMeta ' Статус обязательного текущего НДС с трещинами, если оно рассчитывалось.
Private mWidth As CCrackWidthResult ' Результат формульного расчета a_crc.
Private mLongitudinal As CLongitudinalCrackResult ' Результат проверки продольных трещин.
Private mPolicy As CResultStatusPolicy ' Общие правила выбора итогового статуса и его обозначения для пользователя.

' Создаёт пустые результаты проверок трещин, не подставляя НДС отсутствующего расчёта.
Private Sub Class_Initialize()
    Initialize Nothing, Nothing, Nothing, Nothing
End Sub

' Заполняет все доступные ветви расчета трещин. Любая ветвь может быть Nothing,
' если расчет еще не дошел до нее или она не нужна для выбранного профиля.
Public Sub Initialize(ByVal formation As CCrackFormationResult, _
        ByVal currentStateMeta As CResultMeta, ByVal width As CCrackWidthResult, _
        ByVal longitudinal As CLongitudinalCrackResult)
    AssertWritable
    Set mFormation = formation
    If mFormation Is Nothing Then Set mFormation = New CCrackFormationResult
    If currentStateMeta Is Nothing Then
        Set mCurrentStateMeta = New CResultMeta
        mCurrentStateMeta.SetNotApplicable rkDirectState
    Else
        Set mCurrentStateMeta = currentStateMeta.Clone
    End If
    Set mWidth = width
    If mWidth Is Nothing Then Set mWidth = New CCrackWidthResult
    Set mLongitudinal = longitudinal
    If mLongitudinal Is Nothing Then Set mLongitudinal = New CLongitudinalCrackResult
End Sub

' Отмечает все проверки трещин как отключённые профилем со статусом NotRequested.
' НДС и формулы при этом не рассчитываются.
Friend Sub SetNotRequested()
    AssertWritable
    Dim meta As CResultMeta
    Set meta = New CResultMeta: meta.SetNotRequested rkCrackFormation
    Set mFormation = New CCrackFormationResult: mFormation.InitializeUncalculated meta
    Set mCurrentStateMeta = New CResultMeta: mCurrentStateMeta.SetNotRequested rkDirectState
    Set meta = New CResultMeta: meta.SetNotRequested rkCrackWidth
    Set mWidth = New CCrackWidthResult: mWidth.InitializeFromCalculator Nothing, meta
    Set meta = New CResultMeta: meta.SetNotRequested rkLongitudinalCrack
    Set mLongitudinal = New CLongitudinalCrackResult: mLongitudinal.Initialize meta, 0#, 0#, 0#
End Sub

' Сохраняет общую ошибку ввода до начала проверок трещин. Проверка образования
' получает точную причину, а текущее НДС, ширина и продольные трещины
' отмечаются как недоступные без вымышленного численного запуска.
Friend Sub SetInputFailure(ByVal reason As String)
    AssertWritable
    Dim meta As CResultMeta
    Set meta = New CResultMeta: meta.SetResult rsInvalidInput, rcInvalidInput, rkCrackFormation, reason
    Set mFormation = New CCrackFormationResult: mFormation.InitializeUncalculated meta
    Set mCurrentStateMeta = New CResultMeta
    mCurrentStateMeta.SetBlockedByDependency rkDirectState, "исходных данных сочетания", reason
    Set meta = New CResultMeta
    meta.SetBlockedByDependency rkCrackWidth, "текущего НДС от сочетания", mCurrentStateMeta.ResultComment
    Set mWidth = New CCrackWidthResult: mWidth.InitializeFromCalculator Nothing, meta
    Set meta = New CResultMeta
    meta.SetBlockedByDependency rkLongitudinalCrack, "текущего НДС от сочетания", mCurrentStateMeta.ResultComment
    Set mLongitudinal = New CLongitudinalCrackResult: mLongitudinal.Initialize meta, 0#, 0#, 0#
End Sub

' Результат поиска Ncrc/Mxy,crc.
Public Property Get Formation() As CCrackFormationResult
    Set Formation = mFormation
End Property

' Результат формульного расчета a_crc.
Public Property Get Width() As CCrackWidthResult
    Set Width = mWidth
End Property

' Возвращает копию статуса текущего НДС независимо от образования трещины
' и формулы ширины. Готовый результат этапа через эту копию изменить нельзя.
Public Property Get CurrentStateMeta() As CResultMeta
    If Not mCurrentStateMeta Is Nothing Then Set CurrentStateMeta = mCurrentStateMeta.Clone
End Property

' Результат проверки продольных трещин.
Public Property Get Longitudinal() As CLongitudinalCrackResult
    Set Longitudinal = mLongitudinal
End Property

' Объединяет результаты образования трещины, текущего НДС и ширины.
' Если порог не найден, ширина может быть рассчитана с psi_s = 1.
' Предупреждение поиска не скрывает неудачу текущего НДС или расчёта ширины.
Public Property Get NormalMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim formation As CResultMeta, width As CResultMeta, currentProblem As CResultMeta
    Set formation = mFormation.ResultMeta
    If HasCommonInputFailure Then
        Dim inputComment As String
        formation.AppendCommentTo inputComment
        Set NormalMeta = formation.Clone
        NormalMeta.ResultComment = inputComment
        Exit Property
    End If
    Set width = mWidth.ResultMeta
    If mPolicy.IsMetaFailedOverall(mCurrentStateMeta) Or _
            mCurrentStateMeta.InternalStatus = rsSuccessWithWarning Then Set currentProblem = mCurrentStateMeta
    Dim comment As String
    formation.AppendCommentTo comment
    mCurrentStateMeta.AppendCommentTo comment
    AppendDependentComment width, comment
    Set NormalMeta = mPolicy.WorstResultMeta(rkCrackWidth, formation, currentProblem, width)
    NormalMeta.ResultComment = comment
End Property

' Собирает общий итог проверок трещин, включая независимые продольные трещины.
' Пояснения идут в порядке расчётных этапов; при записи на лист
' чужие причины к этому итогу не добавляются.
Public Property Get ResultMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim normal As CResultMeta, longitudinal As CResultMeta, comment As String
    Set normal = NormalMeta
    If HasCommonInputFailure Then
        Set ResultMeta = normal
        Exit Property
    End If
    Set longitudinal = mLongitudinal.ResultMeta
    normal.AppendCommentTo comment, True
    AppendDependentComment longitudinal, comment
    Set ResultMeta = mPolicy.WorstResultMeta(rkGeneric, normal, longitudinal)
    ResultMeta.ResultComment = comment
End Property

' Различает общую ошибку ввода до расчёта и отдельную неудачу поиска предела.
' Зависимые этапы не повторяют общую причину в комментарии второй раз.
Private Function HasCommonInputFailure() As Boolean
    Select Case mFormation.ResultMeta.InternalStatus
        Case rsInvalidInput, rsInvalidConfiguration
            HasCommonInputFailure = (mCurrentStateMeta.InternalStatus = rsBlockedByDependency And _
                mWidth.ResultMeta.InternalStatus = rsBlockedByDependency And _
                mLongitudinal.ResultMeta.InternalStatus = rsBlockedByDependency And _
                Not mFormation.ResultMeta.Calculated)
    End Select
End Function

' Сохраняет отдельное пояснение недоступной проверки, не повторяя в общем
' комментарии уже записанную причину текущего НДС. Статусы отдельных проверок
' не меняются; сравнение текста нужно только для понятного вывода.
Private Sub AppendDependentComment(ByVal meta As CResultMeta, ByRef comment As String)
    Dim prepared As CResultMeta
    Set prepared = meta.Clone
    If meta.InternalStatus = rsBlockedByDependency Then
        Dim sharedReason As String, ownText As String
        sharedReason = Trim$(mCurrentStateMeta.ResultComment)
        ownText = Trim$(meta.ResultComment)
        If Len(sharedReason) > 0 And Len(ownText) > Len(sharedReason) Then
            If InStr(1, comment, sharedReason, vbBinaryCompare) > 0 Then
                If StrComp(Right$(ownText, Len(sharedReason)), sharedReason, vbBinaryCompare) = 0 Then _
                    prepared.ResultComment = Trim$(Left$(ownText, Len(ownText) - Len(sharedReason)))
            End If
        End If
    End If
    prepared.AppendCommentTo comment
End Sub

' Завершает заполнение общего результата, образования трещины, ширины
' и продольных трещин без повторных расчётов. После этого изменения запрещены;
' повторный вызов безопасен.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mFormation Is Nothing Then mFormation.Freeze
    If Not mWidth Is Nothing Then mWidth.Freeze
    If Not mLongitudinal Is Nothing Then mLongitudinal.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCrackResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
