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
' Сводный результат расчета трещин одного LC
' ==========================================================================
' CCrackResult объединяет три самостоятельных смысловых результата:
' образование нормальной трещины, ширину раскрытия нормальной трещины и
' продольные трещины по сжатому бетону. Он не пишет Results и не запускает
' solver; это контейнер result-tree между orchestration-слоем и batch-снимком.

Private mFormation As CCrackFormationResult ' Результат поиска Ncrc/Mxy,crc.
Private mCurrentStateMeta As CResultMeta ' Meta обязательного CurrentCrackedState, если он уже решался.
Private mWidth As CCrackWidthResult ' Результат формульного расчета a_crc.
Private mLongitudinal As CLongitudinalCrackResult ' Результат проверки продольных трещин.
Private mPolicy As CResultStatusPolicy ' Общая политика typed/display приоритета.

' Создает нейтральные ветви, не создавая физические State при отсутствии расчета.
Private Sub Class_Initialize()
    Initialize Nothing, Nothing, Nothing, Nothing
End Sub

' Заполняет все доступные ветви расчета трещин. Любая ветвь может быть Nothing,
' если расчет еще не дошел до нее или она не нужна для выбранного профиля.
Public Sub Initialize(ByVal formation As CCrackFormationResult, _
        ByVal currentStateMeta As CResultMeta, ByVal width As CCrackWidthResult, _
        ByVal longitudinal As CLongitudinalCrackResult)
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

Public Property Get Formation() As CCrackFormationResult
    Set Formation = mFormation
End Property

Public Property Get Width() As CCrackWidthResult
    Set Width = mWidth
End Property

Public Property Get CurrentStateMeta() As CResultMeta
    If Not mCurrentStateMeta Is Nothing Then Set CurrentStateMeta = mCurrentStateMeta.Clone
End Property

Public Property Get Longitudinal() As CLongitudinalCrackResult
    Set Longitudinal = mLongitudinal
End Property

' Собирает итог нормальных трещин из готовых Formation/CurrentState/Width.
' Недостижение образования делает ширину неприменимой, но не скрывает отказ
' обязательного текущего НДС. Новые solve или инженерные формулы здесь не выполняются.
Public Property Get NormalMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim formation As CResultMeta, width As CResultMeta, currentProblem As CResultMeta, formationWarning As CResultMeta
    Set formation = mFormation.ResultMeta
    Set width = mWidth.ResultMeta
    If mPolicy.IsFailedOverall(mPolicy.ExternalStatus(mCurrentStateMeta)) Or _
            mCurrentStateMeta.InternalStatus = rsSuccessWithWarning Then Set currentProblem = mCurrentStateMeta
    If formation.InternalStatus = rsSuccessWithWarning Then Set formationWarning = formation
    Dim comment As String
    formation.AppendCommentTo comment
    mCurrentStateMeta.AppendCommentTo comment
    width.AppendCommentTo comment
    If mPolicy.IsFailedOverall(mPolicy.ExternalStatus(formation)) Then
        Set NormalMeta = mPolicy.WorstResultMeta(rkCrackFormation, formation, currentProblem)
    ElseIf formation.InternalStatus = rsSuccess And _
            (formation.ResultCode = rcCriterionNotReached Or formation.ResultCode = rcCrackNotFormed) Then
        Set NormalMeta = mPolicy.WorstResultMeta(rkCrackFormation, formation, currentProblem)
    ElseIf mFormation.CrackFormed Then
        Set NormalMeta = mPolicy.WorstResultMeta(rkCrackWidth, currentProblem, width, formationWarning)
    Else
        Set NormalMeta = mPolicy.WorstResultMeta(rkCrackWidth, formation, currentProblem, width)
    End If
    NormalMeta.ResultComment = comment
End Property

' Возвращает итог всей подробной таблицы трещин, включая независимую проверку
' продольных трещин. Комментарии собираются в порядке расчетных стадий, только
' из этого subtree; writer не добавляет чужие причины и не склеивает строки.
Public Property Get ResultMeta() As CResultMeta
    If mPolicy Is Nothing Then Set mPolicy = New CResultStatusPolicy
    Dim normal As CResultMeta, longitudinal As CResultMeta, comment As String
    Set normal = NormalMeta
    Set longitudinal = mLongitudinal.ResultMeta
    normal.AppendCommentTo comment, True
    longitudinal.AppendCommentTo comment
    Set ResultMeta = mPolicy.WorstResultMeta(rkGeneric, normal, longitudinal)
    ResultMeta.ResultComment = comment
End Property
