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

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
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

' Отмечает весь crack workflow как отключенный профилем. Численные State
' и формулы не создаются; каждый leaf сохраняет именно NotRequested.
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

' Сохраняет ошибку общих исходных данных до запуска crack workflow.
' Formation получает точную причину валидации, а Current/Width/Longitudinal
' остаются заблокированными зависимыми этапами без фиктивной численной попытки.
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

Public Property Get Formation() As CCrackFormationResult
    Set Formation = mFormation
End Property

Public Property Get Width() As CCrackWidthResult
    Set Width = mWidth
End Property

' Возвращает копию причины/статуса текущего НДС, независимо от Formation
' и формульной Width. Потребитель не получает права изменить итог этого этапа.
Public Property Get CurrentStateMeta() As CResultMeta
    If Not mCurrentStateMeta Is Nothing Then Set CurrentStateMeta = mCurrentStateMeta.Clone
End Property

Public Property Get Longitudinal() As CLongitudinalCrackResult
    Set Longitudinal = mLongitudinal
End Property

' Собирает итог нормальных трещин из готовых Formation/CurrentState/Width.
' Поиск Formation и независимые формулы сохраняют собственные исходы.
' Отсутствующая точка может сопровождаться рассчитанной шириной с psi_s = 1;
' ни поиск, ни предупреждение не скрывают отказ текущего НДС или самой ширины.
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
    If mPolicy.IsFailedOverall(mPolicy.ExternalStatus(mCurrentStateMeta)) Or _
            mCurrentStateMeta.InternalStatus = rsSuccessWithWarning Then Set currentProblem = mCurrentStateMeta
    Dim comment As String
    formation.AppendCommentTo comment
    mCurrentStateMeta.AppendCommentTo comment
    AppendDependentComment width, comment
    Set NormalMeta = mPolicy.WorstResultMeta(rkCrackWidth, formation, currentProblem, width)
    NormalMeta.ResultComment = comment
End Property

' Возвращает итог всей подробной таблицы трещин, включая независимую проверку
' продольных трещин. Комментарии собираются в порядке расчетных стадий, только
' из этого subtree; writer не добавляет чужие причины и не склеивает строки.
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

' Различает общую ошибку входа до запуска от самостоятельного отказа Search.
' Зависимые этапы при этом не добавляют ту же причину второй раз в комментарий.
Private Function HasCommonInputFailure() As Boolean
    Select Case mFormation.ResultMeta.InternalStatus
        Case rsInvalidInput, rsInvalidConfiguration
            HasCommonInputFailure = (mCurrentStateMeta.InternalStatus = rsBlockedByDependency And _
                mWidth.ResultMeta.InternalStatus = rsBlockedByDependency And _
                mLongitudinal.ResultMeta.InternalStatus = rsBlockedByDependency And _
                Not mFormation.ResultMeta.Calculated)
    End Select
End Function

' Сохраняет отдельное пояснение blocked-проверки, но не повторяет в aggregate
' уже записанную полную причину CurrentCrackedState. Leaf/meta не изменяются.
' Сравнение текста служит только представлению; статус задан typed-результатом.
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

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Закрывает Formation, Width и Longitudinal вместе с их агрегатом, не
' пересчитывая проверки. Повторный вызов безопасен для общих ссылок workflow.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mFormation Is Nothing Then mFormation.Freeze
    If Not mWidth Is Nothing Then mWidth.Freeze
    If Not mLongitudinal Is Nothing Then mLongitudinal.Freeze
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCrackResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
