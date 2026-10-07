VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionStateResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ИМЕНОВАННОЕ НДС ОДНОГО СОЧЕТАНИЯ
' ==========================================================================
' Один LC может иметь несколько конечных расчетных состояний: StrengthState,
' CapacityState, CrackedState и состояния образования трещины
' PreCrackState/PostCrackState. CSectionStateResult хранит плоскость деформаций,
' усилия, невязки, диагностические признаки solver-а и фактическую спецификацию
' материала для такого named-state.
'
' Класс не пишет лист Results и не решает равновесие. Он только переносит
' состояние solver-а в компактный объект, который CCombinationResult хранит
' до блочной записи snapshot.

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mStateType As ESectionStateType ' Какое именованное состояние рассчитано.
Private mMaterialModelRole As ECalculationPurpose ' Смысловая роль модели материала.
Private mMaterialSpec As CMaterialModelSpec ' Фактически использованная спецификация профиля.
Private mEpsilon0 As Double ' Деформация в начале координат расчетной модели.
Private mKappaX As Double ' Кривизна kappaX, 1/мм.
Private mKappaY As Double ' Кривизна kappaY, 1/мм.
Private mTargetN As Double ' Целевое N named-state, Н.
Private mTargetMx As Double ' Целевой Mx named-state, Н*мм.
Private mTargetMy As Double ' Целевой My named-state, Н*мм.
Private mNint As Double ' Внутреннее N, полученное интегрированием напряжений, Н.
Private mMxint As Double ' Внутренний Mx, полученный интегрированием напряжений, Н*мм.
Private mMyint As Double ' Внутренний My, полученный интегрированием напряжений, Н*мм.
Private mResidualN As Double ' Невязка Nint - Ntarget, Н.
Private mResidualMx As Double ' Невязка Mxint - Mxtarget, Н*мм.
Private mResidualMy As Double ' Невязка Myint - Mytarget, Н*мм.
Private mRelativeResidualN As Double ' Относительная невязка по N для диагностики качества состояния.
Private mRelativeResidualMx As Double ' Относительная невязка по Mx для диагностики качества состояния.
Private mRelativeResidualMy As Double ' Относительная невязка по My для диагностики качества состояния.
Private mLastResidualNorm As Double ' Норма невязки последней оценки solver-а.
Private mExtensionUsed As Boolean ' True, если найденное состояние дошло до numerical extension.
Private mWithinPhysicalRange As Boolean ' True, если state принят в пределах физической диаграммы.
Private mConverged As Boolean ' Сошелся ли solver, из которого взята плоскость.
Private mStopReason As String ' Служебная причина остановки solver-а.
Private mResultMeta As CResultMeta ' Типизированная причина результата named-state, сохраненная в repository.
Private mSolverCallCount As Long ' Число запусков CSectionSolver.Solve, связанных с получением state.
Private mIterationCount As Long ' Число итераций последнего solver-а, давшего state.
Private mDiagnosticLog As String ' Диагностика runner-а или solver-а для последующего result graph.
Private mMinConcreteStrain As Double ' Минимальная деформация бетона в этом named-state.
Private mMaxConcreteStrain As Double ' Максимальная деформация бетона в этом named-state.
Private mMinSteelStrain As Double ' Минимальная деформация арматуры в этом named-state.
Private mMaxSteelStrain As Double ' Максимальная деформация арматуры в этом named-state.
Private mMinConcreteStress As Double ' Наиболее сжимающее напряжение бетона со знаком.
Private mMaxConcreteStress As Double ' Наиболее растягивающее напряжение бетона со знаком.
Private mMinSteelStress As Double ' Наиболее сжимающее напряжение арматуры со знаком.
Private mMaxSteelStress As Double ' Наиболее растягивающее напряжение арматуры со знаком.
Private mHasConcreteCompression As Boolean ' True, если в состоянии реально есть сжатый бетон.
Private mHasConcreteTension As Boolean ' True, если бетон реально работает на растяжение.
Private mHasSteelCompression As Boolean ' True, если есть сжатая арматура.
Private mHasSteelTension As Boolean ' True, если есть растянутая арматура.
Private mSnapshotSection As CSectionModel ' Общая фактически решенная модель; ревизия запрещает подмену геометрии.
Private mSnapshotRevision As Long ' Ревизия геометрии при фиксации состояния.
Private mSnapshotConcrete As CMaterialDiagram ' Неизменяемая фактическая диаграмма бетона.
Private mSnapshotSteel As CMaterialDiagram ' Неизменяемая фактическая диаграмма арматуры.
Private Const ELEMENT_ZERO_TOLERANCE As Double = 0.000000001 ' Прежний порог классификации элемента для снимка.

' Создает named-state по результату CSectionSolver.
Public Sub InitializeFromSolver(ByVal stateType As ESectionStateType, _
        ByVal materialModelRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal solver As CSectionSolver, Optional ByVal extensionUsed As Boolean = False, _
        Optional ByVal withinPhysicalRange As Boolean = True, _
        Optional ByVal solverCallCount As Long = 0, _
        Optional ByVal diagnosticLog As String = vbNullString)
    AssertWritable
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3970, "CSectionStateResult", "MaterialModelSpec не передан."
    If solver Is Nothing Then Err.Raise vbObjectError + 3971, "CSectionStateResult", "Решатель для именованного состояния не передан."

    ClearSnapshot
    mStateType = stateType
    mMaterialModelRole = materialModelRole
    Set mMaterialSpec = materialSpec.Clone
    mEpsilon0 = solver.Epsilon0
    mKappaX = solver.KappaX
    mKappaY = solver.KappaY
    CaptureEquilibriumData solver
    mExtensionUsed = extensionUsed
    mWithinPhysicalRange = (withinPhysicalRange And solver.Converged)
    mConverged = solver.Converged
    mStopReason = solver.StopReason
    Set mResultMeta = BuildMetaFromSolver(stateType, solver, extensionUsed, withinPhysicalRange, mStopReason)
    mSolverCallCount = solverCallCount
    mIterationCount = solver.Iterations
    mDiagnosticLog = diagnosticLog
    mMinConcreteStrain = solver.MinConcreteStrain
    mMaxConcreteStrain = solver.MaxConcreteStrain
    mMinSteelStrain = solver.MinSteelStrain
    mMaxSteelStrain = solver.MaxSteelStrain
    mMinConcreteStress = solver.MinConcreteStress
    mMaxConcreteStress = solver.MaxConcreteStress
    mMinSteelStress = solver.MinSteelStress
    mMaxSteelStress = solver.MaxSteelStress
    mHasConcreteCompression = solver.HasConcreteCompression
    mHasConcreteTension = solver.HasConcreteTension
    mHasSteelCompression = solver.HasSteelCompression
    mHasSteelTension = solver.HasSteelTension
    solver.CopyMaterialContext mSnapshotSection, mSnapshotConcrete, mSnapshotSteel, extensionUsed
    If Not mSnapshotSection Is Nothing Then mSnapshotRevision = mSnapshotSection.Revision
End Sub

' Сохраняет диагностический результат запроса, для которого solve не мог
' начаться из-за отсутствующего программного контекста. Это не reusable НДС.
Public Sub InitializeFailure(ByVal request As CStateRequest, ByVal meta As CResultMeta)
    AssertWritable
    If request Is Nothing Then Err.Raise vbObjectError + 3972, "CSectionStateResult", "Запрос диагностического состояния не передан."
    If meta Is Nothing Then Err.Raise vbObjectError + 3973, "CSectionStateResult", "Причина отсутствия состояния не передана."
    ClearSnapshot
    mStateType = request.StateType
    mMaterialModelRole = request.MaterialRole
    Set mMaterialSpec = request.MaterialSpec
    mTargetN = request.TargetN
    mTargetMx = request.TargetMx
    mTargetMy = request.TargetMy
    Set mResultMeta = meta.Clone
    mStopReason = meta.ResultComment
End Sub

' Очищает все физические и диагностические поля перед заполнением snapshot,
' чтобы повторная инициализация отказом не сохраняла числа прошлого НДС.
Private Sub ClearSnapshot()
    Set mSnapshotSection = Nothing
    Set mSnapshotConcrete = Nothing
    Set mSnapshotSteel = Nothing
    mSnapshotRevision = 0
    Set mMaterialSpec = Nothing
    Set mResultMeta = Nothing
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mTargetN = 0#: mTargetMx = 0#: mTargetMy = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mLastResidualNorm = 0#
    mExtensionUsed = False: mWithinPhysicalRange = False: mConverged = False
    mStopReason = vbNullString: mDiagnosticLog = vbNullString
    mSolverCallCount = 0: mIterationCount = 0
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#
    mMinSteelStrain = 0#: mMaxSteelStrain = 0#
    mMinConcreteStress = 0#: mMaxConcreteStress = 0#
    mMinSteelStress = 0#: mMaxSteelStress = 0#
    mHasConcreteCompression = False: mHasConcreteTension = False
    mHasSteelCompression = False: mHasSteelTension = False
End Sub

' Готовит краткоживущий массив ID/type/strain/stress/physical-state по
' сохраненной плоскости и фактическим диаграммам. Solve/Search не вызываются.
' Сетка не копируется на каждый state; измененная модель явно отклоняется.
Public Function PrepareElementSnapshot(ByVal section As CSectionModel) As Variant
    If Not mConverged Then Exit Function
    If section Is Nothing Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Модель для снимка элементов не передана."
    If Not section Is mSnapshotSection Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Снимок НДС относится к другой модели сечения."
    If section.Revision <> mSnapshotRevision Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "После расчета НДС геометрия изменилась; выполните расчет для текущего сечения."
    If Not HasMaterialSnapshot Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Фактические диаграммы состояния недоступны для записи элементов."
    Dim data() As Variant, i As Long, row As Long, strain As Double, stress As Double
    ReDim data(1 To section.ConcreteCount + section.RebarCount, 1 To 5)
    row = 1
    For i = 1 To section.ConcreteCount
        strain = mEpsilon0 + mKappaX * section.ConcreteY(i) + mKappaY * section.ConcreteX(i)
        stress = mSnapshotConcrete.GetStress(strain)
        data(row, 1) = section.ConcreteID(i): data(row, 2) = "Concrete"
        data(row, 3) = strain: data(row, 4) = stress: data(row, 5) = ElementPhysicalState("Concrete", strain, stress)
        row = row + 1
    Next i
    For i = 1 To section.RebarCount
        strain = mEpsilon0 + mKappaX * section.RebarY(i) + mKappaY * section.RebarX(i)
        stress = mSnapshotSteel.GetStress(strain)
        data(row, 1) = section.RebarID(i): data(row, 2) = "Rebar"
        data(row, 3) = strain: data(row, 4) = stress: data(row, 5) = ElementPhysicalState("Rebar", strain, stress)
        row = row + 1
    Next i
    PrepareElementSnapshot = data
End Function

' Классификация относится к фактическому напряжению, не внешнему статусу LC.
Private Function ElementPhysicalState(ByVal materialType As String, ByVal strain As Double, ByVal stress As Double) As String
    If Abs(stress) <= ELEMENT_ZERO_TOLERANCE Then
        If materialType = "Concrete" And strain > ELEMENT_ZERO_TOLERANCE Then
            ElementPhysicalState = "InactiveTensionConcrete"
        Else
            ElementPhysicalState = "NearZero"
        End If
    ElseIf stress < 0# Then
        ElementPhysicalState = "Compression"
    Else
        ElementPhysicalState = "Tension"
    End If
End Function

' Снимки диаграмм закреплены при создании; выдача ссылки не разрешает мутацию.
Public Property Get HasMaterialSnapshot() As Boolean
    HasMaterialSnapshot = Not mSnapshotConcrete Is Nothing And Not mSnapshotSteel Is Nothing
End Property
Public Property Get ConcreteDiagramSnapshot() As CMaterialDiagram
    Set ConcreteDiagramSnapshot = mSnapshotConcrete
End Property
Public Property Get SteelDiagramSnapshot() As CMaterialDiagram
    Set SteelDiagramSnapshot = mSnapshotSteel
End Property

' Копирует из solver-а все данные равновесия, которые определяют качество
' найденного НДС. Target восстанавливается как Nint - residual, потому что
' сам CSectionSolver хранит невязку именно в виде внутреннее усилие минус цель.
Private Sub CaptureEquilibriumData(ByVal solver As CSectionSolver)
    mNint = solver.Nint
    mMxint = solver.Mxint
    mMyint = solver.Myint
    mResidualN = solver.ResidualN
    mResidualMx = solver.ResidualMx
    mResidualMy = solver.ResidualMy
    mTargetN = mNint - mResidualN
    mTargetMx = mMxint - mResidualMx
    mTargetMy = mMyint - mResidualMy
    mRelativeResidualN = solver.RelativeResidualN
    mRelativeResidualMx = solver.RelativeResidualMx
    mRelativeResidualMy = solver.RelativeResidualMy
    mLastResidualNorm = solver.LastResidualNorm
End Sub

Public Property Get StateType() As ESectionStateType
    StateType = mStateType
End Property

Public Property Get StateTypeText() As String
    StateTypeText = SectionStateTypeToText(mStateType)
End Property

Public Property Get MaterialModelRole() As ECalculationPurpose
    MaterialModelRole = mMaterialModelRole
End Property

Public Property Get MaterialModelRoleText() As String
    MaterialModelRoleText = MaterialModelRoleToText(mMaterialModelRole)
End Property

' Выдает независимую небольшую спецификацию, не копируя расчетную сетку.
' Внешний потребитель не может изменить модель уже сохраненного состояния.
Public Property Get MaterialSpec() As CMaterialModelSpec
    If Not mMaterialSpec Is Nothing Then Set MaterialSpec = mMaterialSpec.Clone
End Property

Public Property Get Epsilon0() As Double
    Epsilon0 = mEpsilon0
End Property

Public Property Get KappaX() As Double
    KappaX = mKappaX
End Property

Public Property Get KappaY() As Double
    KappaY = mKappaY
End Property

Public Property Get ExtensionUsed() As Boolean
    ExtensionUsed = mExtensionUsed
End Property

Public Property Get WithinPhysicalRange() As Boolean
    WithinPhysicalRange = mWithinPhysicalRange
End Property

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Преобразует сохраненную typed meta в единственный внешний словарь статусов.
' Отдельной изменяемой строки статуса у физического State нет.
Public Property Get Status() As String
    Status = ExternalStatus(ResultMeta)
End Property

' Возвращает снимок метаданных. Неинициализированный State не является
' успешным равновесием; его отсутствие означает внутреннюю ошибку контракта.
Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then
        Dim missing As CResultMeta
        Set missing = New CResultMeta
        missing.SetResult rsInternalError, rcInternalError, ResultKindFromStateType(mStateType), _
            "Расчетное состояние не заполнено.", vbNullString, True, False
        Set ResultMeta = missing
    Else
        Set ResultMeta = mResultMeta.Clone
    End If
End Property

Public Property Get InternalStatus() As EResultInternalStatus
    InternalStatus = ResultMeta.InternalStatus
End Property

Public Property Get ResultCode() As EResultCode
    ResultCode = ResultMeta.ResultCode
End Property

Public Property Get ResultComment() As String
    ResultComment = ResultMeta.ResultComment
End Property

Public Property Get TargetN() As Double
    TargetN = mTargetN
End Property

Public Property Get TargetMx() As Double
    TargetMx = mTargetMx
End Property

Public Property Get TargetMy() As Double
    TargetMy = mTargetMy
End Property

Public Property Get Nint() As Double
    Nint = mNint
End Property

Public Property Get Mxint() As Double
    Mxint = mMxint
End Property

Public Property Get Myint() As Double
    Myint = mMyint
End Property

Public Property Get ResidualN() As Double
    ResidualN = mResidualN
End Property

Public Property Get ResidualMx() As Double
    ResidualMx = mResidualMx
End Property

Public Property Get ResidualMy() As Double
    ResidualMy = mResidualMy
End Property

Public Property Get RelativeResidualN() As Double
    RelativeResidualN = mRelativeResidualN
End Property

Public Property Get RelativeResidualMx() As Double
    RelativeResidualMx = mRelativeResidualMx
End Property

Public Property Get RelativeResidualMy() As Double
    RelativeResidualMy = mRelativeResidualMy
End Property

Public Property Get LastResidualNorm() As Double
    LastResidualNorm = mLastResidualNorm
End Property

Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

Public Property Get IterationCount() As Long
    IterationCount = mIterationCount
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

Public Property Get MinConcreteStrain() As Double
    MinConcreteStrain = mMinConcreteStrain
End Property

Public Property Get MaxConcreteStrain() As Double
    MaxConcreteStrain = mMaxConcreteStrain
End Property

Public Property Get MinSteelStrain() As Double
    MinSteelStrain = mMinSteelStrain
End Property

Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

Public Property Get MinConcreteStress() As Double
    MinConcreteStress = mMinConcreteStress
End Property

Public Property Get MaxConcreteStress() As Double
    MaxConcreteStress = mMaxConcreteStress
End Property

Public Property Get MinSteelStress() As Double
    MinSteelStress = mMinSteelStress
End Property

Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

Public Property Get HasConcreteCompression() As Boolean
    HasConcreteCompression = mHasConcreteCompression
End Property

Public Property Get HasConcreteTension() As Boolean
    HasConcreteTension = mHasConcreteTension
End Property

Public Property Get HasSteelCompression() As Boolean
    HasSteelCompression = mHasSteelCompression
End Property

Public Property Get HasSteelTension() As Boolean
    HasSteelTension = mHasSteelTension
End Property

' Создает ResultMeta для named-state, построенного напрямую из CSectionSolver.
' Здесь не разбирается текст StopReason: результат определяется только
' машинными признаками solver-а и физической допустимостью диаграммы.
Private Function BuildMetaFromSolver(ByVal stateType As ESectionStateType, ByVal solver As CSectionSolver, _
        ByVal extensionUsed As Boolean, ByVal withinPhysicalRange As Boolean, _
        ByVal stopReason As String) As CResultMeta
    If solver Is Nothing Then
        Dim missing As CResultMeta
        Set missing = New CResultMeta
        missing.SetResult rsInternalError, rcSolverDidNotReturnState, ResultKindFromStateType(stateType), _
            "Решатель не вернул обязательное расчетное состояние.", vbNullString, True, False
        Set BuildMetaFromSolver = missing
    ElseIf Not solver.Converged Then
        Set BuildMetaFromSolver = BuildFailureMeta(stateType, solver, stopReason)
    ElseIf extensionUsed Or Not withinPhysicalRange Then
        Set BuildMetaFromSolver = BuildPhysicalRangeMeta(stateType, extensionUsed, stopReason)
    Else
        Set BuildMetaFromSolver = BuildSuccessMeta(stateType, stopReason)
    End If
End Function

' Создает успешную meta для готовой физической плоскости деформаций.
Private Function BuildSuccessMeta(ByVal stateType As ESectionStateType, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsSuccess, rcCheckPassed, ResultKindFromStateType(stateType), resultComment
    Set BuildSuccessMeta = meta
End Function

' Сохраняет точный машинный отказ solver-а. Ошибка входа/настроек до итераций
' не означает выполненный численный расчет; вырождение отличается от бюджета.
Private Function BuildFailureMeta(ByVal stateType As ESectionStateType, ByVal solver As CSectionSolver, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetSolverFailure solver.FailureCode, ResultKindFromStateType(stateType), _
        resultComment, (solver.Iterations > 0)
    Set BuildFailureMeta = meta
End Function

' Создает meta для сходящегося состояния, которое найдено только за пределами
' физической диаграммы. Это физический FAIL, а не численная несходимость.
Private Function BuildPhysicalRangeMeta(ByVal stateType As ESectionStateType, _
        ByVal extensionUsed As Boolean, ByVal diagnosticDetails As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    Dim comment As String
    If extensionUsed Then
        comment = "Равновесие найдено с расширением диаграмм, но деформации материала вышли за физические пределы."
    Else
        comment = "Равновесие найдено, но деформации материала вышли за физические пределы диаграмм."
    End If
    comment = comment & " Проверка по физической диаграмме не проходит."
    meta.SetResult rsCheckFailed, rcStateOutsidePhysicalRange, ResultKindFromStateType(stateType), comment, diagnosticDetails
    Set BuildPhysicalRangeMeta = meta
End Function

' Возвращает внешний статус для плоских полей Results.
Private Function ExternalStatus(ByVal meta As CResultMeta) As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    ExternalStatus = policy.ExternalStatus(meta)
End Function

' Определяет инженерный контекст результата по типу named-state.
Private Function ResultKindFromStateType(ByVal stateType As ESectionStateType) As EResultKind
    Select Case stateType
        Case sstCapacityState
            ResultKindFromStateType = rkCapacity
        Case sstPreCrackState, sstPostCrackState
            ResultKindFromStateType = rkCrackFormation
        Case sstCrackedState
            ResultKindFromStateType = rkDirectState
        Case Else
            ResultKindFromStateType = rkDirectState
    End Select
End Function

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Повторный вызов безопасен. Плоскость, усилия, признаки и meta этого НДС
' больше нельзя перезаполнить через сохраненную ссылку на state.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CSectionStateResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
