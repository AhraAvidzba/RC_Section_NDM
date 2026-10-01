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
Private mResultMeta As CResultMeta ' Типизированный статус named-state для будущего StateRepository/result graph.
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

' Создает named-state по результату CSectionSolver.
Public Sub InitializeFromSolver(ByVal stateType As ESectionStateType, _
        ByVal materialModelRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal solver As CSectionSolver, Optional ByVal extensionUsed As Boolean = False, _
        Optional ByVal withinPhysicalRange As Boolean = True, _
        Optional ByVal solverCallCount As Long = 0, _
        Optional ByVal diagnosticLog As String = vbNullString)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3970, "CSectionStateResult", "MaterialModelSpec не передан."
    If solver Is Nothing Then Err.Raise vbObjectError + 3971, "CSectionStateResult", "Решатель для именованного состояния не передан."

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
End Sub

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
            "Расчетное состояние не заполнено."
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
        Set BuildMetaFromSolver = BuildNumericalFailureMeta(stateType, "Solver не вернул состояние.")
    ElseIf Not solver.Converged Then
        Set BuildMetaFromSolver = BuildNumericalFailureMeta(stateType, stopReason)
    ElseIf extensionUsed Or Not withinPhysicalRange Then
        Set BuildMetaFromSolver = BuildPhysicalRangeMeta(stateType, stopReason)
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

' Создает meta для состояния, которое численно не сошлось.
Private Function BuildNumericalFailureMeta(ByVal stateType As ESectionStateType, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsNumericalFailure, rcNumericalFailure, ResultKindFromStateType(stateType), resultComment
    Set BuildNumericalFailureMeta = meta
End Function

' Создает meta для сходящегося состояния, которое найдено только за пределами
' физической диаграммы. Это физический FAIL, а не численная несходимость.
Private Function BuildPhysicalRangeMeta(ByVal stateType As ESectionStateType, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsCheckFailed, rcStateOutsidePhysicalRange, ResultKindFromStateType(stateType), resultComment
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
