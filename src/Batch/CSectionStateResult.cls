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
' BeforeMcrcState/AfterMcrcState. CSectionStateResult хранит плоскость деформаций и
' фактическую спецификацию материала для такого named-state.
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
Private mExtensionUsed As Boolean ' True только для StateSolution с numerical extension.
Private mConverged As Boolean ' Сошелся ли solver, из которого взята плоскость.
Private mStopReason As String ' Служебная причина остановки solver-а.
Private mStatus As String ' Короткий внутренний статус, если он нужен writer-у/тестам.
Private mResultMeta As CResultMeta ' Типизированный статус named-state для будущего StateRepository/result graph.
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
        Optional ByVal statusText As String = vbNullString)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3970, "CSectionStateResult", "MaterialModelSpec не передан."
    If solver Is Nothing Then Err.Raise vbObjectError + 3971, "CSectionStateResult", "Решатель для именованного состояния не передан."

    mStateType = stateType
    mMaterialModelRole = materialModelRole
    Set mMaterialSpec = materialSpec.Clone
    mEpsilon0 = solver.Epsilon0
    mKappaX = solver.KappaX
    mKappaY = solver.KappaY
    mExtensionUsed = extensionUsed
    mConverged = solver.Converged
    mStopReason = solver.StopReason
    mStatus = statusText
    Set mResultMeta = BuildMetaFromSolver(stateType, solver, statusText, mStopReason)
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

' Создает named-state по уже известной плоскости деформаций.
' Используется для capacity-точки, когда CCapacitySolver отдает итоговую
' плоскость через LastSolver.
Public Sub InitializeFromPlane(ByVal stateType As ESectionStateType, _
        ByVal materialModelRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        Optional ByVal extensionUsed As Boolean = False, Optional ByVal statusText As String = vbNullString, _
        Optional ByVal stopReason As String = vbNullString)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3972, "CSectionStateResult", "MaterialModelSpec не передан."

    mStateType = stateType
    mMaterialModelRole = materialModelRole
    Set mMaterialSpec = materialSpec.Clone
    mEpsilon0 = eps0
    mKappaX = kx
    mKappaY = ky
    mExtensionUsed = extensionUsed
    mConverged = True
    mStatus = statusText
    mStopReason = stopReason
    Set mResultMeta = BuildMetaFromStatus(stateType, statusText, stopReason)
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

Public Property Get MaterialSpec() As CMaterialModelSpec
    Set MaterialSpec = mMaterialSpec
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

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then Set mResultMeta = BuildMetaFromStatus(mStateType, mStatus, mStopReason)
    Set ResultMeta = mResultMeta
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
Private Function BuildMetaFromSolver(ByVal stateType As ESectionStateType, ByVal solver As CSectionSolver, _
        ByVal statusText As String, ByVal stopReason As String) As CResultMeta
    If Len(Trim$(statusText)) = 0 Then
        If solver Is Nothing Then
            Set BuildMetaFromSolver = MetaFromStatus(stateType, "NumFail", "Solver не вернул состояние.")
        ElseIf solver.Converged Then
            Set BuildMetaFromSolver = MetaFromStatus(stateType, "OK", stopReason)
        ElseIf IsInputStateFailure(stopReason) Then
            Set BuildMetaFromSolver = MetaFromStatus(stateType, "InputErr", stopReason)
        Else
            Set BuildMetaFromSolver = MetaFromStatus(stateType, "NumFail", stopReason)
        End If
    Else
        Set BuildMetaFromSolver = MetaFromStatus(stateType, statusText, stopReason)
    End If
End Function

' Создает ResultMeta для named-state, построенного из готовой плоскости.
Private Function BuildMetaFromStatus(ByVal stateType As ESectionStateType, ByVal statusText As String, _
        ByVal resultComment As String) As CResultMeta
    If Len(Trim$(statusText)) = 0 Then statusText = "OK"
    Set BuildMetaFromStatus = MetaFromStatus(stateType, statusText, resultComment)
End Function

' Переводит status named-state в типизированный ResultMeta через общую policy.
Private Function MetaFromStatus(ByVal stateType As ESectionStateType, ByVal statusText As String, _
        ByVal resultComment As String) As CResultMeta
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Set MetaFromStatus = policy.MetaFromLegacyStatus(statusText, ResultKindFromStateType(stateType), resultComment)
End Function

' Определяет инженерный контекст результата по типу named-state.
Private Function ResultKindFromStateType(ByVal stateType As ESectionStateType) As EResultKind
    Select Case stateType
        Case sstCapacityState
            ResultKindFromStateType = rkCapacity
        Case sstBeforeMcrcState, sstAfterMcrcState
            ResultKindFromStateType = rkCrackFormation
        Case sstCrackedState
            ResultKindFromStateType = rkDirectState
        Case Else
            ResultKindFromStateType = rkDirectState
    End Select
End Function

' Отличает ошибку исходных данных solver-а от настоящей численной несходимости.
' Это нужно, чтобы ResultMeta named-state не превращал InputError в NumFail.
Private Function IsInputStateFailure(ByVal reason As String) As Boolean
    IsInputStateFailure = (InStr(1, reason, "InputError", vbTextCompare) > 0 Or _
        InStr(1, reason, "InvalidInput", vbTextCompare) > 0 Or _
        InStr(1, reason, "InvalidConfiguration", vbTextCompare) > 0)
End Function
