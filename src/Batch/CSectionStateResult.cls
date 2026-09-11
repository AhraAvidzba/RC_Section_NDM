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

' Создает named-state по результату CSectionSolver.
Public Sub InitializeFromSolver(ByVal stateType As ESectionStateType, _
        ByVal materialModelRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal solver As CSectionSolver, Optional ByVal extensionUsed As Boolean = False, _
        Optional ByVal statusText As String = vbNullString)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3970, "CSectionStateResult", "MaterialModelSpec не передан."
    If solver Is Nothing Then Err.Raise vbObjectError + 3971, "CSectionStateResult", "Solver для named-state не передан."

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
