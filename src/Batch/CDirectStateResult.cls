VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDirectStateResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат прямого решения НДС одного сочетания
' ==========================================================================
' CDirectStateResult хранит все данные прямого StateSolution, которые раньше
' жили только в плоских полях CCombinationResult: внешний статус, typed ResultMeta,
' признак использования extension диаграммы, найденную плоскость деформаций,
' внутренние усилия, невязки и экстремальные деформации/напряжения.
'
' Класс не запускает solver и не принимает расчетных решений. Он является узлом
' result-tree для этапа 7 миграции: CCombinationResult пока сохраняет старый
' фасад для writer-ов, но источник meta/status постепенно переносится сюда.

Public Status As String ' Пользовательский статус direct-state блока: OK/FAIL/NumFail/InputErr/N/A.
Public ResultMeta As CResultMeta ' Типизированная причина результата direct-state.
Public ExtensionUsed As Boolean ' True, если фактический solve дошел до numerical extension диаграммы.
Public ExtensionEnabled As Boolean ' Снимок настройки Solver.DirectState.DiagramExtension для этого LC.
Public StateAvailable As Boolean ' True, если есть плоскость деформаций для вывода.
Public StateConverged As Boolean ' True, если solver сошелся численно.
Public Epsilon0 As Double ' epsilon0 в начале внутренней системы координат модели.
Public KappaX As Double ' kappaX, 1/мм.
Public KappaY As Double ' kappaY, 1/мм.
Public Nint As Double ' Внутреннее N найденного состояния, Н.
Public Mxint As Double ' Внутренний Mx найденного состояния, Н*мм.
Public Myint As Double ' Внутренний My найденного состояния, Н*мм.
Public ResidualN As Double ' Невязка равновесия по N.
Public ResidualMx As Double ' Невязка равновесия по Mx.
Public ResidualMy As Double ' Невязка равновесия по My.
Public Iterations As Long ' Число итераций solver-а для последней попытки.
Public MinConcreteStrain As Double ' Минимальная деформация бетонных волокон.
Public MaxConcreteStrain As Double ' Максимальная деформация бетонных волокон.
Public MaxConcreteCompressionStress As Double ' Максимальное по модулю сжатие бетона для summary.
Public MinSteelStrain As Double ' Минимальная деформация арматуры.
Public MaxSteelStrain As Double ' Максимальная деформация арматуры.

' Создает безопасный N/A-result, чтобы пустой direct-state не превращался в CalcErr.
Private Sub Class_Initialize()
    Clear
End Sub

' Сбрасывает объект к нейтральному состоянию, когда direct-state еще не запрошен.
Public Sub Clear()
    Status = "N/A"
    Set ResultMeta = New CResultMeta
    ResultMeta.SetNotApplicable rkDirectState
    ExtensionUsed = False
    ExtensionEnabled = False
    StateAvailable = False
    StateConverged = False
    Epsilon0 = 0#
    KappaX = 0#
    KappaY = 0#
    Nint = 0#
    Mxint = 0#
    Myint = 0#
    ResidualN = 0#
    ResidualMx = 0#
    ResidualMy = 0#
    Iterations = 0
    MinConcreteStrain = 0#
    MaxConcreteStrain = 0#
    MaxConcreteCompressionStress = 0#
    MinSteelStrain = 0#
    MaxSteelStrain = 0#
End Sub

' Устанавливает статус и синхронную typed meta без изменения численных данных state.
Public Sub SetStatus(ByVal statusText As String, Optional ByVal resultComment As String = vbNullString)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    Status = policy.ToUserStatus(statusText)
    Set ResultMeta = policy.MetaFromLegacyStatus(Status, rkDirectState, resultComment)
End Sub

' Запоминает данные direct-state, полученные напрямую из CSectionSolver.
Public Sub StoreSolver(ByVal solver As CSectionSolver, ByVal extensionEnabledValue As Boolean)
    If solver Is Nothing Then Exit Sub

    ExtensionEnabled = extensionEnabledValue
    StateAvailable = True
    StateConverged = solver.Converged
    Epsilon0 = solver.Epsilon0
    KappaX = solver.KappaX
    KappaY = solver.KappaY
    Nint = solver.Nint
    Mxint = solver.Mxint
    Myint = solver.Myint
    ResidualN = solver.ResidualN
    ResidualMx = solver.ResidualMx
    ResidualMy = solver.ResidualMy
    Iterations = solver.Iterations
    MinConcreteStrain = solver.MinConcreteStrain
    MaxConcreteStrain = solver.MaxConcreteStrain
    MaxConcreteCompressionStress = solver.MaxConcreteCompressionStress
    MinSteelStrain = solver.MinSteelStrain
    MaxSteelStrain = solver.MaxSteelStrain
End Sub

' Запоминает данные direct-state, если они пришли из reusable CSectionStateResult.
Public Sub StoreStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabledValue As Boolean)
    If stateResult Is Nothing Then Exit Sub

    ExtensionEnabled = extensionEnabledValue
    ExtensionUsed = stateResult.ExtensionUsed
    StateAvailable = True
    StateConverged = stateResult.Converged
    Epsilon0 = stateResult.Epsilon0
    KappaX = stateResult.KappaX
    KappaY = stateResult.KappaY
    Nint = stateResult.Nint
    Mxint = stateResult.Mxint
    Myint = stateResult.Myint
    ResidualN = stateResult.ResidualN
    ResidualMx = stateResult.ResidualMx
    ResidualMy = stateResult.ResidualMy
    Iterations = stateResult.IterationCount
    MinConcreteStrain = stateResult.MinConcreteStrain
    MaxConcreteStrain = stateResult.MaxConcreteStrain
    MaxConcreteCompressionStress = Abs(stateResult.MinConcreteStress)
    MinSteelStrain = stateResult.MinSteelStrain
    MaxSteelStrain = stateResult.MaxSteelStrain
    Set ResultMeta = stateResult.ResultMeta.Clone
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(ResultMeta)
End Sub

' Возвращает копию typed meta, чтобы внешний код не менял состояние result-tree.
Public Property Get Meta() As CResultMeta
    If ResultMeta Is Nothing Then SetStatus Status
    Set Meta = ResultMeta.Clone
End Property
