VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLimitSearchRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Запрос общего поиска предельной точки
' ==========================================================================
' CLimitSearchRequest хранит физическую постановку lambda-поиска и настройки
' выбора алгоритма, но не выполняет расчет сам. Класс нужен как первый общий
' слой между инженерной задачей (capacity, crack formation) и конкретными
' реализациями UltimateStrain / LoadMultiplier.
'
' На этом этапе миграции request подключен к расчету несущей способности.
' Низкоуровневая математика пока остается в CCapacitySolver, чтобы не менять
' расчетную методику и результаты; общий search-слой только централизует
' маршрутизацию стратегии и передачу lambda-траектории.

Private mCapacitySolver As CCapacitySolver ' Настроенный solver, который фактически выполняет численный поиск.
Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех lambda-проб.
Private mConcreteMaterial As Object ' Диаграмма бетона для выбранной инженерной задачи.
Private mSteelMaterial As Object ' Диаграмма арматуры для выбранной инженерной задачи.
Private mSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier.
Private mNOffset As Double ' Постоянная часть N в формуле N = N0 + lambda*Nbase.
Private mNBase As Double ' Масштабируемая часть N.
Private mMxOffset As Double ' Постоянная часть Mx.
Private mMxBase As Double ' Масштабируемая часть Mx.
Private mMyOffset As Double ' Постоянная часть My.
Private mMyBase As Double ' Масштабируемая часть My.
Private mForceOnly As Boolean ' True для траектории, где пользователь масштабирует только N.
Private mPureAxialForMethod As Boolean ' True, если после переноса к расчетному центру траектория чисто осевая.
Private mResultKind As EResultKind ' Инженерный смысл поиска: capacity, crack formation и т.п.

' Заполняет request для поиска несущей способности.
' Материалы и capacitySolver передаются готовыми, чтобы этот класс не знал о
' Config, профилях и правилах построения диаграмм.
Public Sub InitializeCapacity(ByVal capacitySolver As CCapacitySolver, _
        ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal forceOnly As Boolean, ByVal pureAxialForMethod As Boolean)
    If capacitySolver Is Nothing Then Err.Raise vbObjectError + 4200, "CLimitSearchRequest", "Не передан solver поиска несущей способности."
    If section Is Nothing Then Err.Raise vbObjectError + 4201, "CLimitSearchRequest", "Не передана модель сечения для limit-search."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 4202, "CLimitSearchRequest", "Не передана диаграмма бетона для limit-search."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 4203, "CLimitSearchRequest", "Не передана диаграмма арматуры для limit-search."

    Set mCapacitySolver = capacitySolver
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
    mSolutionStrategy = Trim$(solutionStrategy)
    mNOffset = nOffset
    mNBase = nBase
    mMxOffset = mxOffset
    mMxBase = mxBase
    mMyOffset = myOffset
    mMyBase = myBase
    mForceOnly = forceOnly
    mPureAxialForMethod = pureAxialForMethod
    mResultKind = rkCapacity
End Sub

Public Property Get CapacitySolver() As CCapacitySolver
    Set CapacitySolver = mCapacitySolver
End Property

Public Property Get Section() As CSectionModel
    Set Section = mSection
End Property

Public Property Get ConcreteMaterial() As Object
    Set ConcreteMaterial = mConcreteMaterial
End Property

Public Property Get SteelMaterial() As Object
    Set SteelMaterial = mSteelMaterial
End Property

Public Property Get SolutionStrategy() As String
    SolutionStrategy = mSolutionStrategy
End Property

Public Property Get NOffset() As Double
    NOffset = mNOffset
End Property

Public Property Get NBase() As Double
    NBase = mNBase
End Property

Public Property Get MxOffset() As Double
    MxOffset = mMxOffset
End Property

Public Property Get MxBase() As Double
    MxBase = mMxBase
End Property

Public Property Get MyOffset() As Double
    MyOffset = mMyOffset
End Property

Public Property Get MyBase() As Double
    MyBase = mMyBase
End Property

Public Property Get ForceOnly() As Boolean
    ForceOnly = mForceOnly
End Property

Public Property Get PureAxialForMethod() As Boolean
    PureAxialForMethod = mPureAxialForMethod
End Property

Public Property Get ResultKind() As EResultKind
    ResultKind = mResultKind
End Property
