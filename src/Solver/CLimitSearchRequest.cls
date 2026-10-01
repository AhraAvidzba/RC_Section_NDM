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
' выбора численного метода. Класс ничего не решает сам: его задача передать
' координатору LimitSearch уже подготовленные материалы, сечение, путь нагрузки
' и инженерный контекст результата.
'
' Capacity и CrackFormation используют один и тот же уровень маршрутизации
' UltimateStrain / LoadMultiplier. Различаются только критерий предела,
' материал и смысл траектории, а не сам механизм выбора search-ветки.

Private mProblem As ILimitSearchProblem ' Общий callback инженерной задачи для LimitSearch.
Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех lambda-проб.
Private mConcreteMaterial As Object ' Диаграмма бетона для выбранной инженерной задачи.
Private mSteelMaterial As Object ' Диаграмма арматуры для выбранной инженерной задачи.
Private mMaterialSpec As CMaterialModelSpec ' Спецификация для снимка конечного named-state, не для выбора критерия.
Private mSolutionStrategy As String ' Auto, UltimateStrain или LoadMultiplier.
Private mNOffset As Double ' Постоянная часть N в траектории N = N0 + lambda*Nbase.
Private mNBase As Double ' Масштабируемая часть N в lambda-траектории.
Private mMxOffset As Double ' Постоянная часть Mx.
Private mMxBase As Double ' Масштабируемая часть Mx.
Private mMyOffset As Double ' Постоянная часть My.
Private mMyBase As Double ' Масштабируемая часть My.
Private mForceOnly As Boolean ' True, если выбранная capacity-траектория масштабирует только N.
Private mPureAxialForMethod As Boolean ' True, если после переноса к расчетному центру траектория остается чисто осевой.
Private mResultKind As EResultKind ' Инженерный смысл поиска: capacity, crack formation и т.п.
Private mCrackFormationPath As String ' Auto, LambdaMxy, LambdaN или LambdaNMxy для образования трещины.
Private mTargetN As Double ' Полная N текущего сочетания для служебных стартов CrackFormation.
Private mTargetMx As Double ' Полный Mx текущего сочетания во внутренней системе координат.
Private mTargetMy As Double ' Полный My текущего сочетания во внутренней системе координат.
Private mMomentMxAtReference As Double ' Mx относительно точки, в которой строится Mcrc.
Private mMomentMyAtReference As Double ' My относительно точки, в которой строится Mcrc.
Private mReferenceX As Double ' X расчетной точки CrackFormation во внутренней системе координат.
Private mReferenceY As Double ' Y расчетной точки CrackFormation во внутренней системе координат.

' Заполняет request для поиска несущей способности.
' Материалы и capacitySolver передаются готовыми, чтобы общий search-слой
' не зависел от Config, расчетного профиля и правил построения диаграмм.
Public Sub InitializeCapacity(ByVal capacitySolver As CCapacitySolver, _
        ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal forceOnly As Boolean, ByVal pureAxialForMethod As Boolean, _
        Optional ByVal materialSpec As CMaterialModelSpec = Nothing)
    If capacitySolver Is Nothing Then Err.Raise vbObjectError + 4200, "CLimitSearchRequest", "Не передан solver поиска несущей способности."
    If section Is Nothing Then Err.Raise vbObjectError + 4201, "CLimitSearchRequest", "Не передана модель сечения для limit-search."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 4202, "CLimitSearchRequest", "Не передана диаграмма бетона для limit-search."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 4203, "CLimitSearchRequest", "Не передана диаграмма арматуры для limit-search."

    Dim problem As CCapacityLimitSearchProblem
    Set problem = New CCapacityLimitSearchProblem
    problem.Initialize capacitySolver
    Set mProblem = problem
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
    Set mMaterialSpec = Nothing
    If Not materialSpec Is Nothing Then Set mMaterialSpec = materialSpec.Clone
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

' Заполняет request для поиска стадии образования нормальной трещины.
' На вход уже приходят материалы модели CrackInitiation; общий LimitSearch
' не выбирает профили и не знает о листе Config.
Public Sub InitializeCrackFormation(ByVal crackCalculator As CCrackFormationCalculator, _
        ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, ByVal crackFormationPath As String, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4204, "CLimitSearchRequest", "Не передан расчетчик трещинообразования."
    If section Is Nothing Then Err.Raise vbObjectError + 4205, "CLimitSearchRequest", "Не передана модель сечения для поиска трещинообразования."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 4206, "CLimitSearchRequest", "Не передана диаграмма бетона для поиска трещинообразования."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 4207, "CLimitSearchRequest", "Не передана диаграмма арматуры для поиска трещинообразования."

    Dim problem As CCrackLimitSearchProblem
    Set problem = New CCrackLimitSearchProblem
    problem.Initialize crackCalculator
    Set mProblem = problem
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
    mSolutionStrategy = Trim$(solutionStrategy)
    mCrackFormationPath = Trim$(crackFormationPath)
    Set mMaterialSpec = Nothing
    mTargetN = targetN
    mTargetMx = targetMx
    mTargetMy = targetMy
    mMomentMxAtReference = momentMxAtReference
    mMomentMyAtReference = momentMyAtReference
    mReferenceX = referenceX
    mReferenceY = referenceY
    mResultKind = rkCrackFormation
End Sub

' Заполняет request уже готовой доменной задачей LimitSearch.
' Этот вход нужен для контрактных тестов и будущих инженерных задач, где
' callback-объект уже создан снаружи. Он не заменяет строгие доменные
' фабрики Capacity/CrackFormation: те по-прежнему проверяют обязательные
' материалы и расчетную модель.
Public Sub InitializeWithProblem(ByVal problem As ILimitSearchProblem, _
        ByVal resultKind As EResultKind, ByVal solutionStrategy As String, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal crackFormationPath As String = vbNullString)
    If problem Is Nothing Then Err.Raise vbObjectError + 4208, _
        "CLimitSearchRequest", "Не передана доменная задача LimitSearch."

    Set mProblem = problem
    Set mSection = Nothing
    Set mConcreteMaterial = Nothing
    Set mSteelMaterial = Nothing
    Set mMaterialSpec = Nothing
    mSolutionStrategy = Trim$(solutionStrategy)
    mNOffset = nOffset
    mNBase = nBase
    mMxOffset = mxOffset
    mMxBase = mxBase
    mMyOffset = myOffset
    mMyBase = myBase
    mForceOnly = False
    mPureAxialForMethod = False
    mResultKind = resultKind
    mCrackFormationPath = Trim$(crackFormationPath)
    mTargetN = 0#
    mTargetMx = 0#
    mTargetMy = 0#
    mMomentMxAtReference = 0#
    mMomentMyAtReference = 0#
    mReferenceX = 0#
    mReferenceY = 0#
End Sub

Public Property Get Problem() As ILimitSearchProblem
    Set Problem = mProblem
End Property

Public Property Get Section() As CSectionModel
    Set Section = mSection
End Property

' Возвращает небольшой снимок модели конечного состояния. Низкоуровневые
' задачи с тестовым Object material API могут не иметь profile-spec.
Public Property Get MaterialSpec() As CMaterialModelSpec
    If Not mMaterialSpec Is Nothing Then Set MaterialSpec = mMaterialSpec.Clone
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

Public Property Get CrackFormationPath() As String
    CrackFormationPath = mCrackFormationPath
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

Public Property Get MomentMxAtReference() As Double
    MomentMxAtReference = mMomentMxAtReference
End Property

Public Property Get MomentMyAtReference() As Double
    MomentMyAtReference = mMomentMyAtReference
End Property

Public Property Get ReferenceX() As Double
    ReferenceX = mReferenceX
End Property

Public Property Get ReferenceY() As Double
    ReferenceY = mReferenceY
End Property
