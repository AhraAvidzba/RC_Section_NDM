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

Private mCapacitySolver As CCapacitySolver ' Настроенный контекст несущей способности с допусками и material limits.
Private mCrackCalculator As CCrackWidthCalculator ' Контекст расчета образования трещины: критерий, диагностика и результат.
Private mSection As CSectionModel ' Расчетная модель сечения, общая для всех lambda-проб.
Private mConcreteMaterial As Object ' Диаграмма бетона для выбранной инженерной задачи.
Private mSteelMaterial As Object ' Диаграмма арматуры для выбранной инженерной задачи.
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

' Заполняет request для поиска стадии образования нормальной трещины.
' На вход уже приходят материалы модели CrackInitiation; общий LimitSearch
' не выбирает профили и не знает о листе Config.
Public Sub InitializeCrackFormation(ByVal crackCalculator As CCrackWidthCalculator, _
        ByVal section As CSectionModel, ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, ByVal crackFormationPath As String, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4204, "CLimitSearchRequest", "Не передан расчетчик трещинообразования."
    If section Is Nothing Then Err.Raise vbObjectError + 4205, "CLimitSearchRequest", "Не передана модель сечения для поиска трещинообразования."
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 4206, "CLimitSearchRequest", "Не передана диаграмма бетона для поиска трещинообразования."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 4207, "CLimitSearchRequest", "Не передана диаграмма арматуры для поиска трещинообразования."

    Set mCrackCalculator = crackCalculator
    Set mSection = section
    Set mConcreteMaterial = concreteMaterial
    Set mSteelMaterial = steelMaterial
    mSolutionStrategy = Trim$(solutionStrategy)
    mCrackFormationPath = Trim$(crackFormationPath)
    mTargetN = targetN
    mTargetMx = targetMx
    mTargetMy = targetMy
    mMomentMxAtReference = momentMxAtReference
    mMomentMyAtReference = momentMyAtReference
    mReferenceX = referenceX
    mReferenceY = referenceY
    mResultKind = rkCrackFormation
End Sub

Public Property Get CapacitySolver() As CCapacitySolver
    Set CapacitySolver = mCapacitySolver
End Property

Public Property Get CrackCalculator() As CCrackWidthCalculator
    Set CrackCalculator = mCrackCalculator
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
