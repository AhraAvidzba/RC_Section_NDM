Attribute VB_Name = "modLimitSearchRequestFactory"
Option Explicit

' Создаёт объекты, связывающие инженерную задачу с общим поиском.
' Так CLimitSearchRequest зависит только от интерфейса задачи,
' а не от конкретных классов несущей способности и образования трещины.

' Подключает критерий несущей способности, сохраняя физическую постановку запроса.
Public Sub InitializeCapacitySearchRequest(ByVal request As CLimitSearchRequest, _
        ByVal capacitySolver As CCapacitySolver, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, _
        ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal forceOnly As Boolean, ByVal pureAxialForMethod As Boolean, _
        Optional ByVal materialSpec As CMaterialModelSpec = Nothing)
    If request Is Nothing Then Err.Raise vbObjectError + 4208, "modLimitSearchRequestFactory", "Не переданы исходные данные для поиска предельного состояния."
    If capacitySolver Is Nothing Then Err.Raise vbObjectError + 4200, "CLimitSearchRequest", "Не передан объект расчёта несущей способности."
    Dim problem As CCapacityLimitSearchProblem
    Set problem = New CCapacityLimitSearchProblem
    problem.Initialize capacitySolver
    request.InitializeCapacity problem, section, concreteMaterial, steelMaterial, solutionStrategy, _
        nOffset, nBase, mxOffset, mxBase, myOffset, myBase, forceOnly, pureAxialForMethod, materialSpec
End Sub

' Подключает критерий трещинообразования к тому же типизированному поиску.
Public Sub InitializeCrackFormationSearchRequest(ByVal request As CLimitSearchRequest, _
        ByVal crackCalculator As CCrackFormationCalculator, ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal solutionStrategy As String, ByVal crackFormationPath As String, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal momentMxAtReference As Double, ByVal momentMyAtReference As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double)
    If request Is Nothing Then Err.Raise vbObjectError + 4208, "modLimitSearchRequestFactory", "Не переданы исходные данные для поиска предельного состояния."
    If crackCalculator Is Nothing Then Err.Raise vbObjectError + 4204, "CLimitSearchRequest", "Не передан объект, который должен определить нагрузку образования нормальной трещины."
    Dim problem As CCrackLimitSearchProblem
    Set problem = New CCrackLimitSearchProblem
    problem.Initialize crackCalculator
    request.InitializeCrackFormation problem, section, concreteMaterial, steelMaterial, _
        solutionStrategy, crackFormationPath, targetN, targetMx, targetMy, _
        momentMxAtReference, momentMyAtReference, referenceX, referenceY
End Sub
