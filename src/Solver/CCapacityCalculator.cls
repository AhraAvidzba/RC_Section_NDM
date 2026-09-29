VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCapacityCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Инженерная обвязка расчета несущей способности
' ==========================================================================
' CCapacityCalculator собирает один capacity-запрос поверх общей
' CLimitSearch-архитектуры. Он не хранит batch-строку, не пишет Results и не
' выбирает худшее сочетание. Его задача - получить уже подготовленные сечение,
' материалы, load path и настройки solver-а, выполнить общий LimitSearch и
' вернуть CCapacityResult.
'
' Такой слой оставляет CCapacitySolver переходным фасадом с legacy-полями, но
' больше не заставляет batch-калькулятор вручную разбирать детали search.

' Выполняет расчет несущей способности для одного LC.
' Все аргументы уже нормализованы во внутренние единицы и знаки; Config сюда
' напрямую не читается.
Public Function Calculate(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal strengthSpec As CMaterialModelSpec, ByVal loadPath As CCapacityLoadPath, _
        ByVal configuredSolver As CCapacitySolver, ByVal solutionStrategy As String, _
        ByVal pureAxialForMethod As Boolean, ByVal statusPolicy As CBatchStatusPolicy, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double) As CCapacityResult
    Dim result As CCapacityResult
    Set result = New CCapacityResult

    If loadPath Is Nothing Then
        result.InitializeInvalid statusPolicy.InputErr, "CapacityLoadPath не подготовлен."
        Set Calculate = result
        Exit Function
    End If

    Dim request As CLimitSearchRequest
    Set request = New CLimitSearchRequest
    request.InitializeCapacity configuredSolver, section, concreteMaterial, steelMaterial, _
        solutionStrategy, loadPath.NOffset, loadPath.NBase, loadPath.MxOffset, _
        loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase, loadPath.ForceOnly, _
        pureAxialForMethod

    Dim coordinator As CLimitSearchCoordinator
    Set coordinator = New CLimitSearchCoordinator

    Dim searchResult As CLimitSearchResult
    Set searchResult = coordinator.ExecuteCapacity(request)
    result.InitializeFromSearch searchResult, loadPath, strengthSpec, statusPolicy, _
        loadReferenceX, loadReferenceY, momentFilter, concreteDepthForMx, concreteDepthForMy

    Set Calculate = result
End Function
