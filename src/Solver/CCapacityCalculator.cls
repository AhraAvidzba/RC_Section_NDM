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
        ByVal loadState As CSectionLoadState, ByVal strengthReferenceX As Double, _
        ByVal strengthReferenceY As Double, ByVal statusPolicy As CBatchStatusPolicy, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepthForMx As Double, _
        ByVal concreteDepthForMy As Double) As CCapacityResult
    Dim result As CCapacityResult
    Set result = New CCapacityResult
    If statusPolicy Is Nothing Then Set statusPolicy = New CBatchStatusPolicy

    If loadPath Is Nothing Then
        result.InitializeInvalid statusPolicy.InputErr, "CapacityLoadPath не подготовлен."
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.Valid Then
        result.InitializeInvalid statusPolicy.InputErr, loadPath.ErrorMessage
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasPath Then
        result.InitializeSkipped statusPolicy.NotApplicable, "ZeroBaseLoad"
        Set Calculate = result
        Exit Function
    End If

    If Not loadPath.HasScaledLoad Then
        result.InitializeInvalid statusPolicy.InputErr, _
            "Для выбранного CapacityLoadPath соответствующая нагрузка должна быть ненулевой."
        Set Calculate = result
        Exit Function
    End If

    Dim pureAxialForMethod As Boolean
    pureAxialForMethod = IsPureAxialForMethod(loadPath, loadState, strengthReferenceX, _
        strengthReferenceY, momentFilter, concreteDepthForMx, concreteDepthForMy)

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
        pureAxialForMethod, loadReferenceX, loadReferenceY, momentFilter, _
        concreteDepthForMx, concreteDepthForMy

    Set Calculate = result
End Function

' Определяет, остается ли силовая lambda-траектория чисто осевой после переноса
' к расчетному центру приведенного сечения. Это влияет только на выбор численной
' ветки Capacity: для действительно осевой траектории устойчивее сразу работать
' через LoadMultiplier, а не через вырожденный UltimateStrain.
Private Function IsPureAxialForMethod(ByVal loadPath As CCapacityLoadPath, _
        ByVal loadState As CSectionLoadState, ByVal referenceX As Double, _
        ByVal referenceY As Double, ByVal momentFilter As CMomentZeroFilter, _
        ByVal concreteDepthForMx As Double, ByVal concreteDepthForMy As Double) As Boolean
    If loadPath Is Nothing Then Exit Function
    If loadState Is Nothing Then Exit Function
    If Not loadPath.ForceOnly Then Exit Function

    IsPureAxialForMethod = MomentIsZeroForMethod(loadState.MxAboutPoint(referenceY), _
        momentFilter, concreteDepthForMx) And _
        MomentIsZeroForMethod(loadState.MyAboutPoint(referenceX), momentFilter, concreteDepthForMy)
End Function

' Проверяет инженерно нулевой момент тем же фильтром, которым batch уже
' нормализует нагрузки перед solver-ом. Без фильтра остается только маленький
' численный guard, чтобы прямые unit-тесты не зависели от Nothing-объекта.
Private Function MomentIsZeroForMethod(ByVal momentValue As Double, _
        ByVal momentFilter As CMomentZeroFilter, ByVal concreteDepth As Double) As Boolean
    If momentFilter Is Nothing Then
        MomentIsZeroForMethod = (Abs(momentValue) <= 0.000000001)
    Else
        MomentIsZeroForMethod = momentFilter.IsMomentEffectivelyZero(momentValue, concreteDepth)
    End If
End Function
