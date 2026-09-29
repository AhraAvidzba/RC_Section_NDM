VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadMultiplierSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Адаптер поиска LoadMultiplier
' ==========================================================================
' Класс инкапсулирует запуск LoadMultiplier-ветки для общей lambda-траектории.
' Промежуточные probe-точки остаются внутри CCapacitySolver и не попадают в
' CStateRepository; это сохраняет правило stage 4 о локальности search-probes.

' Выполняет LoadMultiplier для request-а несущей способности.
' Вся численная логика bracket/bisection/brent/secant пока остается в
' CCapacitySolver, чтобы этот рефакторинг не изменил математику.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4220, "CLoadMultiplierSearch", "Не передан запрос limit-search."

    Dim cap As CCapacitySolver
    Set cap = request.CapacitySolver
    cap.SolveByLoadPathMultiplier request.Section, request.ConcreteMaterial, request.SteelMaterial, _
        request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
        request.MyOffset, request.MyBase, request.ForceOnly

    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCapacitySolver cap, request.SolutionStrategy
    Set ExecuteCapacity = result
End Function

' Выполняет LoadMultiplier для поиска образования нормальной трещины.
' Общий LimitSearch выбирает эту ветку, а сам одномерный расчет lambda_crc
' остается внутри CCrackWidthCalculator, где уже хранится crack-контекст.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4221, "CLoadMultiplierSearch", "Не передан запрос limit-search."

    Dim crack As CCrackWidthCalculator
    Set crack = request.CrackCalculator
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = request.ConcreteMaterial
    Set steelMaterial = request.SteelMaterial

    Dim searchSucceeded As Boolean
    searchSucceeded = crack.LimitSearchRunCrackFormationLoadMultiplier(request.Section, _
        concreteMaterial, steelMaterial, Nothing, request.TargetN, _
        request.MomentMxAtReference, request.MomentMyAtReference, _
        request.ReferenceX, request.ReferenceY)

    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCrackFormation crack, request.SolutionStrategy, searchSucceeded
    Set ExecuteCrackFormation = result
End Function
