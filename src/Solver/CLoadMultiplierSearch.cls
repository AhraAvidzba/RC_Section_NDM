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
