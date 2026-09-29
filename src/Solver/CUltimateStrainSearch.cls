VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUltimateStrainSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Адаптер поиска UltimateStrain
' ==========================================================================
' Класс является тонкой оболочкой над текущей реализацией CCapacitySolver.
' Его задача не менять алгоритм, а отделить выбор search-ветки от batch-класса.
' Для силовой траектории с эксцентриситетом сохраняется утвержденный fallback через
' SolveByAutoLoadPath: сначала UltimateStrain, затем LoadMultiplier при неудаче.

' Выполняет ветку UltimateStrain для request-а несущей способности.
' Чисто осевая траектория переводится в LoadMultiplier тем же правилом, что и
' в основной capacity-логике.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4210, "CUltimateStrainSearch", "Не передан запрос limit-search."

    Dim cap As CCapacitySolver
    Set cap = request.CapacitySolver

    If request.PureAxialForMethod Then
        cap.SolveByLoadPathMultiplier request.Section, request.ConcreteMaterial, request.SteelMaterial, _
            request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
            request.MyOffset, request.MyBase, True
    ElseIf request.ForceOnly Then
        cap.SolveByAutoLoadPath request.Section, request.ConcreteMaterial, request.SteelMaterial, _
            request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
            request.MyOffset, request.MyBase, False, True
    Else
        cap.SolveByUltimateLoadPath request.Section, request.ConcreteMaterial, request.SteelMaterial, _
            request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
            request.MyOffset, request.MyBase
    End If

    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCapacitySolver cap, request.SolutionStrategy
    Set ExecuteCapacity = result
End Function

' Выполняет ветку UltimateStrain для поиска образования нормальной трещины.
' Низкоуровневый расчет остается в CCrackWidthCalculator: там находятся
' критерий eps_bt,ult, сохранение Before/After-состояний и crack-диагностика.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4211, "CUltimateStrainSearch", "Не передан запрос limit-search."

    Dim crack As CCrackWidthCalculator
    Set crack = request.CrackCalculator
    Dim concreteMaterial As CMaterialDiagram
    Dim steelMaterial As CMaterialDiagram
    Set concreteMaterial = request.ConcreteMaterial
    Set steelMaterial = request.SteelMaterial

    Dim searchSucceeded As Boolean
    searchSucceeded = crack.LimitSearchRunCrackFormationUltimate(request.Section, _
        concreteMaterial, steelMaterial, request.TargetN, _
        request.TargetMx, request.TargetMy, request.MomentMxAtReference, _
        request.MomentMyAtReference, request.ReferenceX, request.ReferenceY)

    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCrackFormation crack, request.SolutionStrategy, searchSucceeded
    Set ExecuteCrackFormation = result
End Function
