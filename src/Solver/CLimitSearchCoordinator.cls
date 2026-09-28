VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLimitSearchCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Координатор общего поиска предельной точки
' ==========================================================================
' CLimitSearchCoordinator выбирает, какой search-метод запускать для заданной
' lambda-траектории. Он не строит материалы, не читает Config и не записывает
' Results. Эти обязанности остаются у вызывающего слоя.
'
' Сейчас координатор подключен к capacity-расчету. Для Auto сохраняется текущее
' поведение CCapacitySolver.SolveByAutoLoadPath, потому что внутри него уже есть
' проверенная последовательность UltimateStrain -> LoadMultiplier с диагностикой.

' Выполняет поиск несущей способности по стратегии из request.
' Возвращает объект результата даже при неизвестной стратегии, чтобы вызывающий
' слой мог единообразно обработать ошибку настройки.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4230, "CLimitSearchCoordinator", "Не передан запрос limit-search."

    Select Case NormalizeStrategy(request.SolutionStrategy)
        Case "Auto"
            Set ExecuteCapacity = ExecuteAutoCapacity(request)
        Case "UltimateStrain"
            Dim ultimate As CUltimateStrainSearch
            Set ultimate = New CUltimateStrainSearch
            Set ExecuteCapacity = ultimate.ExecuteCapacity(request)
        Case "LoadMultiplier"
            Dim multiplier As CLoadMultiplierSearch
            Set multiplier = New CLoadMultiplierSearch
            Set ExecuteCapacity = multiplier.ExecuteCapacity(request)
        Case Else
            Dim invalidResult As CLimitSearchResult
            Set invalidResult = New CLimitSearchResult
            invalidResult.InitializeInvalidConfiguration request.SolutionStrategy, _
                "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
            Set ExecuteCapacity = invalidResult
    End Select
End Function

' Сохраняет принятую Auto-логику: чисто осевая траектория сразу идет в
' LoadMultiplier, остальные траектории сначала пробуют UltimateStrain и только
' затем fallback на LoadMultiplier внутри CCapacitySolver.
Private Function ExecuteAutoCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim cap As CCapacitySolver
    Set cap = request.CapacitySolver

    If request.PureAxialForMethod Then
        cap.SolveByLoadPathMultiplier request.Section, request.ConcreteMaterial, request.SteelMaterial, _
            request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
            request.MyOffset, request.MyBase, True
    Else
        cap.SolveByAutoLoadPath request.Section, request.ConcreteMaterial, request.SteelMaterial, _
            request.NOffset, request.NBase, request.MxOffset, request.MxBase, _
            request.MyOffset, request.MyBase, False, request.ForceOnly
    End If

    Dim result As CLimitSearchResult
    Set result = New CLimitSearchResult
    result.InitializeFromCapacitySolver cap, request.SolutionStrategy
    Set ExecuteAutoCapacity = result
End Function

' Нормализует пользовательскую строку стратегии без зависимости от Config-reader.
Private Function NormalizeStrategy(ByVal value As String) As String
    Select Case UCase$(Trim$(value))
        Case "AUTO"
            NormalizeStrategy = "Auto"
        Case "ULTIMATESTRAIN"
            NormalizeStrategy = "UltimateStrain"
        Case "LOADMULTIPLIER"
            NormalizeStrategy = "LoadMultiplier"
        Case Else
            NormalizeStrategy = Trim$(value)
    End Select
End Function
