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
' Координатор уже маршрутизирует Capacity и CrackFormation через один уровень
' выбора UltimateStrain / LoadMultiplier. Различия остаются в criterion,
' load path и material spec, а не в параллельной реализации search-диспетчера.


' Выполняет поиск несущей способности по стратегии из request.
' Возвращает объект результата даже при неизвестной стратегии, чтобы вызывающий
' слой мог единообразно обработать ошибку настройки.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4230, "CLimitSearchCoordinator", "Не передан запрос limit-search."
    If request.HasLoadPathSelection Then
        Set ExecuteCapacity = ExecuteLoadPaths(request)
    Else
        Set ExecuteCapacity = ExecuteCapacityCurrentPath(request)
    End If
End Function

' Запускает численную стратегию Capacity на одной уже выбранной траектории.
Private Function ExecuteCapacityCurrentPath(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Select Case NormalizeStrategy(request.SolutionStrategy)
        Case "Auto"
            Set ExecuteCapacityCurrentPath = ExecuteAutoCapacity(request)
        Case "UltimateStrain"
            Dim ultimate As CUltimateStrainSearch
            Set ultimate = New CUltimateStrainSearch
            Set ExecuteCapacityCurrentPath = ultimate.ExecuteCapacity(request)
        Case "LoadMultiplier"
            Dim multiplier As CLoadMultiplierSearch
            Set multiplier = New CLoadMultiplierSearch
            Set ExecuteCapacityCurrentPath = multiplier.Execute(request)
        Case Else
            Dim invalidResult As CLimitSearchResult
            Set invalidResult = New CLimitSearchResult
            invalidResult.InitializeInvalidConfiguration request.SolutionStrategy, _
                "Capacity.SolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain."
            Set ExecuteCapacityCurrentPath = invalidResult
    End Select
End Function

' Выполняет поиск образования нормальной трещины через тот же coordinator,
' который выбирает метод для capacity. Отличия CrackFormation остаются в
' CCrackFormationCalculator: критерий eps_bt,ult, путь нагрузки и material spec.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4231, "CLimitSearchCoordinator", "Не передан запрос limit-search."
    If request.HasLoadPathSelection Then
        Set ExecuteCrackFormation = ExecuteLoadPaths(request)
    Else
        Set ExecuteCrackFormation = ExecuteCrackFormationCurrentPath(request)
    End If
End Function

' Возвращает единый порядок путей по нормализованным нагрузкам в точке
' приложения. Mx/My вместе являются одним моментным действием; момент от
' эксцентриситета N не превращает осевое действие в отдельную компоненту M.
' Критерий, материалы и тип инженерного расчета на порядок не влияют.
Public Function LoadPathKeys(ByVal requestedPath As String, ByVal loadState As CSectionLoadState) As Variant
    If loadState Is Nothing Then Err.Raise vbObjectError + 4232, "CLimitSearchCoordinator", "Не передано сочетание для выбора LoadPath."
    Dim descriptor As CLoadPathDescriptor
    Set descriptor = New CLoadPathDescriptor
    Dim pathKey As String
    pathKey = descriptor.NormalizeKey(requestedPath)
    If Len(pathKey) = 0 Then Err.Raise vbObjectError + 4233, "CLimitSearchCoordinator", "Неизвестный LoadPath."
    If pathKey <> "Auto" Then
        LoadPathKeys = Array(pathKey)
        Exit Function
    End If

    Dim momentPath As String
    If loadState.LoadPointMx <> 0# And loadState.LoadPointMy <> 0# Then
        momentPath = "LambdaMxy"
    ElseIf loadState.LoadPointMx <> 0# Then
        momentPath = "LambdaMx"
    ElseIf loadState.LoadPointMy <> 0# Then
        momentPath = "LambdaMy"
    End If
    If Len(momentPath) = 0 Then
        LoadPathKeys = Array("LambdaN")
    ElseIf loadState.N = 0# Then
        LoadPathKeys = Array(momentPath)
    Else
        LoadPathKeys = Array(momentPath, "LambdaN", "LambdaNMxy")
    End If
End Function

' Перебирает пути одной search-задачи. Успешно найденный физический предел
' принимается независимо от того, проходит ли текущий LC инженерную проверку.
' Ошибки ввода/конфигурации/контракта не маскируются попыткой другого пути.
Private Function ExecuteLoadPaths(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim paths As Variant
    paths = LoadPathKeys(request.RequestedLoadPath, request.LoadState)
    Dim history As String, attemptComment As String
    Dim diagnosticHistory As String, attemptDiagnostics As String
    Dim result As CLimitSearchResult
    Dim meta As CResultMeta
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        request.SelectLoadPath CStr(paths(i))
        request.Problem.BeginLoadPath request
        If Not request.SelectedLoadPath.HasScaledLoad Then
            request.Problem.SetFailure sfcInvalidInput, "Для пути " & request.SelectedLoadPath.DisplayName & _
                " в колонке LoadPath сочетания соответствующая нагрузка должна быть ненулевой."
            Set result = request.Problem.BuildResult(request, False)
            Exit For
        End If
        request.Problem.AppendDiagnostic "LoadPath: попытка " & CStr(i - LBound(paths) + 1) & _
            " из " & CStr(UBound(paths) - LBound(paths) + 1) & " - " & request.SelectedLoadPath.DisplayName & "."
        Select Case request.ResultKind
            Case rkCapacity
                Set result = ExecuteCapacityCurrentPath(request)
            Case rkCrackFormation
                Set result = ExecuteCrackFormationCurrentPath(request)
            Case Else
                Err.Raise vbObjectError + 4234, "CLimitSearchCoordinator", "Тип критерия общего LoadPath не поддержан."
        End Select
        Set meta = result.Meta
        If result.Succeeded And result.HasLimitPoint Then Exit For
        If meta.InternalStatus = rsInvalidInput Or meta.InternalStatus = rsInvalidConfiguration Or _
                meta.InternalStatus = rsInternalError Then Exit For
        If i < UBound(paths) Then
            attemptComment = "Путь " & request.SelectedLoadPath.DisplayName & ": " & Trim$(meta.ResultComment)
            If Len(history) > 0 Then history = history & " "
            history = history & attemptComment
            ' Capacity очищает рабочий журнал при подготовке нового пути.
            ' Formation уже накапливает его; общий snapshot сохраняет оба
            ' варианта без повторного включения накопленного префикса.
            attemptDiagnostics = result.DiagnosticLog
            If Len(attemptDiagnostics) > 0 Then
                If Len(diagnosticHistory) = 0 Then
                    diagnosticHistory = attemptDiagnostics
                ElseIf StrComp(Left$(attemptDiagnostics, Len(diagnosticHistory)), _
                        diagnosticHistory, vbBinaryCompare) = 0 Then
                    diagnosticHistory = attemptDiagnostics
                Else
                    diagnosticHistory = diagnosticHistory & vbCrLf & attemptDiagnostics
                End If
            End If
        End If
    Next i
    request.Problem.FinishLoadPath request, history
    result.AddLoadPathHistory history, request.SelectedLoadPath.DisplayName, diagnosticHistory
    Set ExecuteLoadPaths = result
End Function

' Сохраняет принятую Auto-логику: чисто осевая траектория сразу идет в
' LoadMultiplier, остальные траектории сначала пробуют UltimateStrain и только
' затем fallback на LoadMultiplier. Обе ветки запускаются через общий
' LimitSearch-слой, а CCapacitySolver остается только доменным callback-ом.
Private Function ExecuteAutoCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request.PureAxialForMethod Then
        Dim axialMultiplier As CLoadMultiplierSearch
        Set axialMultiplier = New CLoadMultiplierSearch
        Set ExecuteAutoCapacity = axialMultiplier.Execute(request)
        Exit Function
    End If

    Dim ultimate As CUltimateStrainSearch
    Set ultimate = New CUltimateStrainSearch

    Dim firstResult As CLimitSearchResult
    Set firstResult = ultimate.ExecuteCapacity(request)
    If firstResult.Succeeded Then
        Set ExecuteAutoCapacity = firstResult
        Exit Function
    End If
    If request.Problem.StoppedTerminally Then
        Set ExecuteAutoCapacity = firstResult
        Exit Function
    End If

    Set ExecuteAutoCapacity = ExecuteAutoCapacityLoadMultiplierFallback(request)
End Function

' Повторяет Auto-поиск capacity через LoadMultiplier после неудачного
' UltimateStrain и возвращает обновленный snapshot результата. Диагностика
' первой попытки сохраняется перед тем, как LoadMultiplier очистит поля solver-а.
Private Function ExecuteAutoCapacityLoadMultiplierFallback(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim firstStopReason As String
    Dim firstLimitState As String
    Dim firstDiagnostics As String
    firstStopReason = request.Problem.StopReason
    firstLimitState = request.Problem.LimitState
    firstDiagnostics = request.Problem.DiagnosticLog

    Dim multiplier As CLoadMultiplierSearch
    Set multiplier = New CLoadMultiplierSearch
    Dim ignoredResult As CLimitSearchResult
    Set ignoredResult = multiplier.Execute(request)

    request.Problem.MergeFallback firstLimitState, firstStopReason, _
        firstDiagnostics, _
        "Auto: первая попытка UltimateStrain не дала принятую предельную точку", _
        "Auto: переходим к LoadMultiplier по той же lambda-траектории.", _
        "Auto: UltimateStrain не сошелся, предельная точка найдена через LoadMultiplier."

    Set ExecuteAutoCapacityLoadMultiplierFallback = request.Problem.BuildResult(request, False)
End Function


' Запускает выбранный метод поиска на уже выбранном пути образования трещины.
' Для Auto-стратегии сначала используется UltimateStrain, а при численной
' неудаче - LoadMultiplier по той же физической траектории.
Private Function ExecuteCrackFormationCurrentPath(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Select Case NormalizeStrategy(request.SolutionStrategy)
        Case "LoadMultiplier"
            Dim multiplier As CLoadMultiplierSearch
            Set multiplier = New CLoadMultiplierSearch
            Set ExecuteCrackFormationCurrentPath = multiplier.Execute(request)
        Case "UltimateStrain"
            Dim ultimate As CUltimateStrainSearch
            Set ultimate = New CUltimateStrainSearch
            Set ExecuteCrackFormationCurrentPath = ultimate.ExecuteCrackFormation(request)
        Case "Auto"
            Set ExecuteCrackFormationCurrentPath = ExecuteAutoStrategyCrackFormation(request)
        Case Else
            Dim invalidResult As CLimitSearchResult
            Set invalidResult = New CLimitSearchResult
            invalidResult.InitializeInvalidConfiguration request.SolutionStrategy, _
                "SLS.Crack.InitiationSolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain.", _
                rkCrackFormation
            Set ExecuteCrackFormationCurrentPath = invalidResult
    End Select
End Function

' Внутренняя Auto-стратегия одного crack-пути: сначала прямой критерий
' UltimateStrain, затем fallback LoadMultiplier, если прямой поиск не дал
' принятого результата и не завершился физически окончательным состоянием.
Private Function ExecuteAutoStrategyCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim ultimate As CUltimateStrainSearch
    Set ultimate = New CUltimateStrainSearch

    Dim firstResult As CLimitSearchResult
    Set firstResult = ultimate.ExecuteCrackFormation(request)
    If firstResult.Succeeded Then
        Set ExecuteAutoStrategyCrackFormation = firstResult
        Exit Function
    End If
    If request.Problem.StoppedTerminally Then
        Set ExecuteAutoStrategyCrackFormation = firstResult
        Exit Function
    End If

    Dim firstStopReason As String
    firstStopReason = request.Problem.StopReason
    request.Problem.AppendDiagnostic _
        "UltimateStrain не дал точку образования трещины: " & firstStopReason
    request.Problem.AppendDiagnostic _
        "Пробуем LoadMultiplier по тому же пути образования трещины."
    request.Problem.ClearStopReason

    Dim multiplier As CLoadMultiplierSearch
    Set multiplier = New CLoadMultiplierSearch
    Dim secondResult As CLimitSearchResult
    Set secondResult = multiplier.Execute(request)
    If Not secondResult.Succeeded And Len(request.Problem.StopReason) = 0 Then _
        request.Problem.SetStopReasonIfEmpty firstStopReason
    Set ExecuteAutoStrategyCrackFormation = secondResult
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
