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

Private Const CRACK_PATH_AUTO As String = "Auto" ' Auto-перебор путей образования трещины.
Private Const CRACK_PATH_MXY As String = "LambdaMxy" ' Масштабируется только моментный вектор.
Private Const CRACK_PATH_N As String = "LambdaN" ' Масштабируется только продольная сила.
Private Const CRACK_PATH_NMXY As String = "LambdaNMxy" ' Масштабируется весь вектор N/Mx/My.

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

' Выполняет поиск образования нормальной трещины через тот же coordinator,
' который выбирает метод для capacity. Отличия CrackFormation остаются в
' CCrackWidthCalculator: критерий eps_bt,ult, путь нагрузки и material spec.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4231, "CLimitSearchCoordinator", "Не передан запрос limit-search."

    If StrComp(NormalizeCrackPath(request.CrackFormationPath), CRACK_PATH_AUTO, vbTextCompare) = 0 Then
        Set ExecuteCrackFormation = ExecuteAutoCrackFormation(request)
    Else
        request.Problem.BeginFormationAttempt NormalizeCrackPath(request.CrackFormationPath)
        Dim result As CLimitSearchResult
        Set result = ExecuteCrackFormationCurrentPath(request)
        If result.Succeeded Then _
            request.Problem.SetResolvedFormationPath NormalizeCrackPath(request.CrackFormationPath)
        Set ExecuteCrackFormation = result
    End If
End Function

' Сохраняет принятую Auto-логику: чисто осевая траектория сразу идет в
' LoadMultiplier, остальные траектории сначала пробуют UltimateStrain и только
' затем fallback на LoadMultiplier. Обе ветки запускаются через общий
' LimitSearch-слой, а CCapacitySolver остается только доменным callback-ом.
Private Function ExecuteAutoCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request.PureAxialForMethod Then
        Dim axialMultiplier As CLoadMultiplierSearch
        Set axialMultiplier = New CLoadMultiplierSearch
        Set ExecuteAutoCapacity = axialMultiplier.ExecuteCapacity(request)
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
    If request.Problem.IsInvalidInputState(request.Problem.LimitState) Then
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
    Set ignoredResult = multiplier.ExecuteCapacity(request)

    request.Problem.MergeFallback firstLimitState, firstStopReason, _
        firstDiagnostics, _
        "Auto: первая попытка UltimateStrain не дала принятую предельную точку", _
        "Auto: переходим к LoadMultiplier по той же lambda-траектории.", _
        "Auto: UltimateStrain не сошелся, предельная точка найдена через LoadMultiplier."

    Set ExecuteAutoCapacityLoadMultiplierFallback = request.Problem.BuildResult(request, False)
End Function

' В Auto-режиме образования трещины по очереди пробует lambda*Mxy, lambda*N и
' lambda*NMxy. Если путь дал только резервный psi_s=1 без точки Ncrc/Mcrc, это
' не ошибка, но Auto продолжает искать более информативную траекторию.
Private Function ExecuteAutoCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Dim savedPath As String
    savedPath = NormalizeCrackPath(request.CrackFormationPath)

    Dim paths(1 To 3) As String
    paths(1) = CRACK_PATH_MXY
    paths(2) = CRACK_PATH_N
    paths(3) = CRACK_PATH_NMXY

    crack.LimitSearchAppendDiagnostic "Auto-путь: начинаем поиск образования нормальной трещины. Порядок попыток: " & _
        ChrW$(&H3BB) & "*Mxy -> " & ChrW$(&H3BB) & "*N -> " & ChrW$(&H3BB) & "*NMxy."

    Dim lastResult As CLimitSearchResult
    Dim lastStopReason As String
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        crack.LimitSearchBeginCrackFormationAttempt paths(i)
        crack.LimitSearchAppendDiagnostic "Auto-путь: попытка " & CStr(i) & " из " & _
            CStr(UBound(paths)) & " - пробуем " & _
            crack.LimitSearchCrackFormationPathDisplayName() & "."

        Set lastResult = ExecuteCrackFormationCurrentPath(request)
        If lastResult.Succeeded Then
            If crack.LimitSearchFormationFallbackOnly And i < UBound(paths) Then
                lastStopReason = crack.StopReason
                crack.LimitSearchAppendDiagnostic "Auto-путь: " & _
                    crack.LimitSearchCrackFormationPathDisplayName() & _
                    " дал только резервный результат psi_s=1 без найденной точки образования трещины. Для более информативного результата переходим к следующему пути."
            Else
                crack.LimitSearchSetResolvedCrackFormationPath paths(i)
                crack.LimitSearchAppendDiagnostic "Auto-путь: выбран " & _
                    crack.LimitSearchCrackFormationPathDisplayName() & "."
                crack.LimitSearchSetCurrentFormationPath savedPath
                Set ExecuteAutoCrackFormation = lastResult
                Exit Function
            End If
        Else
            lastStopReason = crack.StopReason
            crack.LimitSearchAppendDiagnostic "Auto-путь: " & _
                crack.LimitSearchCrackFormationPathDisplayName() & _
                " не дал точку образования трещины: " & crack.StopReason
        End If
    Next i

    crack.LimitSearchSetCurrentFormationPath savedPath
    If Not lastResult Is Nothing Then Set ExecuteAutoCrackFormation = lastResult
End Function

' Запускает выбранный метод поиска на уже выбранном пути образования трещины.
' Для Auto-стратегии сначала используется UltimateStrain, а при численной
' неудаче - LoadMultiplier по той же физической траектории.
Private Function ExecuteCrackFormationCurrentPath(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    Select Case NormalizeStrategy(request.SolutionStrategy)
        Case "LoadMultiplier"
            Dim multiplier As CLoadMultiplierSearch
            Set multiplier = New CLoadMultiplierSearch
            Set ExecuteCrackFormationCurrentPath = multiplier.ExecuteCrackFormation(request)
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
    If request.CrackCalculator.LimitSearchStoppedTerminally() Then
        Set ExecuteAutoStrategyCrackFormation = firstResult
        Exit Function
    End If

    Dim firstStopReason As String
    firstStopReason = request.CrackCalculator.StopReason
    request.CrackCalculator.LimitSearchAppendDiagnostic _
        "UltimateStrain не дал точку образования трещины: " & firstStopReason
    request.CrackCalculator.LimitSearchAppendDiagnostic _
        "Пробуем LoadMultiplier по тому же пути образования трещины."
    request.CrackCalculator.LimitSearchClearStopReason

    Dim multiplier As CLoadMultiplierSearch
    Set multiplier = New CLoadMultiplierSearch
    Dim secondResult As CLimitSearchResult
    Set secondResult = multiplier.ExecuteCrackFormation(request)
    If Not secondResult.Succeeded And Len(request.CrackCalculator.StopReason) = 0 Then _
        request.CrackCalculator.LimitSearchSetStopReasonIfEmpty firstStopReason
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

' Нормализует ключ пути образования трещины без обращения к Config-reader.
Private Function NormalizeCrackPath(ByVal value As String) As String
    Select Case UCase$(Trim$(value))
        Case "AUTO"
            NormalizeCrackPath = CRACK_PATH_AUTO
        Case "LAMBDAN", "LAMBDA*N", "N"
            NormalizeCrackPath = CRACK_PATH_N
        Case "LAMBDANMXY", "LAMBDA*NMXY", "NMXY", "ALL"
            NormalizeCrackPath = CRACK_PATH_NMXY
        Case "LAMBDAMXY", "LAMBDA*MXY", "MXY", vbNullString
            NormalizeCrackPath = CRACK_PATH_MXY
        Case Else
            NormalizeCrackPath = Trim$(value)
    End Select
End Function
