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
' Выбор и запуск общего поиска предельной нагрузки
' ==========================================================================
' Выбирает UltimateStrain или LoadMultiplier для заданного изменения нагрузки.
' Несущая способность и образование трещины используют один механизм выбора;
' различия задаются критериями и материалами соответствующей задачи.
' Готовые материалы, сечение и настройки передаёт вызывающий расчёт.


' Выполняет поиск несущей способности по стратегии из request.
' Возвращает объект результата даже при неизвестной стратегии, чтобы вызывающий
' слой мог единообразно обработать ошибку настройки.
Public Function ExecuteCapacity(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4230, "CLimitSearchCoordinator", "Не переданы исходные данные для поиска предельного состояния."
    If request.HasLoadPathSelection Then
        Set ExecuteCapacity = ExecuteLoadPaths(request)
    Else
        Set ExecuteCapacity = ExecuteCapacityCurrentPath(request)
    End If
End Function

' Запускает выбранный численный метод несущей способности на одной готовой траектории нагрузки.
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

' Запускает образование трещины через общий выбор численного метода.
' Критерий eps_bt,ult, материалы и физический смысл нагрузки
' задаёт CCrackFormationCalculator.
Public Function ExecuteCrackFormation(ByVal request As CLimitSearchRequest) As CLimitSearchResult
    If request Is Nothing Then Err.Raise vbObjectError + 4231, "CLimitSearchCoordinator", "Не переданы исходные данные для поиска предельного состояния."
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
    If Len(pathKey) = 0 Then Err.Raise vbObjectError + 4233, "CLimitSearchCoordinator", "Не распознан способ изменения нагрузки (LoadPath). Выберите Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy."
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

' Последовательно пробует разрешённые способы изменения нагрузки. Подтверждённый
' физический предел принимается и при lambda < 1, когда текущее сочетание
' не проходит проверку. Ошибки ввода, настроек и программы сменой пути не скрываются.
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
            ' Сохраняем журналы предыдущих путей независимо от того, очищает ли их
            ' подготовка следующей задачи. Уже накопленный текст не добавляется повторно.
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

' В режиме Auto для чисто осевого изменения нагрузки сразу выбирает LoadMultiplier.
' Для остальных сначала пробует UltimateStrain, затем при неудаче LoadMultiplier.
' Оба способа используют общие численные классы; CCapacitySolver задаёт критерий и оценку
' нагрузки.
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
    ' Подтвержденный физический отказ при lambda=0 не требует другого
    ' численного метода. Выбор следующего Auto-пути остается прежним.
    If firstResult.Meta.ResultCode = rcInitialStateBeyondLimit Then
        Set ExecuteAutoCapacity = firstResult
        Exit Function
    End If
    If firstResult.Succeeded Then
        Set ExecuteAutoCapacity = firstResult
        Exit Function
    End If
    If request.Problem.StoppedTerminally Then
        Set ExecuteAutoCapacity = firstResult
        Exit Function
    End If

    Set ExecuteAutoCapacity = ExecuteAutoCapacityLoadMultiplierFallback(request, firstResult.SearchExecuted)
End Function

' После неудачи UltimateStrain повторяет поиск несущей способности через LoadMultiplier.
' Перед очисткой рабочих данных сохраняет диагностику первой попытки
' и возвращает обновлённый общий результат.
Private Function ExecuteAutoCapacityLoadMultiplierFallback(ByVal request As CLimitSearchRequest, _
        ByVal firstSearchExecuted As Boolean) As CLimitSearchResult
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
        "Auto: UltimateStrain не сошелся, предельная точка найдена через LoadMultiplier.", firstSearchExecuted

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
            ' При масштабировании только N сохраняем путь с его постоянными
            ' моментами. Отрицательная lambda не доказывает NotCracked:
            ' первая физическая граница подтверждается общим LoadMultiplier.
            Dim forcePathFallback As Boolean
            forcePathFallback = request.ForceOnly
            If Not request.SelectedLoadPath Is Nothing Then
                forcePathFallback = request.SelectedLoadPath.ScalesN And _
                    Not request.SelectedLoadPath.ScalesMx And Not request.SelectedLoadPath.ScalesMy
            End If
            Set ExecuteCrackFormationCurrentPath = ExecuteUltimateStrategyCrackFormation(request, forcePathFallback)
        Case "Auto"
            Set ExecuteCrackFormationCurrentPath = ExecuteUltimateStrategyCrackFormation(request, True)
        Case Else
            Dim invalidResult As CLimitSearchResult
            Set invalidResult = New CLimitSearchResult
            invalidResult.InitializeInvalidConfiguration request.SolutionStrategy, _
                "SLS.Crack.InitiationSolutionStrategy должен быть Auto, LoadMultiplier или UltimateStrain.", _
                rkCrackFormation
            Set ExecuteCrackFormationCurrentPath = invalidResult
    End Select
End Function

' На одном пути образования трещины сначала пробует UltimateStrain,
' затем разрешённый LoadMultiplier. В Auto переход разрешён для всех путей;
' при явном UltimateStrain — когда масштабируется N без пользовательских моментов.
Private Function ExecuteUltimateStrategyCrackFormation(ByVal request As CLimitSearchRequest, _
        ByVal allowMultiplierFallback As Boolean) As CLimitSearchResult
    Dim ultimate As CUltimateStrainSearch
    Set ultimate = New CUltimateStrainSearch

    Dim firstResult As CLimitSearchResult
    Set firstResult = ultimate.ExecuteCrackFormation(request)
    If firstResult.Succeeded Then
        Set ExecuteUltimateStrategyCrackFormation = firstResult
        Exit Function
    End If
    If request.Problem.StoppedTerminally Or Not allowMultiplierFallback Then
        Set ExecuteUltimateStrategyCrackFormation = firstResult
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
    Set ExecuteUltimateStrategyCrackFormation = secondResult
End Function

' Проверяет написание стратегии, заданное пользователем, независимо от чтения Config.
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
