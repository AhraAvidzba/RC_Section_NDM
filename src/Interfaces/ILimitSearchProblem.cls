VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILimitSearchProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Контракт доменной задачи для общего LimitSearch
' ==========================================================================
' Интерфейс отделяет одномерный численный search от инженерной задачи.
' Capacity и CrackFormation предоставляют одинаковые операции: подготовить путь,
' оценить одну lambda-точку, классифицировать пробу и принять итог. Сам search
' после этого не знает, какой доменный калькулятор стоит за callback-ом.

' Возвращает тип результата, для которого выполняется search.
Public Property Get ResultKind() As EResultKind
End Property

' Возвращает начальный шаг расширения lambda-скобки.
Public Property Get InitialLambdaStep() As Double
End Property

' Возвращает техническую верхнюю границу lambda для поиска.
Public Property Get MaxLambda() As Double
End Property

' Возвращает требуемую точность lambda-скобки.
Public Property Get LambdaTolerance() As Double
End Property

' Возвращает предельное число итераций одномерного search.
Public Property Get MaxIterations() As Long
End Property

' Отдельный лимит уточнения скобки Bisection; 0 означает уточнение до допуска.
' Он не обязан совпадать с лимитом Newton внутренней equilibrium-пробы.
Public Property Get BisectionMaxIterations() As Long
End Property

' Возвращает имя одномерного метода: Bisection, Brent или Secant.
Public Property Get SearchMethod() As String
End Property

Public Property Get Converged() As Boolean
End Property

Public Property Get LimitState() As String
End Property

Public Property Get DiagnosticLog() As String
End Property

' Готовит инженерный контекст нового конкретного пути общего Auto-перебора.
' Численные состояния предыдущей search-session не переиспользуются как
' физические состояния нового пути; общая диагностика может сохраняться.
Public Sub BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' Сохраняет фактический путь и причины предыдущих попыток у владельца результата.
' Writer не склеивает эти комментарии и не определяет по ним машинный статус.
Public Sub FinishLoadPath(ByVal request As CLimitSearchRequest, ByVal previousPathComments As String)
End Sub

' Выполняет доменную подготовку перед LoadMultiplier.
Public Function Prepare(ByVal request As CLimitSearchRequest) As Boolean
End Function

' Создает lambda-траекторию для выбранной инженерной задачи.
Public Function BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
End Function

' Решает, можно ли продолжить поиск от подтвержденной начальной точки.
' Отказ или физический результат без точки оформляет владелец задачи.
Public Function AcceptInitialPoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal functionValue As Double, _
        ByVal probeState As String, ByVal solver As CSectionSolver) As Boolean
End Function

' Оценивает одну lambda-точку и возвращает значение функции search, код probe
' и, если нужно, solver найденного состояния.
Public Function EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
End Function

' Проверяет, что probe еще находится на допустимой стороне предела.
Public Function IsAcceptableState(ByVal probeState As String) As Boolean
End Function

' Проверяет, что probe достиг физического критерия предела.
Public Function IsLimitState(ByVal probeState As String) As Boolean
End Function

' Проверяет, что probe завершился численной неудачей.
Public Function IsNumericalFailureState(ByVal probeState As String) As Boolean
End Function

' Отличает отказ входа от численной неудачи, которую разрешено повторять.
' Domain adapter классифицирует свой probe-код, общий Search не разбирает текст.
Public Function IsInvalidInputState(ByVal probeState As String) As Boolean
End Function

' Предлагает промежуточную lambda после несошедшейся верхней пробы.
' Общий Search отдельно проверяет границы и строгий прогресс recovery.
Public Function ReduceLambdaAfterNumericalFailure(ByVal lowLambda As Double, _
        ByVal highLambda As Double, ByVal probeState As String) As Double
End Function

' Запоминает последнюю допустимую lambda-точку.
Public Sub RecordAcceptable(ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double)
End Sub

' Обрабатывает отказ начальной точки lambda = 0.
Public Sub HandleInitialFailure(ByVal probeState As String)
End Sub

' Обрабатывает отсутствие верхней границы физического предела.
Public Sub HandleBracketNotFound(ByVal lowLambda As Double)
End Sub

' Обрабатывает численную неудачу внутри найденной скобки.
Public Sub HandleBracketNumericalFailure(ByVal lowLambda As Double, _
        ByVal failedLambda As Double, ByVal probeState As String)
End Sub

' Обрабатывает исчерпание итераций до достижения заданной точности.
Public Sub HandleIterationLimit(ByVal maxIterations As Long, _
        ByVal remainingWidth As Double, ByVal tolerance As Double)
End Sub

' Обрабатывает неизвестный probe-код как нарушение контракта callback-а.
Public Sub HandleUnexpectedState(ByVal probeState As String)
End Sub

' Фиксирует ошибку настройки или исходных данных.
Public Sub SetInvalidInput(ByVal message As String)
End Sub

' Фиксирует численную неудачу search.
Public Sub SetNumericalFailure(ByVal message As String)
End Sub

' Сохраняет точную техническую причину общего solve/search без разбора текста.
' Ошибка настройки или внутреннего контракта не становится численной неудачей.
Public Sub SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
End Sub

' Добавляет строку диагностики в доменный журнал.
Public Sub AppendDiagnostic(ByVal message As String)
End Sub

' Финализирует найденную lambda-точку.
Public Function FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
End Function

' Принимает уточненную скобку по физическому критерию конкретной задачи.
' Выбор допустимой стороны и дополнительное подтверждение не принадлежат Search.
Public Sub FinalizeBracket(ByVal request As CLimitSearchRequest, ByVal loadPath As CLoadPathVector, _
        ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long)
End Sub

' Пробует специальный подтвержденный метод после численной неудачи поиска.
' Если такого метода у задачи нет, ее первичная причина сохраняется.
Public Sub TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
End Sub

' Собирает общий snapshot результата search.
Public Function BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
End Function

' Проверяет постановку перед общим Newton-поиском; численный цикл здесь
' не реализуется, а material/criterion-настройки остаются у владельца задачи.
Public Function PrepareUltimate(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector) As Boolean
End Function

' Подбирает последовательность численных стартов одного инженерного поиска.
' Каждая попытка должна использовать общий UltimateStrain Newton-цикл.
Public Function RunUltimateAttempts(ByVal request As CLimitSearchRequest) As Boolean
End Function

' Завершает принятую попытку без повторного решения найденной точки.
Public Sub CompleteUltimate(ByVal loadPath As CLoadPathVector)
End Sub

' Строит стартовую плоскость для общей системы путь + физический критерий.
Public Sub LimitSearchInitializeUltimateGuess(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
End Sub

' Оценивает две невязки выбранного пути и невязку физического критерия.
' Здесь допустима оценка плоскости, но не фиктивное подтверждение равновесия.
Public Function LimitSearchEvaluateUltimateResidual(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal eps0 As Double, _
        ByVal kx As Double, ByVal ky As Double, ByRef solver As CSectionSolver, _
        ByRef rPath1 As Double, ByRef rPath2 As Double, ByRef rLimit As Double, _
        ByRef lambdaValue As Double, ByRef state As String) As Boolean
End Function

' Проверяет допуски критерия и задает согласованную норму для line search.
Public Function LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
End Function

' Нормирует те же три невязки для сравнения line-search шагов.
' Уменьшение нормы не заменяет отдельную проверку всех допусков и физического финала.
Public Function LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
End Function

' Принимает точку только после проверки ее усилий, load path и критерия.
Public Function LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As Boolean
End Function

' Сохраняет реальную численную неудачу общего Newton-цикла.
Public Sub LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
End Sub

' Настройки одного Newton-цикла: лимит, демпфирование и ограничения шага.
Public Property Get LimitSearchUltimateMaxIterations() As Long
End Property

Public Property Get LimitSearchUltimateDampingInitial() As Double
End Property

Public Property Get LimitSearchUltimateMinLineSearchAlpha() As Double
End Property

Public Property Get LimitSearchUltimateLineSearchEnabled() As Boolean
End Property

Public Property Get LimitSearchUltimateMaxDeltaEpsilon0() As Double
End Property

Public Property Get LimitSearchUltimateMaxDeltaKappa() As Double
End Property

' Сообщает численному слою об окончательном исходе задачи, который нельзя
' подменять последующей попыткой или общей численной ошибкой.
Public Property Get StoppedTerminally() As Boolean
End Property

Public Property Get StopReason() As String
End Property

' Снимает причину законченной попытки перед разрешенным резервным стартом.
' Не должен отменять typed терминальную ошибку или принятый физический исход.
Public Sub ClearStopReason()
End Sub

' Добавляет контекст остановки только при отсутствии более точной причины.
' Первичная диагностика solver/search сохраняет приоритет.
Public Sub SetStopReasonIfEmpty(ByVal reason As String)
End Sub

' Объединяет диагностику первой и резервной попыток в результирующем контексте.
' Последняя принятая попытка определяет исход; история объясняет переход метода.
' Факт первой численной попытки сохраняется явно и не восстанавливается
' по наличию/тексту журнала, который может быть отключен настройкой Diagnostics.
Public Sub MergeFallback(ByVal firstLimitState As String, ByVal firstStopReason As String, _
        ByVal firstDiagnostics As String, ByVal firstAttemptMessage As String, _
        ByVal transitionMessage As String, ByVal successMessage As String, _
        Optional ByVal firstSearchExecuted As Boolean = False)
End Sub
