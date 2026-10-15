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
' Интерфейс инженерной задачи для общего поиска предела
' ==========================================================================
' Задаёт операции подготовки нагрузки, оценки пробной точки, проверки
' критерия и приёма результата. Расчёты несущей способности и образования
' трещины предоставляют эти операции общим численным алгоритмам поиска.

' Возвращает вид инженерной проверки, для которой ищется предел.
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

' Возвращает максимальное число итераций одномерного поиска нагрузки.
Public Property Get MaxIterations() As Long
End Property

' Отдельный лимит уточнения скобки Bisection; 0 означает уточнение до допуска.
' Он не обязан совпадать с лимитом Newton внутренней equilibrium-пробы.
Public Property Get BisectionMaxIterations() As Long
End Property

' Возвращает имя одномерного метода: Bisection, Brent или Secant.
Public Property Get SearchMethod() As String
End Property

' Признак достигнутого численного равновесия; не заменяет проверку физических пределов.
Public Property Get Converged() As Boolean
End Property

' Признак достигнутого физического предела по проверяемому критерию.
Public Property Get LimitState() As String
End Property

' Накопленный диагностический журнал текущего расчёта или поиска.
Public Property Get DiagnosticLog() As String
End Property

' Готовит задачу для нового способа изменения нагрузки при переборе Auto.
' Пробы предыдущего пути не принимаются за физические состояния нового;
' общая история диагностики может сохраняться.
Public Sub BeginLoadPath(ByVal request As CLimitSearchRequest)
End Sub

' Сохраняет выбранный путь и причины прежних попыток в результате задачи.
' При записи на лист пояснения не собираются заново и не определяют код статуса.
Public Sub FinishLoadPath(ByVal request As CLimitSearchRequest, ByVal previousPathComments As String)
End Sub

' Готовит материалы, критерий и рабочие данные задачи перед LoadMultiplier.
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

' Оценивает пробную нагрузку lambda и возвращает значение критерия,
' код исхода и при необходимости найденное НДС.
Public Function EvaluatePoint(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal startSolver As CSectionSolver, _
        ByVal lambdaValue As Double, ByRef functionValue As Double, _
        ByRef probeState As String, ByRef solver As CSectionSolver) As Boolean
End Function

' Проверяет, находится ли пробная нагрузка с допустимой стороны предела.
Public Function IsAcceptableState(ByVal probeState As String) As Boolean
End Function

' Проверяет, достигнут ли физический предельный критерий в пробной нагрузке.
Public Function IsLimitState(ByVal probeState As String) As Boolean
End Function

' Проверяет, не удалось ли численно получить НДС пробной нагрузки.
Public Function IsNumericalFailureState(ByVal probeState As String) As Boolean
End Function

' Различает ошибку ввода и численную неудачу, после которой допустим повтор.
' Код исхода определяет сама инженерная задача; общий поиск текст причины не разбирает.
Public Function IsInvalidInputState(ByVal probeState As String) As Boolean
End Function

' Предлагает уменьшенное lambda после неудачной верхней пробы.
' Общий поиск проверяет границы и действительное изменение следующей нагрузки.
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

' Сохраняет неизвестный код исхода пробы как ошибку программного ответа задачи.
Public Sub HandleUnexpectedState(ByVal probeState As String)
End Sub

' Фиксирует ошибку настройки или исходных данных.
Public Sub SetInvalidInput(ByVal message As String)
End Sub

' Сохраняет численную неудачу общего поиска предельной нагрузки.
Public Sub SetNumericalFailure(ByVal message As String)
End Sub

' Сохраняет фактический код причины решения НДС или поиска предела.
' Ошибка настройки и внутренняя ошибка программы не становятся численной несходимостью.
Public Sub SetFailure(ByVal failureCode As ESolverFailureCode, ByVal reason As String)
End Sub

' Добавляет сообщение общего поиска в журнал инженерной задачи.
Public Sub AppendDiagnostic(ByVal message As String)
End Sub

' Финализирует найденную lambda-точку.
Public Function FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long) As Boolean
End Function

' Подтверждает одну из сторон уточнённого интервала по физическому критерию задачи.
' Задача выбирает сторону и дополнительные проверки; True означает
' подтверждённую точку, а не завершение последующих расчётных этапов.
Public Function FinalizeBracket(ByVal request As CLimitSearchRequest, ByVal loadPath As CLoadPathVector, _
        ByVal lowLambda As Double, ByVal lowSolver As CSectionSolver, _
        ByVal highLambda As Double, ByVal highSolver As CSectionSolver, _
        ByVal methodName As String, ByVal iterations As Long) As Boolean
End Function

' Пробует специальный подтвержденный метод после численной неудачи поиска.
' Если такого метода у задачи нет, ее первичная причина сохраняется.
Public Sub TryFinalizeAfterSearchFailure(ByVal request As CLimitSearchRequest)
End Sub

' Собирает готовый результат общего поиска предельной нагрузки.
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
        ByRef lambdaValue As Double, ByRef state As String, Optional ByVal centeredPlane As Boolean = False) As Boolean
End Function

' Проверяет допуски критерия и определяет общую меру невязок для подбора шага.
Public Function LimitSearchUltimateResidualReached(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Boolean
End Function

' Нормирует три невязки для сравнения пробных шагов. Их уменьшение
' не заменяет окончательную проверку всех допусков и физических пределов.
Public Function LimitSearchUltimateResidualNorm(ByVal rPath1 As Double, _
        ByVal rPath2 As Double, ByVal rLimit As Double) As Double
End Function

' Определяет масштаб общей меры невязок для подбора шага.
' Допуски окончательного подтверждения точки задаются независимо от него.
Public Property Get LimitSearchUseJacobianMeritScaling() As Boolean
End Property

' Принимает точку после проверки усилий, load path и критерия. Отказ точки
' завершает эту попытку, но не запрещает другой разрешенный старт или путь.
Public Function LimitSearchFinalizeUltimateResult(ByVal section As CSectionModel, _
        ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal solver As CSectionSolver, ByVal state As String, ByVal iterations As Long) As EUltimateFinalization
End Function

' Сохраняет реальную численную неудачу общего Newton-цикла.
Public Sub LimitSearchSetUltimateNumericalFailure(ByVal reason As String)
End Sub

' Настройки одного Newton-цикла: лимит, демпфирование и ограничения шага.
Public Property Get LimitSearchUltimateMaxIterations() As Long
End Property

' Сообщает численному слою об окончательном исходе задачи, который нельзя
' подменять последующей попыткой или общей численной ошибкой.
Public Property Get StoppedTerminally() As Boolean
End Property

' Фактическая причина завершения расчёта или поиска.
Public Property Get StopReason() As String
End Property

' Очищает причину завершённой попытки перед разрешённым дополнительным стартом.
' Окончательная ошибка или подтверждённый физический исход не отменяются.
Public Sub ClearStopReason()
End Sub

' Добавляет общее пояснение остановки только при отсутствии более точной причины
' решения НДС или поиска предела.
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
