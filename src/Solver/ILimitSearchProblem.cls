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

' Возвращает исходный доменный объект только для тех общих алгоритмов, где
' исторически уже используется callback-набор UltimateStrain.
Public Property Get DomainContext() As Object
End Property

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

' Возвращает имя одномерного метода: Bisection, Brent или Secant.
Public Property Get SearchMethod() As String
End Property

' True, если после неудачи bracket-а можно принять удержанную чисто осевую
' несущую способность. Для CrackFormation всегда False.
Public Property Get UsesRetainedForceCapacity() As Boolean
End Property

' Выполняет доменную подготовку перед LoadMultiplier.
Public Function Prepare(ByVal request As CLimitSearchRequest) As Boolean
End Function

' Создает lambda-траекторию для выбранной инженерной задачи.
Public Function BuildLoadPath(ByVal request As CLimitSearchRequest) As CLoadPathVector
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

' Добавляет строку диагностики в доменный журнал.
Public Sub AppendDiagnostic(ByVal message As String)
End Sub

' Финализирует найденную lambda-точку.
Public Sub FinalizeAt(ByVal request As CLimitSearchRequest, _
        ByVal loadPath As CLoadPathVector, ByVal lambdaValue As Double, _
        ByVal probeState As String, ByVal methodName As String, _
        ByVal solver As CSectionSolver, ByVal iterations As Long)
End Sub

' Принимает удержанную чисто осевую несущую способность, если это применимо.
Public Sub AcceptRetainedCapacity(ByVal request As CLimitSearchRequest)
End Sub

' Собирает общий snapshot результата search.
Public Function BuildResult(ByVal request As CLimitSearchRequest, _
        ByVal searchSucceeded As Boolean) As CLimitSearchResult
End Function

' Ниже идут методы, нужные CrackFormation Auto. Для capacity-адаптера они
' являются безопасными no-op, чтобы coordinator оставался независимым от класса.
Public Sub BeginFormationAttempt(ByVal pathName As String)
End Sub

Public Sub SetResolvedFormationPath(ByVal pathName As String)
End Sub

Public Property Get FormationFallbackOnly() As Boolean
End Property

Public Property Get StoppedTerminally() As Boolean
End Property

Public Property Get StopReason() As String
End Property

Public Sub ClearStopReason()
End Sub

Public Sub SetStopReasonIfEmpty(ByVal reason As String)
End Sub

Public Function FormationPathDisplayName() As String
End Function
