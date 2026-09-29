VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLongitudinalCrackCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Проверка продольных трещин по уже найденному CurrentCrackedState
' ==========================================================================
' Класс не запускает CSectionSolver и не выбирает диаграммы материалов. Он получает
' готовое равновесное состояние с трещинами, берет из него максимальное сжатие
' бетона и оформляет самостоятельный result-объект продольных трещин.
' Если обязательное состояние не получено или не может использоваться физически,
' проверка помечается как заблокированная зависимостью, а не как новая
' численная несходимость формульного этапа.

Private Const STRESS_ZERO_TOLERANCE As Double = 0.000000001 ' МПа; ниже считаем, что сжатого бетона нет.

' Выполняет проверку продольных трещин по сохраненному named-state CrackedState.
Public Function Calculate(ByVal currentState As CSectionStateResult, ByVal rbMc2 As Double, _
        ByVal statusPolicy As CBatchStatusPolicy) As CLongitudinalCrackResult
    Dim result As CLongitudinalCrackResult
    Set result = New CLongitudinalCrackResult

    Dim policy As CBatchStatusPolicy
    Set policy = EffectivePolicy(statusPolicy)

    If currentState Is Nothing Then
        Set Calculate = BlockedByDependency(policy, "CurrentCrackedState", _
            "Состояние НДС с трещинами не получено.", rbMc2)
        Exit Function
    End If
    If Not currentState.Converged Then
        Set Calculate = BlockedByDependency(policy, "CurrentCrackedState", _
            "Состояние НДС с трещинами не сошлось: " & currentState.StopReason, rbMc2)
        Exit Function
    End If
    If currentState.ExtensionUsed Or Not currentState.WithinPhysicalRange Then
        Set Calculate = BlockedByDependency(policy, "CurrentCrackedState", _
            "Состояние НДС с трещинами найдено вне физического диапазона диаграмм.", rbMc2)
        Exit Function
    End If

    Dim maxCompressionStress As Double
    maxCompressionStress = Abs(currentState.MinConcreteStress)

    Dim utilization As Double
    If maxCompressionStress > STRESS_ZERO_TOLERANCE And rbMc2 > 0# Then
        utilization = maxCompressionStress / rbMc2
    End If

    result.Initialize policy.LongitudinalCrackMeta(maxCompressionStress, rbMc2), _
        maxCompressionStress, rbMc2, utilization
    Set Calculate = result
End Function

' Возвращает result продольных трещин для случая, когда обязательное НДС отсутствует.
Public Function BlockedByDependency(ByVal statusPolicy As CBatchStatusPolicy, _
        ByVal dependencyName As String, ByVal reason As String, _
        Optional ByVal rbMc2 As Double = 0#) As CLongitudinalCrackResult
    Dim policy As CBatchStatusPolicy
    Set policy = EffectivePolicy(statusPolicy)

    Dim result As CLongitudinalCrackResult
    Set result = New CLongitudinalCrackResult
    result.Initialize policy.BlockedByDependencyMeta(rkLongitudinalCrack, dependencyName, reason), _
        0#, rbMc2, 0#
    Set BlockedByDependency = result
End Function

' Создает локальную policy только для защитных сценариев тестов; в рабочем batch она передается снаружи.
Private Function EffectivePolicy(ByVal statusPolicy As CBatchStatusPolicy) As CBatchStatusPolicy
    If statusPolicy Is Nothing Then
        Set EffectivePolicy = New CBatchStatusPolicy
    Else
        Set EffectivePolicy = statusPolicy
    End If
End Function
