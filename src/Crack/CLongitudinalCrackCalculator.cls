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
' Проверка продольных трещин по готовому НДС
' ==========================================================================
' Получает текущее равновесное состояние с трещинами CurrentCrackedState,
' берёт максимальное сжатие бетона и формирует результат проверки продольных трещин.
' Если обязательное состояние непригодно или не получено, проверка отмечается
' как недоступная по исходным данным; новая численная несходимость не назначается.

Private Const STRESS_ZERO_TOLERANCE As Double = 0.000000001 ' МПа; ниже считаем, что сжатого бетона нет.

' Проверяет продольные трещины по сохранённому текущему НДС с трещинами CrackedState.
Public Function Calculate(ByVal currentState As CSectionStateResult, ByVal rbMc2 As Double) As CLongitudinalCrackResult
    If currentState Is Nothing Then
        Set Calculate = BlockedByDependency("CurrentCrackedState", _
            "Состояние НДС с трещинами не получено.", rbMc2)
        Exit Function
    End If
    If Not currentState.Converged Then
        Set Calculate = BlockedByDependency("CurrentCrackedState", _
            "Состояние НДС с трещинами не сошлось: " & currentState.StopReason, rbMc2)
        Exit Function
    End If
    If currentState.ExtensionUsed Or Not currentState.WithinPhysicalRange Then
        Set Calculate = BlockedByDependency("CurrentCrackedState", _
            "Состояние НДС с трещинами найдено вне физического диапазона диаграмм.", rbMc2)
        Exit Function
    End If

    Dim maxCompressionStress As Double
    maxCompressionStress = Abs(currentState.MinConcreteStress)

    Set Calculate = CalculateFromStress(maxCompressionStress, rbMc2)
End Function

' Проверяет переданное максимальное сжатие бетона в МПа без повторного решения НДС.
' Здесь определяются результат и причина проверки, а общие правила статусов
' только задают обозначение при выводе.
Public Function CalculateFromStress(ByVal maxCompressionStress As Double, ByVal rbMc2 As Double) As CLongitudinalCrackResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    Dim utilization As Double
    If maxCompressionStress <= STRESS_ZERO_TOLERANCE Then
        meta.SetNotApplicable rkLongitudinalCrack, "Сжатого бетона для проверки продольных трещин нет."
    ElseIf rbMc2 <= 0# Then
        meta.SetResult rsInvalidInput, rcInvalidInput, rkLongitudinalCrack, "Допускаемое напряжение бетона Rb,mc2 должно быть больше нуля."
    Else
        utilization = maxCompressionStress / rbMc2
        Dim reason As String
        reason = "Максимальное сжатие бетона " & FormatStress(maxCompressionStress) & _
            " МПа; допускаемое Rb,mc2 = " & FormatStress(rbMc2) & " МПа."
        If maxCompressionStress <= rbMc2 Then
            meta.SetResult rsSuccess, rcCheckPassed, rkLongitudinalCrack, "Проверка продольных трещин проходит. " & reason
        Else
            meta.SetResult rsCheckFailed, rcLongitudinalCrackExceeded, rkLongitudinalCrack, _
                "Проверка продольных трещин не проходит: сжатие бетона выше допускаемого. " & reason
        End If
    End If
    Dim result As CLongitudinalCrackResult
    Set result = New CLongitudinalCrackResult
    result.Initialize meta, maxCompressionStress, rbMc2, utilization
    Set CalculateFromStress = result
End Function

' Форматирует напряжение во внутренних единицах для самостоятельного
' инженерного комментария, не обращаясь к Excel или пользовательской палитре.
Private Function FormatStress(ByVal value As Double) As String
    FormatStress = Replace$(Format$(value, "0.000"), ",", ".")
End Function

' Создаёт результат недоступной проверки продольных трещин, если обязательное текущее НДС не
' получено.
Public Function BlockedByDependency(ByVal dependencyName As String, ByVal reason As String, _
        Optional ByVal rbMc2 As Double = 0#) As CLongitudinalCrackResult
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetBlockedByDependency rkLongitudinalCrack, dependencyName, reason
    Dim result As CLongitudinalCrackResult
    Set result = New CLongitudinalCrackResult
    result.Initialize meta, _
        0#, rbMc2, 0#
    Set BlockedByDependency = result
End Function
