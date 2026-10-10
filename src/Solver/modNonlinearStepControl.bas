Attribute VB_Name = "modNonlinearStepControl"
Option Explicit

' Общая автоматическая приемка нелинейного шага; допуски принадлежат задачам.
' Достаточное снижение Armijo: SUNDIALS KINSOL, Mathematical Considerations 7.2.5.
Public Const NONLINEAR_MAX_BACKTRACKS As Long = 53
Private Const ARMIJO_FACTOR As Double = 0.0001
Private Const SIGNIFICANT_REDUCTION As Double = 0.000000000000007105427357601

' Снижение должно отличаться от шума Double (32 машинных epsilon).
' Это не заменяет и не ослабляет допуски равновесия задачи.
' Малая поправка сама по себе никогда не подтверждает равновесие.
Public Function NonlinearStepImproves(ByVal baseNorm As Double, ByVal trialNorm As Double, _
        ByVal alpha As Double) As Boolean
    Dim reduction As Double
    reduction = ARMIJO_FACTOR * alpha
    If reduction < SIGNIFICANT_REDUCTION Then reduction = SIGNIFICANT_REDUCTION
    NonlinearStepImproves = trialNorm < baseNorm And _
        trialNorm <= baseNorm * (1# - reduction)
End Function

' Ближайшие два перелома диаграмм вдоль плоскости деформаций. В стержнях
' учитывается и вытесненный бетон, поскольку его напряжение вычитается из стали.
Public Function NonlinearMaterialTransitions(ByVal section As CSectionModel, _
        ByVal frame As CSectionNumericalFrame, ByVal concrete As CMaterialDiagram, _
        ByVal steel As CMaterialDiagram, ByRef q() As Double, ByRef direction() As Double, _
        ByRef first As Double, ByRef second As Double) As Boolean
    first = 0#: second = 0#
    Dim i As Long, x As Double, y As Double, strain As Double, delta As Double
    For i = 1 To section.ConcreteCount + section.RebarCount
        If i <= section.ConcreteCount Then
            x = section.ConcreteX(i): y = section.ConcreteY(i)
        Else
            x = section.RebarX(i - section.ConcreteCount): y = section.RebarY(i - section.ConcreteCount)
        End If
        x = (x - frame.OriginX) / frame.LengthScale
        y = (y - frame.OriginY) / frame.LengthScale
        strain = q(1) + q(2) * y + q(3) * x
        delta = direction(1) + direction(2) * y + direction(3) * x
        If delta <> 0# Then
            CollectMaterialTransitions concrete, strain, delta, first, second
            If i > section.ConcreteCount Then CollectMaterialTransitions steel, strain, delta, first, second
        End If
    Next i
    NonlinearMaterialTransitions = first > 0#
End Function

' Узлы с одинаковым наклоном не меняют работу материала и не ограничивают шаг.
' Крайние узлы тоже важны: за ними касательный модуль равен нулю.
Private Sub CollectMaterialTransitions(ByVal diagram As CMaterialDiagram, ByVal strain As Double, _
        ByVal delta As Double, ByRef first As Double, ByRef second As Double)
    Dim i As Long, leftSlope As Double, rightSlope As Double, node As Double, distance As Double
    For i = 1 To diagram.PointCount
        node = diagram.PointStrain(i)
        leftSlope = 0#: rightSlope = 0#
        If i > 1 Then leftSlope = (diagram.PointStress(i) - diagram.PointStress(i - 1)) / _
            (node - diagram.PointStrain(i - 1))
        If i < diagram.PointCount Then rightSlope = (diagram.PointStress(i + 1) - diagram.PointStress(i)) / _
            (diagram.PointStrain(i + 1) - node)
        If leftSlope <> rightSlope Then
            distance = (node - strain) / delta
            If distance > 0# And strain + distance * delta <> strain Then
                If first = 0# Or distance < first Then
                    second = first: first = distance
                ElseIf distance > first Then
                    If second = 0# Or distance < second Then second = distance
                End If
            End If
        End If
    Next i
End Sub
