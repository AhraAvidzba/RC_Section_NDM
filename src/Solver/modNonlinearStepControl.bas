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
    If reduction < SIGNIFICANT_REDUCTION Then reduction = SIGNIFICANT_REDUCTION ' Снижение должно различаться в Double даже при малой alpha; это критерий шага, а не допуска равновесия.
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
            concrete.CollectStrainTransitions strain, delta, first, second
            If i > section.ConcreteCount Then steel.CollectStrainTransitions strain, delta, first, second
        End If
    Next i
    NonlinearMaterialTransitions = first > 0#
End Function
