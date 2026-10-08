Attribute VB_Name = "modRebarLayoutGeometry"
Option Explicit

' Общая геометрия раскладки арматуры: шаг по траектории и дополнительные ряды.
' Не знает типа сечения, Config, материалов и аннотаций.

' Расстояние между осями учитывает оба радиуса и пропущенный промежуточный ряд.
Public Sub RebarOffsetAdditionalRow(ByVal baseX As Double, ByVal baseY As Double, _
        ByVal row1Diameter As Double, ByVal targetDiameter As Double, ByVal skippedDiameter As Double, _
        ByVal locationMode As String, ByVal tangentX As Double, ByVal tangentY As Double, _
        ByVal inwardX As Double, ByVal inwardY As Double, ByRef xCoord As Double, ByRef yCoord As Double)
    Dim distance As Double
    distance = row1Diameter / 2# + skippedDiameter + targetDiameter / 2#
    If StrComp(locationMode, "SideBySide", vbTextCompare) = 0 Then
        xCoord = baseX + tangentX * distance
        yCoord = baseY + tangentY * distance
    Else
        xCoord = baseX + inwardX * distance
        yCoord = baseY + inwardY * distance
    End If
End Sub

' EverySecondBar отсчитывается от начала своей грани, а не от индекса layout.
Public Function RebarShouldAddAdditionalRow(ByVal baseBarOrdinal As Long, ByVal bindingMode As String) As Boolean
    If StrComp(bindingMode, "EverySecondBar", vbTextCompare) = 0 Then
        RebarShouldAddAdditionalRow = ((baseBarOrdinal Mod 2) = 1)
    Else
        RebarShouldAddAdditionalRow = True
    End If
End Function

Public Function RebarRowLocationIsValid(ByVal value As String) As Boolean
    RebarRowLocationIsValid = (StrComp(value, "SideBySide", vbTextCompare) = 0 Or StrComp(value, "Stacked", vbTextCompare) = 0)
End Function

Public Function RebarRowBindingIsValid(ByVal value As String) As Boolean
    RebarRowBindingIsValid = (StrComp(value, "EachBar", vbTextCompare) = 0 Or StrComp(value, "EverySecondBar", vbTextCompare) = 0)
End Function

' Длина именно траектории, не хорды между ее концами.
Public Function RebarPolylineLength(ByRef px() As Double, ByRef py() As Double) As Double
    Dim i As Long
    For i = LBound(px) To UBound(px) - 1
        RebarPolylineLength = RebarPolylineLength + Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
    Next i
End Function

' Равномерный шаг по ломаной сохраняет прежние координаты и локальную касательную.
Public Sub RebarPointAtDistance(ByRef px() As Double, ByRef py() As Double, ByVal targetDistance As Double, _
        ByRef xCoord As Double, ByRef yCoord As Double, ByRef tangentX As Double, ByRef tangentY As Double)
    Dim accumulated As Double, i As Long, segLen As Double, t As Double
    For i = LBound(px) To UBound(px) - 1
        segLen = Sqr((px(i + 1) - px(i)) ^ 2 + (py(i + 1) - py(i)) ^ 2)
        If segLen <= GEOM_TOLERANCE Then GoTo NextSegment
        If accumulated + segLen >= targetDistance - GEOM_TOLERANCE Then
            t = (targetDistance - accumulated) / segLen
            If t < 0# Then t = 0#
            If t > 1# Then t = 1#
            xCoord = px(i) + t * (px(i + 1) - px(i))
            yCoord = py(i) + t * (py(i + 1) - py(i))
            tangentX = (px(i + 1) - px(i)) / segLen
            tangentY = (py(i + 1) - py(i)) / segLen
            Exit Sub
        End If
        accumulated = accumulated + segLen
NextSegment:
    Next i
    xCoord = px(UBound(px)): yCoord = py(UBound(py))
    tangentX = 0#: tangentY = -1#
End Sub

Public Sub RebarNormalizeVector(ByRef xValue As Double, ByRef yValue As Double)
    Dim length As Double
    length = Sqr(xValue * xValue + yValue * yValue)
    If length <= GEOM_TOLERANCE Then Exit Sub
    xValue = xValue / length: yValue = yValue / length
End Sub
