Attribute VB_Name = "Interpolate"
Option Explicit
'функция интерполяции по одному параметру
'интерполяция: 0 - по горизонтали, 1 - по вертикали
Public Function InterRange(RangeX As Range, rangeValues As Range, X#, Gor0Vert1%) As Variant
  Dim i%
    InterRange = 0:   i = 0
        'f(x)=f1+(x-x1)*(f2-f1)/(x2-x1)
        'if (x2-x1)=0
    If Gor0Vert1 = 0 Then ' По горизонтали
        While X > RangeX.Offset(0, i).Value
            i = i + 1
        Wend
        If i = 0 Then
           InterRange = rangeValues.Offset(0, 0).Value
        ElseIf (RangeX.Offset(0, i).Value - RangeX.Offset(0, i - 1).Value) = 0 Then
           InterRange = rangeValues.Offset(0, 0).Value
        Else
           InterRange = rangeValues.Offset(0, i - 1).Value + (X - RangeX.Offset(0, i - 1).Value) * (rangeValues.Offset(0, i).Value - rangeValues.Offset(0, i - 1).Value) / (RangeX.Offset(0, i).Value - RangeX.Offset(0, i - 1).Value)
        End If
    ElseIf Gor0Vert1 = 1 Then ' По вертикали
        While X > RangeX.Offset(i, 0).Value
            i = i + 1
        Wend
        If i = 0 Then
          InterRange = rangeValues.Offset(0, 0).Value
        ElseIf (RangeX.Offset(i, 0).Value - RangeX.Offset(i - 1, 0).Value) = 0 Then
          InterRange = rangeValues.Offset(0, 0).Value
        Else
          InterRange = rangeValues.Offset(i - 1, 0).Value + (X - RangeX.Offset(i - 1, 0).Value) * (rangeValues.Offset(i, 0).Value - rangeValues.Offset(i - 1, 0).Value) / (RangeX.Offset(i, 0).Value - RangeX.Offset(i - 1, 0).Value)
        End If
    Else
      InterRange = "???"
    End If
End Function
'функция двойной интерполяции
Public Function AB_inter(Horiz!, Vertic!, Table As Range) As Single
   Dim pos As Range, stg#, a#
   Dim col_c%, col_r%
   Dim i%, j%, k%, L%
   Dim gor!, vert!, dxt!, dx!, dyt!, dy!, a11!, a12!, a21!, a22!, ah!, ahh!
   
   gor = Horiz
   vert = Vertic
   
   col_c = Table.Columns.count
   col_r = Table.Rows.count
   stg = -1
   If gor > Table(1, col_c) Then
      gor = Table(1, col_c)
   End If
   If vert > Table(col_r, 1) Then
      vert = Table(col_r, 1)
   End If
   a = 1
   For i = 2 To col_c
     If gor <= Table(1, i) Then
       If i <= 2 Then
         dxt = Table(1, i)
         If gor = 0 Then
           dx = 0
         Else
           dx = dxt - gor
         End If
       Else
         dxt = Table(1, i) - Table(1, i - 1)
         dx = Table(1, i) - gor
       End If
       j = i
       Exit For
     End If
   Next
   For i = 2 To col_r
     If vert <= Table(i, 1) Then
       If j <= 2 And i <= 2 Then
         a11 = 0
       Else
         a11 = Table(i - 1, j - 1)
       End If
       If i <= 2 Then
         a12 = 0
         a11 = 0
       Else
         a12 = Table(i - 1, j)
       End If
       If j <= 2 Then
        a21 = 0
        a11 = 0
       Else
        a21 = Table(i, j - 1)
       End If
       a22 = Table(i, j)
       dyt = Table(i, 1) - Table(i - 1, 1)
       dy = Table(i, 1)
       dy = dy - vert
       If dx = 0 Then dxt = 1
       ah = a12 - (a12 - a11) * dx / dxt
       ahh = a22 - (a22 - a21) * dx / dxt
       a = ahh - (ahh - ah) * dy / dyt
       Exit For
     End If
   Next
   AB_inter = a
End Function

