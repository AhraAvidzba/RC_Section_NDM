VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISectionShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Параметрическое описание формы. Принадлежность точек, площадь, центр
' тяжести и сетка вычисляются по этим контурам общим геометрическим ядром.
' Новый тип задает наружные границы и любое число отверстий с явными ролями.
Public Function IsValid(ByRef message As String) As Boolean
End Function

Public Sub BuildContours(ByVal contours As CSectionContours)
End Sub
