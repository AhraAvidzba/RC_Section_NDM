VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialPointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Набор точек материалной диаграммы
' ==========================================================================
' Маленький контейнер передает результат работы builders в расчетные
' материалы. Он хранит только пары "деформация - напряжение" во внутренних
' единицах и не содержит нормативной логики.
'
' Такой слой нужен, чтобы CConcreteDiagramBuilder и CSteelDiagramBuilder
' могли формировать точки независимо от конкретного material-класса, а
' CConcreteDiagramMaterial/CSteelDiagramMaterial оставались простыми
' интерполяторами для solver-а.

Private mStrains() As Double ' Деформации узлов диаграммы.
Private mStresses() As Double ' Напряжения узлов диаграммы, МПа.
Private mCount As Long ' Количество добавленных точек.

' Добавляет одну точку диаграммы в конец контейнера.
Public Sub AddPoint(ByVal strain As Double, ByVal stress As Double)
    mCount = mCount + 1
    ReDim Preserve mStrains(1 To mCount)
    ReDim Preserve mStresses(1 To mCount)
    mStrains(mCount) = strain
    mStresses(mCount) = stress
End Sub

' Количество точек в контейнере.
Public Property Get Count() As Long
    Count = mCount
End Property

' Возвращает деформацию точки по индексу.
Public Function Strain(ByVal index As Long) As Double
    ValidateIndex index
    Strain = mStrains(index)
End Function

' Возвращает напряжение точки по индексу.
Public Function Stress(ByVal index As Long) As Double
    ValidateIndex index
    Stress = mStresses(index)
End Function

' Защищает material-классы от выхода за границы массива.
Private Sub ValidateIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then
        Err.Raise vbObjectError + 3180, "CMaterialPointSet", "Индекс точки материалной диаграммы вне диапазона."
    End If
End Sub
