VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CDiscreteForceIntegrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Суммирование усилий исходных дискретных волокон по линейным ветвям
' ==========================================================================
' Применяется только к пробам усилий без касательной и физической диагностики.
' Сетку и закон материала не меняет. При однородной ветви использует шесть
' моментов тех же площадей; неоднородные листья вычисляет штатной диаграммой.
' Физические экстремумы, критические элементы и матрицу K собирает прямой
' проход CSectionSolver. Индекс принадлежит одной модели и не удерживает её ссылку.
Private Type TForceNode
    First As Long ' Начало диапазона индексов исходных волокон.
    Count As Long ' Число волокон узла.
    Left As Long ' Левый дочерний узел либо ноль для листа.
    Right As Long ' Правый дочерний узел.
    MinX As Double ' Минимальная центрированная X.
    MaxX As Double ' Максимальная центрированная X.
    MinY As Double ' Минимальная центрированная Y.
    MaxY As Double ' Максимальная центрированная Y.
    A As Double ' Сумма площадей исходных волокон.
    AX As Double ' Сумма A*x.
    AY As Double ' Сумма A*y.
    AXX As Double ' Сумма A*x*x.
    AXY As Double ' Сумма A*x*y.
    AYY As Double ' Сумма A*y*y.
End Type
Private mFiberX() As Double, mFiberY() As Double, mArea() As Double ' Неизменяемые центрированные координаты и реальные площади.
Private mOrder() As Long, mStack() As Long ' Индексы исходных волокон и стек одной оценки.
Private mNodes() As TForceNode, mNodeCount As Long ' Пространственное дерево и фактическое число узлов.
Private mRevision As Long, mOriginX As Double, mOriginY As Double ' Точная идентичность ревизии и численного центра.
Private Const LEAF_CAPACITY As Long = 24 ' Малый неоднородный лист выгоднее обойти напрямую.
Private Const ROUNDING_GUARD As Double = 3.552713678800501E-15 ' 16 машинных epsilon для консервативной границы аффинного поля.

' Проверяет точные параметры подготовки; округлённых ключей нет.
Public Function Matches(ByVal revision As Long, ByVal originX As Double, ByVal originY As Double) As Boolean
    Matches = mNodeCount > 0 And revision = mRevision And originX = mOriginX And originY = mOriginY
End Function

' Копирует геометрию один раз для данной ревизии и строит независимый индекс.
Public Sub Initialize(ByVal section As CSectionModel, ByVal originX As Double, ByVal originY As Double)
    Dim count As Long, i As Long, x As Double, y As Double, area As Double
    count = section.ConcreteCount
    mRevision = section.Revision: mOriginX = originX: mOriginY = originY: mNodeCount = 0
    ReDim mFiberX(1 To count): ReDim mFiberY(1 To count): ReDim mArea(1 To count): ReDim mOrder(1 To count)
    ReDim mNodes(1 To count * 2)
    For i = 1 To count
        section.ConcreteElementForSolver i, x, y, area
        mFiberX(i) = x - originX: mFiberY(i) = y - originY: mArea(i) = area: mOrder(i) = i
    Next i
    BuildNode 1, count
    ReDim Preserve mNodes(1 To mNodeCount): ReDim mStack(1 To mNodeCount)
End Sub

' Делит по длинной стороне рамки. Моменты считаются в листьях, затем складываются вверх.
Private Function BuildNode(ByVal first As Long, ByVal count As Long) As Long
    mNodeCount = mNodeCount + 1
    Dim id As Long, i As Long, p As Long, splitX As Boolean, middle As Double
    id = mNodeCount
    With mNodes(id)
        .First = first: .Count = count
        .MinX = mFiberX(mOrder(first)): .MaxX = .MinX: .MinY = mFiberY(mOrder(first)): .MaxY = .MinY
        For i = first + 1 To first + count - 1
            p = mOrder(i)
            If mFiberX(p) < .MinX Then .MinX = mFiberX(p)
            If mFiberX(p) > .MaxX Then .MaxX = mFiberX(p)
            If mFiberY(p) < .MinY Then .MinY = mFiberY(p)
            If mFiberY(p) > .MaxY Then .MaxY = mFiberY(p)
        Next i
        splitX = .MaxX - .MinX >= .MaxY - .MinY
        If splitX Then middle = .MinX + (.MaxX - .MinX) / 2# Else middle = .MinY + (.MaxY - .MinY) / 2#
    End With
    Dim left As Long, right As Long, coordinate As Double, swap As Long
    left = first: right = first + count - 1
    If count > LEAF_CAPACITY Then
        Do While left <= right
            p = mOrder(left)
            If splitX Then coordinate = mFiberX(p) Else coordinate = mFiberY(p)
            If coordinate <= middle Then
                left = left + 1
            Else
                swap = mOrder(left): mOrder(left) = mOrder(right): mOrder(right) = swap: right = right - 1
            End If
        Loop
    End If
    Dim a As Long, b As Long, x As Double, y As Double, weight As Double
    If count <= LEAF_CAPACITY Or left = first Or left = first + count Then
        With mNodes(id)
            For i = first To first + count - 1
                p = mOrder(i): x = mFiberX(p): y = mFiberY(p): weight = mArea(p)
                .A = .A + weight: .AX = .AX + weight * x: .AY = .AY + weight * y
                .AXX = .AXX + weight * x * x: .AXY = .AXY + weight * x * y: .AYY = .AYY + weight * y * y
            Next i
        End With
    Else
        a = BuildNode(first, left - first): b = BuildNode(left, first + count - left)
        With mNodes(id)
            .Left = a: .Right = b
            .A = mNodes(a).A + mNodes(b).A: .AX = mNodes(a).AX + mNodes(b).AX: .AY = mNodes(a).AY + mNodes(b).AY
            .AXX = mNodes(a).AXX + mNodes(b).AXX: .AXY = mNodes(a).AXY + mNodes(b).AXY: .AYY = mNodes(a).AYY + mNodes(b).AYY
        End With
    End If
    BuildNode = id
End Function

' Вычисляет консервативный интервал аффинных деформаций по рамке узла.
Private Sub StrainBounds(ByRef node As TForceNode, ByVal eps As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef low As Double, ByRef high As Double, ByRef guard As Double)
    If kx >= 0# Then low = eps + kx * node.MinY: high = eps + kx * node.MaxY Else low = eps + kx * node.MaxY: high = eps + kx * node.MinY
    If ky >= 0# Then low = low + ky * node.MinX: high = high + ky * node.MaxX Else low = low + ky * node.MaxX: high = high + ky * node.MinX
    guard = ROUNDING_GUARD * (Abs(eps) + Abs(kx) * (Abs(node.MinY) + Abs(node.MaxY)) + Abs(ky) * (Abs(node.MinX) + Abs(node.MaxX)))
End Sub

' Возвращает усилия по той же сетке. Неоднозначный почти постоянный интервал
' у перелома сразу передаёт прямому обходу, чтобы не строить бесполезный индекс.
Public Function TryForces(ByVal diagram As CMaterialDiagram, ByVal eps As Double, ByVal kx As Double, ByVal ky As Double, _
        ByRef n As Double, ByRef mx As Double, ByRef my As Double) As Boolean
    Dim node As TForceNode, low As Double, high As Double, guard As Double, a As Double, b As Double
    node = mNodes(1)
    If kx = 0# And ky = 0# Then
        b = diagram.StressAtStrain(eps): n = b * node.A: mx = b * node.AY: my = b * node.AX
        TryForces = True: Exit Function
    End If
    StrainBounds node, eps, kx, ky, low, high, guard
    If high - low <= 8# * guard Then
        If Not diagram.TryLinearStressRange(low - guard, high + guard, a, b) Then Exit Function
    End If
    n = 0#: mx = 0#: my = 0#
    Dim top As Long, id As Long, i As Long, p As Long, x As Double, y As Double, force As Double
    top = 1: mStack(top) = 1
    Do While top > 0
        id = mStack(top): top = top - 1: node = mNodes(id)
        StrainBounds node, eps, kx, ky, low, high, guard
        If diagram.TryLinearStressRange(low - guard, high + guard, a, b) Then
            n = n + (a * eps + b) * node.A + a * (kx * node.AY + ky * node.AX)
            mx = mx + (a * eps + b) * node.AY + a * (kx * node.AYY + ky * node.AXY)
            my = my + (a * eps + b) * node.AX + a * (kx * node.AXY + ky * node.AXX)
        ElseIf node.Left > 0 Then
            top = top + 1: mStack(top) = node.Right
            top = top + 1: mStack(top) = node.Left
        Else
            For i = node.First To node.First + node.Count - 1
                p = mOrder(i): x = mFiberX(p): y = mFiberY(p)
                force = diagram.StressAtStrain(eps + kx * y + ky * x) * mArea(p)
                n = n + force: mx = mx + force * y: my = my + force * x
            Next i
        End If
    Loop
    TryForces = True
End Function
