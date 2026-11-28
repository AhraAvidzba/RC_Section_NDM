VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CThickLayerIntegrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Экспериментальное интегрирование бетона по контуру и полосам диаграммы
' ==========================================================================
' MTLI: Bonet et al., Computers & Structures 82 (2004), 213–225,
' doi:10.1016/j.compstruc.2003.10.009, раздел 3, формулы (1)–(6).
' Интегрирует непрерывную область, а не сумму площадей в центрах сетки.
' Применим к единой кусочно-линейной диаграмме без истории нагружения.
' Полосы отделяются точно по узлам диаграммы. Замыкающие стороны полос
' параллельны u и имеют dv=0, поэтому достаточно разрезать исходные рёбра.
' Прямые интегрируются двумя точками Гаусса; дуги — восемью на интервалах
' не больше pi/8. Дуги сохраняются аналитическими, без замены хордами.
' Возвращает R и согласованную K; принятие равновесия и физических пределов
' остаётся обязанностью решателя. Excel и настройки листов не используются.

Private Type TBoundary
    X1 As Double
    Y1 As Double
    X2 As Double
    Y2 As Double
    CX As Double
    CY As Double
    Radius As Double
    StartAngle As Double
    Sweep As Double
End Type
Private mEdges() As TBoundary, mCount As Long ' Проверенная граница с бетоном слева, в численном центре.
Private mR(1 To 3) As Double, mK(1 To 3, 1 To 3) As Double ' Усилия [N,Mx,My] и их производные по [eps,kx,ky].
Private mH(0 To 5) As Double, mF(0 To 2) As Double ' Интегралы Et*[1,u,v,u^2,uv,v^2] и sigma*[1,u,v].
Private mNX As Double, mNY As Double, mGradient As Double, mEps As Double ' v=nx*x+ny*y; u=ny*x-nx*y.
Private mMinStrain As Double, mMaxStrain As Double, mEvaluations As Long ' Экстремумы по точной границе и фактическое число точек Гаусса.
Private Const PI As Double = 3.1415926535897931

' Копирует нормализованную область общего геометрического ядра.
' Перенос координат перед вычислениями уменьшает потерю точности в моментах.
Public Sub Initialize(ByVal region As CConcreteRegion, ByVal originX As Double, ByVal originY As Double)
    mCount = 0
    If region Is Nothing Then Err.Raise vbObjectError + 5901, "CThickLayerIntegrator", "Не передана бетонная область."
    If region.SegmentCount = 0 Then Err.Raise vbObjectError + 5901, "CThickLayerIntegrator", "Бетонная область не содержит границ."
    mCount = region.SegmentCount: ReDim mEdges(1 To mCount)
    Dim i As Long, loopID As Long
    For i = 1 To mCount
        With mEdges(i)
            region.GetSegment i, .X1, .Y1, .X2, .Y2, .CX, .CY, .Sweep, loopID
            .X1 = .X1 - originX: .Y1 = .Y1 - originY
            .X2 = .X2 - originX: .Y2 = .Y2 - originY
            .CX = .CX - originX: .CY = .CY - originY
            If .Sweep <> 0# Then
                .Radius = Sqr((.X1 - .CX) ^ 2 + (.Y1 - .CY) ^ 2)
                .StartAngle = AngleOf(.X1 - .CX, .Y1 - .CY)
            End If
        End With
    Next i
End Sub

' Находит интегралы для одной плоскости. При нулевой кривизне выбирается
' произвольная ортонормальная система u=x, v=y без деления на кривизну.
Public Sub Evaluate(ByVal diagram As CMaterialDiagram, ByVal epsilon As Double, ByVal kx As Double, ByVal ky As Double)
    If mCount = 0 Then Err.Raise vbObjectError + 5902, "CThickLayerIntegrator", "Интегратор не подготовлен."
    If diagram Is Nothing Then Err.Raise vbObjectError + 5902, "CThickLayerIntegrator", "Диаграмма бетона не передана."
    Erase mR: Erase mK: Erase mF: Erase mH
    mEps = epsilon: mEvaluations = 0
    Dim curvatureScale As Double: curvatureScale = Abs(kx): If Abs(ky) > curvatureScale Then curvatureScale = Abs(ky)
    mGradient = 0#: mNX = 0#: mNY = 1#
    If curvatureScale > 0# Then
        mGradient = curvatureScale * Sqr((kx / curvatureScale) ^ 2 + (ky / curvatureScale) ^ 2)
        mNX = ky / mGradient: mNY = kx / mGradient
    End If
    mMinStrain = 1E+100: mMaxStrain = -1E+100
    Dim i As Long
    For i = 1 To mCount: IntegrateEdge diagram, mEdges(i): Next i
    mR(1) = mF(0)
    mR(2) = -mNX * mF(1) + mNY * mF(2)
    mR(3) = mNY * mF(1) + mNX * mF(2)
    mK(1, 1) = mH(0)
    mK(1, 2) = -mNX * mH(1) + mNY * mH(2)
    mK(1, 3) = mNY * mH(1) + mNX * mH(2)
    mK(2, 2) = mNX * mNX * mH(3) - 2# * mNX * mNY * mH(4) + mNY * mNY * mH(5)
    mK(2, 3) = -mNX * mNY * mH(3) + (mNY * mNY - mNX * mNX) * mH(4) + mNX * mNY * mH(5)
    mK(3, 3) = mNY * mNY * mH(3) + 2# * mNX * mNY * mH(4) + mNX * mNX * mH(5)
    mK(2, 1) = mK(1, 2): mK(3, 1) = mK(1, 3): mK(3, 2) = mK(2, 3)
End Sub

' Разрезает ребро по всем узлам диаграммы, включая границы продолжений.
' У дуги учитывает оба пересечения каждого уровня и точки экстремума v.
Private Sub IntegrateEdge(ByVal diagram As CMaterialDiagram, ByRef edge As TBoundary)
    Dim cuts() As Double, count As Long, i As Long, j As Long, knot As Double
    ReDim cuts(1 To 2 + 18 * diagram.PointCount + 32)
    count = 2: cuts(1) = 0#: cuts(2) = 1#
    Dim e1 As Double, e2 As Double
    e1 = StrainAt(edge.X1, edge.Y1): e2 = StrainAt(edge.X2, edge.Y2)
    RecordStrain e1: RecordStrain e2
    If edge.Sweep = 0# Then
        If e1 <> e2 Then
            For i = 1 To diagram.PointCount
                knot = diagram.PointStrain(i)
                If knot > MinValue(e1, e2) And knot < MaxValue(e1, e2) Then AddCut cuts, count, (knot - e1) / (e2 - e1)
            Next i
        End If
    Else
        Dim parts As Long: parts = Int(Abs(edge.Sweep) / (PI / 8#)) + 1
        For i = 1 To parts - 1: AddCut cuts, count, CDbl(i) / parts: Next i
        Dim phase As Double: phase = AngleOf(mNX, mNY)
        AddArcAngles cuts, count, edge, phase
        AddArcAngles cuts, count, edge, phase + PI
        Dim centerV As Double, level As Double, angle As Double
        centerV = mNX * edge.CX + mNY * edge.CY
        If mGradient <> 0# Then
            For i = 1 To diagram.PointCount
                level = ((diagram.PointStrain(i) - mEps) / mGradient - centerV) / edge.Radius
                If Abs(level) < 1# Then
                    angle = PI / 2# - Atn(level / Sqr(1# - level * level))
                    AddArcAngles cuts, count, edge, phase + angle
                    AddArcAngles cuts, count, edge, phase - angle
                End If
            Next i
        End If
    End If
    Dim temp As Double
    For i = 2 To count
        temp = cuts(i): j = i - 1
        Do While j >= 1
            If cuts(j) <= temp Then Exit Do
            cuts(j + 1) = cuts(j): j = j - 1
        Loop
        cuts(j + 1) = temp
    Next i
    For i = 1 To count - 1
        If cuts(i + 1) > cuts(i) Then IntegrateInterval diagram, edge, cuts(i), cuts(i + 1)
    Next i
End Sub

' Сохраняет только внутренние точки без округления близких пересечений.
Private Sub AddCut(ByRef cuts() As Double, ByRef count As Long, ByVal value As Double)
    If value <= 0# Or value >= 1# Then Exit Sub
    count = count + 1
    If count > UBound(cuts) Then ReDim Preserve cuts(1 To count * 2)
    cuts(count) = value
End Sub

' Приводит эквивалентные углы к параметру данной дуги; знак обхода сохранён.
Private Sub AddArcAngles(ByRef cuts() As Double, ByRef count As Long, ByRef edge As TBoundary, ByVal angle As Double)
    Dim turn As Long, t As Double, actualAngle As Double
    For turn = -3 To 3
        actualAngle = angle + 2# * PI * turn
        t = (actualAngle - edge.StartAngle) / edge.Sweep
        If t > 0# And t < 1# Then
            AddCut cuts, count, t
            RecordStrain StrainAt(edge.CX + edge.Radius * Cos(actualAngle), edge.CY + edge.Radius * Sin(actualAngle))
        End If
    Next turn
End Sub

' Выполняет квадратуру одной гладкой части ребра. Две точки точны для
' полиномов до третьей степени на прямой, восемь разрешают кривизну дуги.
Private Sub IntegrateInterval(ByVal diagram As CMaterialDiagram, ByRef edge As TBoundary, ByVal low As Double, ByVal high As Double)
    Dim mid As Double, half As Double, i As Long, point As Double, weight As Double
    mid = (low + high) / 2#: half = (high - low) / 2#
    If edge.Sweep = 0# Then
        AddPoint diagram, edge, mid - half / Sqr(3#), half
        AddPoint diagram, edge, mid + half / Sqr(3#), half
    Else
        For i = 1 To 4
            Select Case i
                Case 1: point = 0.1834346424956498: weight = 0.362683783378362
                Case 2: point = 0.525532409916329: weight = 0.3137066458778873
                Case 3: point = 0.7966664774136267: weight = 0.2223810344533745
                Case 4: point = 0.9602898564975363: weight = 0.1012285362903763
            End Select
            AddPoint diagram, edge, mid - half * point, half * weight
            AddPoint diagram, edge, mid + half * point, half * weight
        Next i
    End If
End Sub

' Формулы Грина в системе (u,v): интеграл u^p*v^q равен
' обходу u^(p+1)*v^q/(p+1)*dv. Переходы материала уже разделены.
Private Sub AddPoint(ByVal diagram As CMaterialDiagram, ByRef edge As TBoundary, ByVal t As Double, ByVal weight As Double)
    Dim x As Double, y As Double, dv As Double, angle As Double
    If edge.Sweep = 0# Then
        x = edge.X1 + t * (edge.X2 - edge.X1): y = edge.Y1 + t * (edge.Y2 - edge.Y1)
        dv = mNX * (edge.X2 - edge.X1) + mNY * (edge.Y2 - edge.Y1)
    Else
        angle = edge.StartAngle + t * edge.Sweep
        x = edge.CX + edge.Radius * Cos(angle): y = edge.CY + edge.Radius * Sin(angle)
        dv = edge.Radius * edge.Sweep * (-mNX * Sin(angle) + mNY * Cos(angle))
    End If
    If dv = 0# Then Exit Sub
    Dim u As Double, v As Double, stress As Double, tangent As Double, w As Double
    u = mNY * x - mNX * y: v = mNX * x + mNY * y
    diagram.EvaluateAtStrain mEps + mGradient * v, stress, tangent
    mEvaluations = mEvaluations + 1: w = weight * dv
    mF(0) = mF(0) + w * u * stress
    mF(1) = mF(1) + w * u * u * stress / 2#
    mF(2) = mF(2) + w * u * v * stress
    mH(0) = mH(0) + w * u * tangent
    mH(1) = mH(1) + w * u * u * tangent / 2#
    mH(2) = mH(2) + w * u * v * tangent
    mH(3) = mH(3) + w * u * u * u * tangent / 3#
    mH(4) = mH(4) + w * u * u * v * tangent / 2#
    mH(5) = mH(5) + w * u * v * v * tangent
End Sub

' Вычисляет деформацию точки в сохранённой численной системе координат.
Private Function StrainAt(ByVal x As Double, ByVal y As Double) As Double
    StrainAt = mEps + mGradient * (mNX * x + mNY * y)
End Function

' Обновляет экстремумы на концах рёбер и в стационарных точках дуг.
Private Sub RecordStrain(ByVal value As Double)
    If value < mMinStrain Then mMinStrain = value
    If value > mMaxStrain Then mMaxStrain = value
End Sub

' Возвращает угол в [-pi,pi], включая вертикальное направление.
Private Function AngleOf(ByVal x As Double, ByVal y As Double) As Double
    If x = 0# Then
        If y > 0# Then AngleOf = PI / 2#
        If y < 0# Then AngleOf = -PI / 2#
    Else
        AngleOf = Atn(y / x)
        If x < 0# Then
            If y < 0# Then AngleOf = AngleOf - PI Else AngleOf = AngleOf + PI
        End If
    End If
End Function

' Выбирает меньшую границу интервала без изменения её точности.
Private Function MinValue(ByVal first As Double, ByVal second As Double) As Double
    MinValue = first: If second < first Then MinValue = second
End Function

' Выбирает большую границу интервала без изменения её точности.
Private Function MaxValue(ByVal first As Double, ByVal second As Double) As Double
    MaxValue = first: If second > first Then MaxValue = second
End Function

' Внутреннее усилие: 1=N, 2=Mx, 3=My, относительно центра подготовки.
Public Property Get Force(ByVal component As Long) As Double
    Force = mR(component)
End Property

' Производная усилия по параметру поля [eps,kx,ky] в тех же координатах.
Public Property Get Tangent(ByVal row As Long, ByVal column As Long) As Double
    Tangent = mK(row, column)
End Property

' Минимальная деформация на точной границе бетонной области.
Public Property Get MinStrain() As Double
    MinStrain = mMinStrain
End Property

' Максимальная деформация на точной границе бетонной области.
Public Property Get MaxStrain() As Double
    MaxStrain = mMaxStrain
End Property

' Число фактически вычисленных пар stress/Et при последней квадратуре.
Public Property Get EvaluationCount() As Long
    EvaluationCount = mEvaluations
End Property
