VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CTorsionConcreteCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Независимая проверка бетона при кручении по выбранной методике
' ==========================================================================
' Strut: ACI 445.1R-12, изложение EC2-04, 9.2.2(b,d,e), пример 9.4.3.
' GorskyRectangle: Горский, 1961, §3, (13). Только ненапрягаемая арматура.
' Геометрия и сопротивление готовятся один раз; НДС НДМ и арматура не меняются.
Private mArea As Double, mThickness As Double, mDesignStrength As Double
Private mEfficiency As Double, mResistance As Double, mLimit As Double, mCotangent As Double
Private mFailureKeys As String

' Подготавливает единую эффективную стенку Strut либо сплошной прямоугольник Горского.
' Толщина стенки и площадь её средней линии не подменяются осью наружного хомута.
Friend Sub Prepare(ByVal domain As CConcreteRegion, ByRef p As TTorsionPostParameters)
    Select Case p.ConcreteMethod
        Case "None": Exit Sub
        Case "GorskyRectangle"
            mFailureKeys = "Torsion.ConcreteMethod,Torsion.GorskyCoefficient,Torsion.Rb"
            If p.GorskyCoefficient <= 0# Or p.ConcreteResistance <= 0# Then CheckError "Сопротивление бетона и коэффициент Горского должны быть положительными."
            Dim b As Double, h As Double: RectangleDimensions domain, b, h
            mLimit = p.GorskyCoefficient * p.ConcreteResistance * b * b * h
        Case "Strut"
            PrepareStrut domain, p
        Case Else
            mFailureKeys = "Torsion.ConcreteMethod": CheckError "Неизвестная методика проверки бетона при кручении."
    End Select
End Sub

' Вычисляет сопротивление раскоса по f_ck и gamma_c; alpha_cw = 1 для ненапрягаемого бетона.
' Проверяет предел cot(theta) из источника независимо от доступности арматурной проверки.
Private Sub PrepareStrut(ByVal domain As CConcreteRegion, ByRef p As TTorsionPostParameters)
    mFailureKeys = "Torsion.StrutAngle"
    If p.StrutAngle <= 0# Or p.StrutAngle >= 90# Then CheckError "Угол раскосов должен находиться строго между 0 и 90 градусами."
    mCotangent = 1# / Tan(p.StrutAngle * GEOM_PI / 180#)
    If mCotangent < 1# - GEOM_PARAMETER_TOLERANCE Or mCotangent > 2.5 + GEOM_PARAMETER_TOLERANCE Then _
        CheckError "Для Strut источник требует 1 <= cot(theta) <= 2,5: угол от 21,801409 до 45 градусов."
    mFailureKeys = "Torsion.Fck,Torsion.GammaC"
    If p.CharacteristicConcreteStrength <= 0# Or p.CharacteristicConcreteStrength >= 250# Then _
        CheckError "Цилиндрическая характеристическая прочность f_ck должна быть положительной и давать положительный коэффициент 0,6(1-f_ck/250 МПа)."
    If p.ConcretePartialFactor < 1# Then CheckError "Коэффициент надёжности gamma_c должен быть не меньше 1."
    mDesignStrength = p.CharacteristicConcreteStrength / p.ConcretePartialFactor
    mEfficiency = 0.6 * (1# - p.CharacteristicConcreteStrength / 250#)
    mResistance = mEfficiency * mDesignStrength ' ACI (9.2.2b,d,e): alpha_cw = 1 для ненапрягаемой арматуры.
    mFailureKeys = "Torsion.RebarAxisOffset,Torsion.WallThickness"
    If p.LongitudinalAxisOffset <= 0# Or p.WallThickness < 0# Then CheckError "Расстояние до центра продольной арматуры должно быть положительным, а заданная толщина — неотрицательной."
    PrepareWall domain, p.LongitudinalAxisOffset, p.WallThickness
    mLimit = 2# * mArea * mThickness * mResistance / (mCotangent + 1# / mCotangent)
End Sub

' Строит среднюю линию стенки на t_eff/2 от наружного контура.
' A включает отверстие; u относится только к наружной границе. Для полого сечения
' фактическая минимальная толщина ограничивает единую эффективную стенку сверху.
Private Sub PrepareWall(ByVal domain As CConcreteRegion, ByVal centreOffset As Double, ByVal specified As Double)
    Dim query As New CSectionGeometryQuery, outer As CGeometryPath, opening As CGeometryPath
    Dim loopIndex As Long, outerIndex As Long, holes As Long
    For loopIndex = 1 To domain.LoopCount
        If domain.LoopIsOpening(loopIndex) Then
            holes = holes + 1: Set opening = BoundaryPath(domain, loopIndex)
        Else
            If outerIndex > 0 Then CheckError "Для одной эффективной стенки Strut нужен один наружный контур."
            outerIndex = loopIndex
        End If
    Next loopIndex
    If outerIndex = 0 Or holes > 1 Then CheckError "Для одной эффективной стенки Strut требуется один наружный контур и не более одного отверстия."
    Set outer = BoundaryPath(domain, outerIndex)
    Dim outerRegion As CConcreteRegion: Set outerRegion = query.RegionFromPath(outer)
    Dim wallLimit As Double: wallLimit = 1E+100
    If Not opening Is Nothing Then wallLimit = PathDistance(outer, opening)
    If 2# * centreOffset > wallLimit + GEOM_BOUNDARY_TOLERANCE Then _
        CheckError "Не существует эффективной стенки: удвоенное расстояние до центра продольной арматуры превышает фактическую толщину бетона."
    mThickness = specified
    If mThickness = 0# Then mThickness = GeomMin(GeomMax(outerRegion.Area / outer.Length, 2# * centreOffset), wallLimit) ' ACI, пример 9.4.3: max(A/u; 2c_l); EC2 6.3.2(1): фактическая стенка — верхний предел.
    If mThickness + GEOM_BOUNDARY_TOLERANCE < 2# * centreOffset Then CheckError "Толщина эффективной стенки должна быть не меньше удвоенного расстояния до центра продольной арматуры."
    If mThickness > wallLimit + GEOM_BOUNDARY_TOLERANCE Then CheckError "Заданная эффективная толщина превышает фактическую минимальную толщину бетонной стенки."
    Dim axis As CGeometryPath
    On Error GoTo InvalidOffset
    Set axis = query.OffsetPath("Torsion.Concrete.Midline", outer, 0.5 * mThickness)
    Dim effective As CConcreteRegion: Set effective = query.RegionFromPath(axis)
    On Error GoTo 0
    If effective.LoopCount <> 1 Or effective.Area <= 0# Then CheckError "Средняя линия эффективной стенки не образует одну пригодную замкнутую область."
    Dim firstEdge As TRegionEdge: axis.CopySegment 1, firstEdge
    If Not query.ContainsPoint(domain, firstEdge.X1, firstEdge.Y1, False) Then CheckError "Средняя линия эффективной стенки находится вне бетона."
    If PathDistance(axis, outer) + GEOM_BOUNDARY_TOLERANCE < 0.5 * mThickness Then _
        CheckError "Эффективная стенка выходит за наружную бетонную границу. Проверьте допустимость одной стенки и её толщину."
    If Not opening Is Nothing Then
        opening.CopySegment 1, firstEdge
        If Not query.ContainsPoint(effective, firstEdge.X1, firstEdge.Y1, True) Then CheckError "Средняя линия эффективной стенки не охватывает внутреннюю полость."
        If PathDistance(axis, opening) + GEOM_BOUNDARY_TOLERANCE < 0.5 * mThickness Then _
            CheckError "Эффективная бетонная стенка пересекает отверстие. Уменьшите толщину, сохранив нижний предел 2c_l."
    End If
    mArea = effective.Area ' Площадь внутри средней линии включает внутреннюю полость согласно определению A_k источника.
    Exit Sub
InvalidOffset:
    Dim reason As String: reason = Err.Description
    On Error GoTo 0
    CheckError "Не удалось построить среднюю линию эффективной стенки. " & reason & " Для сложного контура проверьте допустимость одной замкнутой стенки и её толщину."
End Sub

' Переносит ориентированную границу без аппроксимации дуг и без изменения исходного сечения.
Private Function BoundaryPath(ByVal domain As CConcreteRegion, ByVal loopIndex As Long) As CGeometryPath
    Dim edges() As TRegionEdge, count As Long, i As Long, path As New CGeometryPath
    domain.CopyEdges edges, count: path.Initialize "Torsion.Concrete.Boundary"
    For i = 1 To count
        If edges(i).LoopID = loopIndex Then path.Append "E" & CStr(i), edges(i)
    Next i
    path.Seal True: Set BoundaryPath = path
End Function

' Находит точное минимальное расстояние между всеми прямыми/дугами двух контуров.
Private Function PathDistance(ByVal first As CGeometryPath, ByVal second As CGeometryPath) As Double
    Dim metric As New CGeometryDistance, a As TRegionEdge, b As TRegionEdge, i As Long, j As Long
    PathDistance = 1E+100
    For i = 1 To first.SegmentCount
        first.CopySegment i, a
        For j = 1 To second.SegmentCount
            second.CopySegment j, b
            PathDistance = GeomMin(PathDistance, metric.EdgeDistance(a, b))
        Next j
    Next i
End Function

' Проверяет фактический сплошной прямоугольник независимо от поворота и названия рецепта.
Private Sub RectangleDimensions(ByVal domain As CConcreteRegion, ByRef b As Double, ByRef h As Double)
    Dim edges() As TRegionEdge, count As Long, i As Long, j As Long
    domain.CopyEdges edges, count
    If domain.LoopCount <> 1 Or count <> 4 Then CheckError "GorskyRectangle применим только к сплошному прямоугольнику без фасок, дуг и отверстий. Выберите Strut или None."
    Dim lengths(1 To 4) As Double, dot As Double
    For i = 1 To 4
        If Abs(edges(i).Sweep) > GEOM_PARAMETER_TOLERANCE Then CheckError "GorskyRectangle не применяется к скруглённым углам."
        lengths(i) = EdgeLength(edges(i))
    Next i
    For i = 1 To 4
        j = i Mod 4 + 1
        dot = (edges(i).X2 - edges(i).X1) * (edges(j).X2 - edges(j).X1) + (edges(i).Y2 - edges(i).Y1) * (edges(j).Y2 - edges(j).Y1)
        If Abs(dot) > 1E-8 * lengths(i) * lengths(j) Then CheckError "GorskyRectangle требует прямые углы бетонного контура."
    Next i
    b = GeomMin(lengths(1), lengths(2)): h = GeomMax(lengths(1), lengths(2))
End Sub

' Возвращает сжатие диагонального раскоса от модуля T во внутренних единицах.
Friend Function Stress(ByVal torque As Double) As Double
    If mArea > 0# And mThickness > 0# Then Stress = Abs(torque) * (mCotangent + 1# / mCotangent) / (2# * mArea * mThickness)
End Function

' Возвращает проверенную площадь внутри средней линии эффективной стенки, мм2.
Public Property Get EffectiveArea() As Double: EffectiveArea = mArea: End Property
' Возвращает полную эффективную толщину, мм.
Public Property Get Thickness() As Double: Thickness = mThickness: End Property
' Возвращает расчётную цилиндрическую прочность f_cd, МПа.
Public Property Get DesignStrength() As Double: DesignStrength = mDesignStrength: End Property
' Возвращает коэффициент снижения прочности растрескавшегося бетона nu.
Public Property Get Efficiency() As Double: Efficiency = mEfficiency: End Property
' Возвращает допустимое сжатие диагонального раскоса, МПа.
Public Property Get Resistance() As Double: Resistance = mResistance: End Property
' Возвращает предельный крутящий момент по бетону, Н*мм.
Public Property Get Limit() As Double: Limit = mLimit: End Property
' Возвращает ключи активных исходных данных для общего адресного диагностического слоя.
Public Property Get FailureSettingKeys() As String: FailureSettingKeys = mFailureKeys: End Property

' Сохраняет общий код ошибки исходных данных без привязки численного ядра к Excel.
Private Sub CheckError(ByVal reason As String)
    Err.Raise vbObjectError + 5670, TypeName(Me), reason
End Sub
