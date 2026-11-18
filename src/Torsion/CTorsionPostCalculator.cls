VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPostCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Отдельная приближённая проверка замкнутой фермы на T после готового НДМ
' ==========================================================================
' Не вызывает решатель, не меняет нагрузки, диаграммы, состояния и результаты
' НДМ. Тяга от T распределяется по периметру на выбранные продольные стержни.
' sigmaNM + deltaSigmaT служит проверкой растяжения с учётом знака НДМ,
' а не новым НДС. Самостоятельные отказы НДМ остаются в его блоке Results.
' Бетонные диагонали проверяются отдельно от нормальных напряжений НДМ.

Private mSection As CSectionModel, mOuter As CTorsionHoopGeometry
Private mP As TTorsionPostParameters
Private mCoefficients() As Double
Private mCotangent As Double, mHoopArea As Double, mConcreteLimit As Double, mThickness As Double
Private mOuterArea As Double
Private mFailureSettingKeys As String
Private mConcreteFailureStatus As String, mConcreteFailureReason As String, mConcreteFailureKeys As String ' Отказ бетонной проверки не меняет подготовку продольной арматуры и хомута.

' Готовит геометрию один раз для всех сочетаний T. Входы уже приведены к
' внутренним единицам и проверены адаптером; геометрическая применимость здесь.
Friend Sub Prepare(ByVal section As CSectionModel, ByRef parameters As TTorsionPostParameters)
    If Not mSection Is Nothing Then ModelError "Геометрия этапа T уже подготовлена."
    If section Is Nothing Then ModelError "Не передана сохранённая геометрия сечения."
    If section.RebarCount < 3 Then ModelError "Для проверки T нужны как минимум три наружных продольных стержня."
    Set mSection = section
    mP.StateName = parameters.StateName: mP.AxisOffset = parameters.AxisOffset
    mP.HoopDiameter = parameters.HoopDiameter: mP.HoopSpacing = parameters.HoopSpacing
    mP.SteelResistance = parameters.SteelResistance: mP.HoopResistance = parameters.HoopResistance
    mP.BarBand = parameters.BarBand: mP.StrutAngle = parameters.StrutAngle
    mP.ConcreteMethod = parameters.ConcreteMethod: mP.WallThickness = parameters.WallThickness
    mP.ConcreteInputInvalid = parameters.ConcreteInputInvalid
    mP.ConcreteResistance = parameters.ConcreteResistance: mP.ConcreteEfficiency = parameters.ConcreteEfficiency
    mP.GorskyCoefficient = parameters.GorskyCoefficient
    mFailureSettingKeys = "Torsion.HoopAxisOffset,Torsion.HoopDiameter"
    Set mOuter = New CTorsionHoopGeometry
    mOuter.InitializeAutomatic section, mP.AxisOffset, mP.HoopDiameter
    Dim query As New CSectionGeometryQuery, domain As CConcreteRegion, holes As Long, i As Long
    Set domain = query.RegionFromContours(section.Contours)
    For i = 1 To domain.LoopCount
        If domain.LoopIsOpening(i) Then holes = holes + 1
    Next i
    If holes > 1 Then ModelError "Однокамерная проверка T не применяется к нескольким отверстиям."
    Dim demand As New CTorsionLongitudinalDemand
    mFailureSettingKeys = "Torsion.BarBand,Torsion.HoopAxisOffset,Torsion.HoopDiameter"
    demand.Initialize section, mOuter, mP.BarBand
    demand.CopyStressPerUnitForce mCoefficients
    For i = 1 To section.RebarCount
        If mCoefficients(i) > 0# Then mOuterArea = mOuterArea + section.RebarArea(i)
    Next i
    mCotangent = 1# / Tan(mP.StrutAngle * 4# * Atn(1#) / 180#)
    mHoopArea = 4# * Atn(1#) * mP.HoopDiameter ^ 2 / 4#
    PrepareConcreteCheck domain
End Sub

' Считает T независимо от физических пределов готового состояния. При
' отсутствии сохранённых напряжений только сумма T+NM получает N/A.
Friend Function Calculate(ByVal identifier As String, ByVal torque As Double, ByVal state As CSectionStateResult) As CTorsionPostResult
    If mSection Is Nothing Then ModelError "Геометрия отдельного этапа T не подготовлена."
    Dim row(1 To 1, 1 To TORSION_POST_COLUMNS) As Variant, bars() As Variant
    torque = Abs(torque)
    If torque = 0# Then ModelError "Нулевая строка T не требует расчёта."
    Dim q As Double, force As Double, hoopStress As Double, concreteStress As Double
    q = torque / (2# * mOuter.EnclosedArea)
    force = q * mOuter.Length * mCotangent ' Вся тяга от T распределяется только на наружную продольную арматуру.
    hoopStress = q * mP.HoopSpacing / (mHoopArea * mCotangent)
    If mThickness > 0# Then concreteStress = q * (mCotangent + 1# / mCotangent) / mThickness
    Dim available As Boolean
    If Not state Is Nothing Then available = state.Converged And state.HasMaterialSnapshot
    Dim snapshot As Variant
    If available Then snapshot = state.PrepareElementSnapshot(mSection)
    Dim count As Long, i As Long, outputRow As Long
    For i = 1 To mSection.RebarCount
        If mCoefficients(i) > 0# Then count = count + 1
    Next i
    ReDim bars(1 To count, 1 To TORSION_BAR_COLUMNS)
    Dim delta As Double, normal As Double, combined As Double, maxDelta As Double, maxCombined As Double
    Dim criticalT As Long, criticalCombined As Long, criticalNM As Double, criticalDelta As Double, criticalForce As Double
    Dim policy As New CResultStatusPolicy
    For i = 1 To mSection.RebarCount
        If mCoefficients(i) > 0# Then
            outputRow = outputRow + 1
            delta = force * mCoefficients(i)
            bars(outputRow, 1) = identifier: bars(outputRow, 2) = mSection.RebarID(i)
            bars(outputRow, 3) = delta * mSection.RebarArea(i)
            If criticalT = 0 Or delta > maxDelta Then criticalT = i: maxDelta = delta
            If available Then
                normal = CDbl(snapshot(mSection.ConcreteCount + i, 4))
                combined = normal + delta ' Положительное напряжение — растяжение; сжатие НДМ сохраняет знак минус.
                If criticalCombined = 0 Or combined > maxCombined Then
                    criticalCombined = i: maxCombined = combined
                    criticalNM = normal: criticalDelta = delta: criticalForce = combined * mSection.RebarArea(i)
                End If
            End If
        End If
    Next i
    row(1, 1) = identifier: row(1, 2) = mP.StateName: row(1, 3) = torque
    row(1, 4) = mOuter.EnclosedArea: row(1, 5) = mOuter.Length: row(1, 6) = mCotangent
    row(1, 7) = q: row(1, 8) = force: row(1, 9) = mOuterArea
    row(1, 10) = force / mP.SteelResistance: row(1, 11) = mSection.RebarID(criticalT)
    row(1, 12) = maxDelta: row(1, 13) = mP.SteelResistance
    row(1, 14) = mP.SteelResistance / maxDelta: row(1, 15) = TorsionPostCheckStatus(CDbl(row(1, 14)))
    row(1, 16) = mHoopArea: row(1, 17) = mP.HoopSpacing
    row(1, 18) = hoopStress: row(1, 19) = mP.HoopResistance
    row(1, 20) = mP.HoopResistance / hoopStress: row(1, 21) = TorsionPostCheckStatus(CDbl(row(1, 20)))
    row(1, 22) = mP.ConcreteMethod
    If Len(mConcreteFailureStatus) > 0 Then
        row(1, 28) = mConcreteFailureStatus
    ElseIf mP.ConcreteMethod = "None" Then
        row(1, 28) = policy.NotApplicable
    Else
        row(1, 23) = mThickness
        If mP.ConcreteMethod = "Strut" Then
            row(1, 24) = concreteStress: row(1, 25) = mP.ConcreteEfficiency * mP.ConcreteResistance
        End If
        row(1, 26) = mConcreteLimit: row(1, 27) = mConcreteLimit / torque
        row(1, 28) = TorsionPostCheckStatus(CDbl(row(1, 27)))
    End If
    If available Then
        row(1, 29) = mSection.RebarID(criticalCombined): row(1, 30) = criticalNM
        row(1, 31) = criticalDelta: row(1, 32) = criticalForce: row(1, 33) = maxCombined
        row(1, 34) = mP.SteelResistance
        If maxCombined > 0# Then
            row(1, 35) = mP.SteelResistance / maxCombined: row(1, 36) = TorsionPostCheckStatus(CDbl(row(1, 35)))
        Else
            row(1, 36) = policy.OK ' При сжатии растягивающий предел не достигается; числовой запас не ограничен.
        End If
    Else
        row(1, 36) = policy.NotApplicable
        If state Is Nothing Then
            row(1, 39) = "Сохранённое состояние " & mP.StateName & " не найдено; T рассчитан, сумма NMT недоступна. Включите получение этого состояния в расчётном профиле."
        Else
            row(1, 39) = "Сохранённые напряжения выбранного состояния недоступны для сложения с добавкой T."
        End If
    End If
    Dim stateFailureReason As String: stateFailureReason = CStr(row(1, 39))
    If Len(mConcreteFailureReason) > 0 Then
        If Len(CStr(row(1, 39))) > 0 Then row(1, 39) = CStr(row(1, 39)) & vbCrLf
        row(1, 39) = CStr(row(1, 39)) & mConcreteFailureReason
    End If
    row(1, 37) = mP.BarBand: row(1, 38) = policy.OK
    ApplyOverallStatus row, policy
    Dim result As New CTorsionPostResult: result.Initialize row, bars, stateFailureReason
    Set Calculate = result
End Function

' Собирает общий внешний статус только для блока NMT, не для сочетания НДМ.
Private Sub ApplyOverallStatus(ByRef row As Variant, ByVal policy As CResultStatusPolicy)
    Dim column As Variant
    For Each column In Array(21, 28, 36)
        If CStr(row(1, column)) = policy.CalcErr Then row(1, 38) = policy.CalcErr: Exit Sub
    Next column
    For Each column In Array(21, 28, 36)
        If CStr(row(1, column)) = policy.InputErr Then row(1, 38) = policy.InputErr: Exit Sub
    Next column
    For Each column In Array(21, 28, 36)
        If CStr(row(1, column)) = policy.Fail Then row(1, 38) = policy.Fail: Exit Sub
    Next column
    If CStr(row(1, 36)) = policy.NotApplicable Then row(1, 38) = policy.NotApplicable
End Sub

' Изолирует только подготовку бетонного предела. Общая геометрия, тяга и
' хомут к этому моменту готовы; неизвестный внутренний сбой сохраняет CalcErr.
Private Sub PrepareConcreteCheck(ByVal domain As CConcreteRegion)
    Dim policy As New CResultStatusPolicy
    If mP.ConcreteInputInvalid Then
        mConcreteFailureStatus = policy.InputErr
        mConcreteFailureReason = "Проверка бетона недоступна: исправьте её исходные данные."
        Exit Sub
    End If
    On Error GoTo Failed
    PrepareConcreteLimit domain
    Exit Sub
Failed:
    mConcreteFailureReason = Err.Description: mConcreteFailureKeys = mFailureSettingKeys
    If Err.Number = vbObjectError + 5670 Then mConcreteFailureStatus = policy.InputErr Else mConcreteFailureStatus = policy.CalcErr
    mThickness = 0#: mConcreteLimit = 0# ' Неподтверждённый предел не попадает в вывод как действительное число.
End Sub

' Возвращает причину независимого отказа бетонной проверки без адресов Excel.
Public Property Get ConcreteFailureReason() As String
    ConcreteFailureReason = mConcreteFailureReason
End Property

' Возвращает ключи параметров только отказавшей бетонной проверки.
Public Property Get ConcreteFailureSettingKeys() As String
    ConcreteFailureSettingKeys = mConcreteFailureKeys
End Property

' Готовит независимый предел бетонных раскосов или эмпирический предел
' сплошного прямоугольника по Горскому, §3, формула (13), стр.23 книги.
Private Sub PrepareConcreteLimit(ByVal domain As CConcreteRegion)
    mFailureSettingKeys = "Torsion.ConcreteMethod"
    Select Case mP.ConcreteMethod
        Case "None"
            Exit Sub
        Case "Strut"
            mFailureSettingKeys = "Torsion.WallThickness,Torsion.HoopAxisOffset,Torsion.HoopDiameter,Torsion.ConcreteEfficiency,Torsion.Rb"
            mThickness = mP.WallThickness
            If mThickness = 0# Then mThickness = 2# * mOuter.MinimumClearCover + mP.HoopDiameter
            If mThickness > 2# * mOuter.MinimumClearCover + mP.HoopDiameter + GEOM_BOUNDARY_TOLERANCE Then _
                ModelError "Torsion.WallThickness превышает толщину располагаемой бетонной полосы вокруг оси хомута. Уменьшите толщину или выберите 0 для автоматического определения. Эта настройка не задаёт полосу отбора арматуры; для неё используется Torsion.BarBand."
            mConcreteLimit = 2# * mOuter.EnclosedArea * mThickness * mP.ConcreteEfficiency * mP.ConcreteResistance / _
                (mCotangent + 1# / mCotangent)
        Case "GorskyRectangle"
            mFailureSettingKeys = "Torsion.ConcreteMethod,Torsion.GorskyCoefficient,Torsion.Rb"
            Dim b As Double, h As Double
            RectangleDimensions domain, b, h
            mConcreteLimit = mP.GorskyCoefficient * mP.ConcreteResistance * b * b * h
        Case Else
            ModelError "Неизвестный метод отдельной проверки бетона на T."
    End Select
End Sub

' Допускает формулу b^2*h только для одного сплошного прямоугольника;
' поворот прямоугольника не меняет длины его сторон.
Private Sub RectangleDimensions(ByVal domain As CConcreteRegion, ByRef b As Double, ByRef h As Double)
    Dim edges() As TRegionEdge, count As Long, i As Long, j As Long
    domain.CopyEdges edges, count
    If domain.LoopCount <> 1 Or count <> 4 Then ModelError "GorskyRectangle применим только к сплошному прямоугольнику без фасок, дуг и отверстий. Выберите Strut или None."
    Dim lengths(1 To 4) As Double, dot As Double
    For i = 1 To 4
        If Abs(edges(i).Sweep) > GEOM_PARAMETER_TOLERANCE Then ModelError "GorskyRectangle не применяется к скруглённым углам."
        lengths(i) = EdgeLength(edges(i))
    Next i
    For i = 1 To 4
        j = i Mod 4 + 1
        dot = (edges(i).X2 - edges(i).X1) * (edges(j).X2 - edges(j).X1) + _
            (edges(i).Y2 - edges(i).Y1) * (edges(j).Y2 - edges(j).Y1)
        If Abs(dot) > 1E-8 * lengths(i) * lengths(j) Then ModelError "GorskyRectangle требует прямые углы бетонного контура."
    Next i
    b = GeomMin(lengths(1), lengths(2)): h = GeomMax(lengths(1), lengths(2))
End Sub

' Возвращает ключи входов текущего геометрического этапа для общего адаптера
' адресов; никаких ссылок на Excel или текстовых разборов причины в ядре нет.
Public Property Get FailureSettingKeys() As String
    FailureSettingKeys = mFailureSettingKeys
End Property

' Возвращает методическую или геометрическую причину без обращения к Excel.
Private Sub ModelError(ByVal reason As String)
    Err.Raise vbObjectError + 5670, TypeName(Me), reason
End Sub
