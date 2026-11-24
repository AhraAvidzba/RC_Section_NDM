VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilitySP35Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Нормативный расчёт устойчивости по СП 35
' ==========================================================================
' Работает с двумя заранее подготовленными главными плоскостями.
' Геометрию и исходные нагрузки не читает из Excel; готовые моменты возвращает фасаду.

' Диагностическая ветвь передаётся ByRef: завершение первой плоскости остаётся
' наблюдаемым и при отказе второй. Объект не хранит состояние между вызовами.

' ----------------------------- СП 35 --------------------------------------

' Расчет по СП 35. Для малых эксцентриситетов используется табличная ветвь
' phi по таблице 7.21, для больших - ветвь eta/Ncr по п. 7.54.
Friend Sub CalculateSP35(ByRef normBranch As String, ByVal normTableInputError As String, _
        ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal valueSet As EMaterialValueSet, ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByVal eb As Double, ByVal es As Double, ByVal systemType As String, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane)
    Dim rb As Double
    Dim rsc As Double
    rb = materialProvider.ConcreteCompressionResistanceFromValueSet(valueSet)
    rsc = materialProvider.SteelCompressionResistanceFromValueSet(valueSet)
    If rb <= 0# Or rsc <= 0# Then StabilityRaiseInput "Для расчета устойчивости СП 35 нужны положительные Eb, Es, Rb и Rsc."
    CalculateSP35Plane normBranch, normTableInputError, concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane1, "1"
    CalculateSP35Plane normBranch, normTableInputError, concreteArea, totalSteelArea, nCompression, sustainedCompression, eb, es, rb, rsc, systemType, phiLMode, phiP, nOverNcrLimit, table721, plane2, "2"
    normBranch = CombinedSP35Branch(plane1, plane2)
End Sub

' Возвращает общий диагностический признак СП 35 без потери смысла, когда две
' главные плоскости попали в разные нормативные ветви расчета.
Private Function CombinedSP35Branch(ByRef plane1 As TStabilityPlane, ByRef plane2 As TStabilityPlane) As String
    Dim branch1 As String
    Dim branch2 As String
    branch1 = plane1.BranchName
    branch2 = plane2.BranchName

    If Not plane1.Applicable Then branch1 = vbNullString
    If Not plane2.Applicable Then branch2 = vbNullString

    If Len(branch1) = 0 And Len(branch2) = 0 Then
        CombinedSP35Branch = STABILITY_CODE_SP35
    ElseIf Len(branch1) = 0 Then
        CombinedSP35Branch = branch2
    ElseIf Len(branch2) = 0 Then
        CombinedSP35Branch = branch1
    ElseIf StrComp(branch1, branch2, vbTextCompare) = 0 Then
        CombinedSP35Branch = branch1
    Else
        CombinedSP35Branch = "SP35-mixed"
    End If
End Function

' Считает одну главную плоскость по СП 35: табличная ветвь при |ec|/r <= 1,
' иначе eta/Ncr. Таблица 7.21 передается с Config и не зашита в код.
Private Sub CalculateSP35Plane(ByRef normBranch As String, ByVal normTableInputError As String, _
        ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal rsc As Double, ByVal systemType As String, ByVal phiLMode As String, _
        ByVal phiP As Double, ByVal nOverNcrLimit As Double, ByVal table721 As Variant, _
        ByRef plane As TStabilityPlane, ByVal planeLabel As String)
    StabilityEnsurePositivePlaneGeometry plane, STABILITY_CODE_SP35, planeLabel

    plane.Eccentricity = StabilityInitialEccentricity(plane.Moment, nCompression, _
        plane.AccidentalEccentricity, systemType)

    Dim coreDistance As Double
    coreDistance = CoreDistanceForEccentricity(plane)
    If coreDistance <= STABILITY_SMALL_VALUE Then StabilityRaiseInput "Для расчета устойчивости СП 35 нужно положительное ядровое расстояние r."
    plane.CoreDistance = coreDistance

    Dim q As Double
    q = Abs(plane.Eccentricity) / coreDistance ' Модуль эксцентриситета делится на ядровое расстояние выбранной стороны.
    If q <= 1# Then
        CalculateSP35TablePlane normBranch, normTableInputError, concreteArea, totalSteelArea, nCompression, sustainedCompression, rb, rsc, table721, q, plane
    Else
        CalculateSP35EtaPlane normBranch, nCompression, sustainedCompression, eb, es, rb, phiLMode, phiP, nOverNcrLimit, plane
    End If
End Sub

' Табличная ветвь СП 35 для |ec|/r <= 1.
Private Sub CalculateSP35TablePlane(ByRef normBranch As String, ByVal normTableInputError As String, _
        ByVal concreteArea As Double, ByVal totalSteelArea As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double, ByVal rb As Double, ByVal rsc As Double, _
        ByVal table721 As Variant, ByVal q As Double, _
        ByRef plane As TStabilityPlane)
    Dim slenderness As Double
    slenderness = plane.Slenderness

    Dim phiM As Double
    Dim phiLTable As Double
    InterpolateSP35Table normTableInputError, table721, slenderness, q, phiM, phiLTable
    plane.PhiL = phiLTable
    plane.PhiM = phiM
    plane.PhiLTable = phiLTable

    Dim nUltimate As Double
    Dim phiValue As Double
    phiValue = PhiForSustainedCompression(phiM, plane.PhiL, nCompression, sustainedCompression)
    plane.PhiValue = phiValue
    nUltimate = phiValue * (rb * SP35ConcreteAreaForTableNult(concreteArea, totalSteelArea) + _
        rsc * totalSteelArea)

    plane.NultimateForce = nUltimate
    plane.CriticalForce = plane.NultimateForce
    plane.ReserveFactor = StabilityReserveByCriticalForce(nUltimate, nCompression)
    plane.Eta = 1#
    plane.StiffnessD = 0#
    plane.DesignMoment = nCompression * plane.Eccentricity
    plane.CheckPassed = (plane.ReserveFactor + STABILITY_PASS_TOLERANCE >= 1#)
    normBranch = "SP35-table"
    plane.BranchName = "SP35-table"
End Sub

' Возвращает бетонную площадь для формулы Nult табличной ветви СП 35.
' Приведенное сечение уже строится через (n - 1)As и этим определяет центр,
' оси и ядровое расстояние. Здесь работает другая нормативная поправка:
' если суммарная площадь стали больше 3% gross-площади бетона, бетонный
' вклад в Nult уменьшается на As, чтобы не завышать сжатую бетонную часть.
Private Function SP35ConcreteAreaForTableNult(ByVal concreteArea As Double, _
        ByVal totalSteelArea As Double) As Double
    If concreteArea <= STABILITY_SMALL_VALUE Then StabilityRaiseInput "Для табличной ветви СП 35 нужна положительная площадь бетона."

    SP35ConcreteAreaForTableNult = concreteArea
    If totalSteelArea <= STABILITY_SMALL_VALUE Then Exit Function

    If totalSteelArea / concreteArea > STABILITY_SP35_REBAR_AREA_LIMIT_RATIO Then
        SP35ConcreteAreaForTableNult = concreteArea - totalSteelArea
        If SP35ConcreteAreaForTableNult <= STABILITY_SMALL_VALUE Then _
            StabilityRaiseInput "Суммарная площадь арматуры не меньше общей площади сечения до вычитания арматуры. Формула Nult СП 35 неприменима; проверьте геометрию, диаметры и единицы."
    End If
End Function

' Рассчитывает ветвь СП 35 для |ec|/r > 1: phi_l по моментам относительно
' требуемой грани, delta, D, Ncr, eta и запас по условию N <= k*Ncr.
Private Sub CalculateSP35EtaPlane(ByRef normBranch As String, ByVal nCompression As Double, ByVal sustainedCompression As Double, _
        ByVal eb As Double, ByVal es As Double, _
        ByVal rb As Double, ByVal phiLMode As String, ByVal phiP As Double, _
        ByVal nOverNcrLimit As Double, ByRef plane As TStabilityPlane)
    Dim faceProjection As Double
    faceProjection = StabilityConcreteFaceProjectionForMoment(plane)
    plane.PhiL = PhiLSP35Eta(phiLMode, StabilityMomentAboutProjection(plane.Moment, nCompression, faceProjection), _
        StabilityMomentAboutProjection(plane.SustainedMoment, sustainedCompression, faceProjection), _
        plane.Eccentricity, plane.Depth)

    Dim deltaValue As Double
    deltaValue = Abs(plane.Eccentricity) / plane.Depth
    Dim deltaMin As Double
    deltaMin = 0.5 - 0.01 * (plane.EffectiveLength / plane.Depth) - 0.01 * rb
    deltaValue = GeomMax(deltaValue, deltaMin)

    Dim phiPValue As Double
    phiPValue = phiP
    plane.PhiP = phiPValue
    Dim stiffnessFactor As Double
    stiffnessFactor = 0.11 / (0.1 + deltaValue / phiPValue) + 0.1
    plane.DeltaValue = deltaValue
    plane.StiffnessD = 6.4 * eb * (plane.InertiaConcrete / plane.PhiL * stiffnessFactor + _
        (es / eb) * plane.InertiaSteel)
    plane.NcrForce = plane.StiffnessD / (plane.EffectiveLength * plane.EffectiveLength)
    plane.CriticalForce = plane.NcrForce
    plane.ReserveFactor = SP35EtaReserveByLimit(plane.NcrForce, nCompression, nOverNcrLimit)
    plane.Eta = StabilityEtaByCriticalForce(plane.NcrForce, nCompression)
    plane.DesignMoment = StabilityDesignMomentWithEta(nCompression, plane.Eccentricity, plane.Eta)
    plane.CheckPassed = (plane.ReserveFactor + STABILITY_PASS_TOLERANCE >= 1# And plane.Eta > 0#)
    normBranch = "SP35-eta"
    plane.BranchName = "SP35-eta"
End Sub

' Вычисляет запас для условия СП 35 N/Ncr <= Stability.SP35.NOverNcrLimit.
' В Results показывается запас limit*Ncr/N, соответствующий
' проверяемому условию, а не простое отношение Ncr/N.
Private Function SP35EtaReserveByLimit(ByVal ncrForce As Double, ByVal nCompression As Double, _
        ByVal nOverNcrLimit As Double) As Double
    If nOverNcrLimit <= 0# Then Exit Function
    SP35EtaReserveByLimit = StabilityReserveByCriticalForce(ncrForce, nCompression) * nOverNcrLimit
End Function

' Для табличной ветви СП 35 учитывает долю постоянной нагрузки в полном
' сжимающем усилии по формуле (7.15). Если длительная часть равна нулю,
' остается кратковременный коэффициент phi_m; если вся сила длительная,
' результат переходит к phi_l.
Private Function PhiForSustainedCompression(ByVal phiM As Double, ByVal phiL As Double, _
        ByVal nCompression As Double, ByVal sustainedCompression As Double) As Double
    If phiM <= 0# Then Exit Function
    If phiL <= 0# Then
        PhiForSustainedCompression = phiM
        Exit Function
    End If
    If nCompression <= STABILITY_SMALL_FORCE Then
        PhiForSustainedCompression = phiM
        Exit Function
    End If

    Dim sustainedRatio As Double
    sustainedRatio = sustainedCompression / nCompression ' Доля длительного сжатия в полном; используется только в табличной ветви СП 35.

    Dim shortRatio As Double
    shortRatio = 1# - sustainedRatio

    Dim denominator As Double
    denominator = sustainedRatio * (phiM / phiL) + shortRatio ' Смешиваем кратковременное и длительное влияние по долям силы, а не линейно по самим коэффициентам phi.
    If denominator > STABILITY_SMALL_VALUE Then
        PhiForSustainedCompression = phiM / denominator
    Else
        StabilityRaiseInput "Для СП 35 получен некорректный знаменатель при учете постоянной/длительной продольной силы."
    End If
End Function

' Возвращает phi_l для ветви eta/Ncr по СП 35. В Auto обычно используется
' отношение длительного момента к полному, но при противоположных знаках
' полного и длительного моментов применяются специальные правила п. 7.54.
Private Function PhiLSP35Eta(ByVal phiLMode As String, ByVal totalMoment As Double, _
        ByVal sustainedMoment As Double, ByVal eccentricity As Double, ByVal depthValue As Double) As Double
    If UCase$(Trim$(phiLMode)) = UCase$(STABILITY_PHIL_2) Then
        PhiLSP35Eta = 2#
        Exit Function
    End If

    If Abs(totalMoment) > STABILITY_SMALL_VALUE And Abs(sustainedMoment) > STABILITY_SMALL_VALUE Then
        If Sgn(totalMoment) <> Sgn(sustainedMoment) And depthValue > STABILITY_SMALL_VALUE Then
            If Abs(eccentricity) >= 0.1 * depthValue Then
                PhiLSP35Eta = 1#
            Else
                PhiLSP35Eta = 1.05
            End If
            Exit Function
        End If
    End If

    PhiLSP35Eta = StabilityPhiLByMode(phiLMode, totalMoment, sustainedMoment)
End Function

' Выбирает ядровое расстояние с той стороны сечения, куда направлен
' расчетный эксцентриситет. Для несимметричного сечения положительная и
' отрицательная стороны могут иметь разные r.
Private Function CoreDistanceForEccentricity(ByRef plane As TStabilityPlane) As Double
    If plane.Eccentricity < 0# Then
        CoreDistanceForEccentricity = plane.CoreDistanceNegative
    Else
        CoreDistanceForEccentricity = plane.CoreDistancePositive
    End If
End Function

' Интерполирует таблицу 7.21 СП 35, переданную с листа Config. В книге
' таблица содержит l0/i, четыре phi_m для |ec|/r = 0; 0.25; 0.5; 1 и phi_l.
Private Sub InterpolateSP35Table(ByVal normTableInputError As String, ByVal table721 As Variant, ByVal slenderness As Double, ByVal q As Double, _
        ByRef phiM As Double, ByRef phiL As Double)
    If Len(normTableInputError) > 0 Then StabilityRaiseInput normTableInputError

    Dim invalidRow As Long, invalidColumn As Long, reason As String
    If Not TryValidateSP35Table721(table721, invalidRow, invalidColumn, reason) Then
        If invalidRow > 0 Then reason = "Строка " & CStr(invalidRow) & ", колонка " & _
            CStr(invalidColumn) & ": " & reason
        StabilityRaiseInput "Таблица 7.21 СП 35: " & reason
    End If

    Dim rowCount As Long, numericRow As Long, numericColumn As Long, numericValue As Double
    rowCount = UBound(table721, 1)
    For numericRow = 2 To rowCount
        For numericColumn = 1 To 6
            If Not TryReadStrictDouble(table721(numericRow, numericColumn), numericValue) Then StabilityRaiseInput "Таблица 7.21 СП 35 содержит неверное число."
            table721(numericRow, numericColumn) = numericValue ' Только локальная копия проверенной таблицы: интерполяция получает однозначные Double независимо от написания десятичного знака.
        Next numericColumn
    Next numericRow

    Dim lowerRow As Long
    Dim upperRow As Long
    FindBoundingRows table721, rowCount, slenderness, lowerRow, upperRow

    Dim phiMLower As Double
    Dim phiMUpper As Double
    Dim phiLLower As Double
    Dim phiLUpper As Double
    phiMLower = InterpolatePhiMByQ(table721, lowerRow, q)
    phiMUpper = InterpolatePhiMByQ(table721, upperRow, q)
    phiLLower = CDbl(table721(lowerRow, 6))
    phiLUpper = CDbl(table721(upperRow, 6))

    Dim t As Double
    If lowerRow = upperRow Then
        t = 0#
    Else
        t = (slenderness - CDbl(table721(lowerRow, 1))) / _
            (CDbl(table721(upperRow, 1)) - CDbl(table721(lowerRow, 1)))
    End If
    phiM = phiMLower + t * (phiMUpper - phiMLower)
    phiL = phiLLower + t * (phiLUpper - phiLLower)
End Sub

' Проверяет расчетные колонки таблицы без обращения к Excel и без скрытых
' нормативных defaults. Возвращает первую неверную строку/колонку для адресной
' диагностики вызывающего слоя. Гибкость неотрицательна; коэффициенты положительны.
' Положительность phi и возрастающие узлы защищают интерполяцию и предельную
' силу; дополнительные ограничения вроде phi <= 1 здесь не навязываются.
Friend Function TryValidateSP35Table721(ByVal table721 As Variant, ByRef invalidRow As Long, _
        ByRef invalidColumn As Long, ByRef reason As String) As Boolean
    invalidRow = 0: invalidColumn = 0: reason = vbNullString
    If Not IsArray(table721) Then
        reason = "Не передана таблица. Восстановите именованный диапазон rngSP35Table721 на Config."
        Exit Function
    End If
    On Error GoTo InvalidStructure
    Dim rowCount As Long, columnCount As Long, firstValue As Variant
    rowCount = UBound(table721, 1): columnCount = UBound(table721, 2)
    If LBound(table721, 1) <> 1 Or LBound(table721, 2) <> 1 Or _
            rowCount < 2 Or columnCount <> 6 Then GoTo InvalidStructure
    firstValue = table721(2, 1)
    On Error GoTo 0

    Dim row As Long, column As Long, value As Double, previousNode As Double
    For row = 2 To rowCount
        For column = 1 To 6
            invalidRow = row: invalidColumn = column
            If Not SP35TableNumber(table721(row, column), value, column = 1) Then
                If column = 1 Then
                    reason = "Гибкость l0/i должна быть числом не меньше нуля. Исправьте значение в таблице."
                Else
                    reason = "Коэффициент " & SP35TableColumnName(column) & _
                        " должен быть числом больше нуля. Заполните ячейку корректным коэффициентом."
                End If
                Exit Function
            End If
            If column = 1 Then
                If row > 2 And value <= previousNode Then
                    reason = "Узлы l0/i должны строго возрастать сверху вниз. Исправьте эту строку или предыдущий узел."
                    Exit Function
                End If
                previousNode = value
            End If
        Next column
    Next row
    invalidRow = 0: invalidColumn = 0
    TryValidateSP35Table721 = True
    Exit Function
InvalidStructure:
    invalidRow = 0: invalidColumn = 0
    reason = "Нужен единый прямоугольный диапазон с заголовком, строками данных и шестью колонками: " & _
        "l0/i, phi_m(0), phi_m(0.25), phi_m(0.5), phi_m(1), phi_l. Восстановите структуру таблицы."
End Function

' Отличает пустоту, ошибку формулы и нечисловое значение от допустимого
' числа: ноль допускается только для узла гибкости. Числовые строки допустимы;
' логические значения и даты не трактуются как инженерные числа.
Private Function SP35TableNumber(ByVal rawValue As Variant, ByRef value As Double, ByVal allowZero As Boolean) As Boolean
    On Error GoTo InvalidValue
    If IsError(rawValue) Or IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    If IsObject(rawValue) Or IsArray(rawValue) Then Exit Function
    If VarType(rawValue) = vbBoolean Or VarType(rawValue) = vbDate Then Exit Function
    If Not TryReadStrictDouble(rawValue, value) Then Exit Function
    SP35TableNumber = (value > 0# Or (allowZero And value = 0#))
    Exit Function
InvalidValue:
End Function

' Называет неверный коэффициент по заголовку таблицы, чтобы сообщение
' оставалось понятным и при автономном расчете без Excel-адреса.
Private Function SP35TableColumnName(ByVal column As Long) As String
    Select Case column
        Case 2: SP35TableColumnName = "phi_m при ec/r = 0"
        Case 3: SP35TableColumnName = "phi_m при ec/r = 0.25"
        Case 4: SP35TableColumnName = "phi_m при ec/r = 0.5"
        Case 5: SP35TableColumnName = "phi_m при ec/r = 1"
        Case 6: SP35TableColumnName = "phi_l"
    End Select
End Function

' Находит соседние строки таблицы по гибкости.
Private Sub FindBoundingRows(ByVal tableData As Variant, ByVal rowCount As Long, ByVal slenderness As Double, _
        ByRef lowerRow As Long, ByRef upperRow As Long)
    lowerRow = 2
    upperRow = rowCount
    Dim r As Long
    For r = 2 To rowCount
        If CDbl(tableData(r, 1)) <= slenderness Then lowerRow = r
        If CDbl(tableData(r, 1)) >= slenderness Then
            upperRow = r
            Exit For
        End If
    Next r
    If slenderness <= CDbl(tableData(2, 1)) Then lowerRow = 2: upperRow = 2
    If slenderness >= CDbl(tableData(rowCount, 1)) Then lowerRow = rowCount: upperRow = rowCount
End Sub

' Интерполяция phi_m по отношению |ec|/r.
Private Function InterpolatePhiMByQ(ByVal tableData As Variant, ByVal rowIndex As Long, ByVal q As Double) As Double
    Dim qValues(1 To 4) As Double
    qValues(1) = 0#: qValues(2) = 0.25: qValues(3) = 0.5: qValues(4) = 1#
    If q <= qValues(1) Then
        InterpolatePhiMByQ = CDbl(tableData(rowIndex, 2))
        Exit Function
    End If
    If q >= qValues(4) Then
        InterpolatePhiMByQ = CDbl(tableData(rowIndex, 5))
        Exit Function
    End If

    Dim i As Long
    For i = 1 To 3
        If q >= qValues(i) And q <= qValues(i + 1) Then
            Dim t As Double
            t = (q - qValues(i)) / (qValues(i + 1) - qValues(i))
            InterpolatePhiMByQ = CDbl(tableData(rowIndex, i + 1)) + _
                t * (CDbl(tableData(rowIndex, i + 2)) - CDbl(tableData(rowIndex, i + 1)))
            Exit Function
        End If
    Next i
End Function
