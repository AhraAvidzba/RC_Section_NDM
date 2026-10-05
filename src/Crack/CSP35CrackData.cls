VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSP35CrackData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подготовленные группы, ряды и области взаимодействия СП 35
' ==========================================================================
' По готовой плоскости деформаций формирует связные группы растянутой арматуры
' и численные входы каждого якорного расчета. Нормативные границы задаются
' здесь, а пересечения, пустоты и площадь выполняет CSectionGeometryQuery.
' CCrackWidthCalculator применяет к готовым числам чистые формулы и выбирает
' максимум. Здесь нет solver/search, Excel/AutoCAD и собственной status-policy.
' Хранятся реальные стержни всех рядов, а не условный диаметр вместо beta*nd.

Private Const LENGTH_TOL As Double = 0.000001 ' Только геометрическая погрешность, мм; не RowTolerance.
Private Const STRAIN_TOL As Double = 0.000000000001 ' Внутренний численный ноль деформации.
Private Const STRESS_TOL As Double = 0.000000001 ' Нулевая растягивающая арматура, МПа.
Private mSection As CSectionModel
Private mQuery As CSectionGeometryQuery
Private mDomain As CConcreteRegion
Private mGroupCount As Long
Private mBarGroup() As Long ' 0 - стержень не растянут; иначе неделимая связная группа.
Private mBarStress() As Double ' Фактические напряжения заданного НДС, МПа.
Private mBarStrain() As Double, mBarDiameter() As Double, mBarArea() As Double
Private mBarID() As String ' Собственные копии для результата, не чтение измененной модели позже.
Private mGroupX() As Double, mGroupY() As Double ' Центр по реальным площадям As.
Private mGroupN() As Double, mGroupT() As Double ' Проекции по нормали к Н.О. и вдоль нее.
Private mGroupArea() As Double, mGroupBars() As Long
Private mGroupBeta() As Double ' Сохраненный коэффициент группы для формулы и пассивного вывода.
Private mGroupRow() As Long ' Глобальный ряд нужен только для боковых соседей якоря.
Private mCandidateCount As Long
Private mZones() As CConcreteRegion ' Снимки той же геометрии, площадь которой пойдет в формулу.
Private mRows() As Variant ' Для кандидата массив [группа, локальный ряд]; только включенные группы.
Private mAnchorGroup() As Long, mStressBar() As Long
Private mBetaDiameterSum() As Double
Private mSideDiameter() As Double, mRowDiameter() As Double
Private mSideRadius() As Double, mNormalRadius() As Double
Private mReferenceRow() As Long, mRowCount() As Long
Private mNx As Double, mNy As Double, mNeutralN As Double
Private mGapTolerance As Double, mRowTolerance As Double
Private mDiameterMode As String, mRadiusMode As String
Private mCentralTension As Boolean
Private mFrozen As Boolean ' Опубликованный снимок нельзя повторно подготовить или изменить.
Private mSectionRevision As Long ' Версия исходной геометрии, зафиксированная при подготовке.
Private mCandidateRadius() As Double, mCandidatePsi() As Double, mCandidateWidth() As Double

' Формирует полный неизменяемый снимок подготовки одного сочетания. Плоскость
' и диаграмма уже получены владельцем НДС; метод не ищет равновесие. При чистом
' центральном растяжении создается одна общая область без искусственной Н.О.
Public Sub Prepare(ByVal section As CSectionModel, ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal steelMaterial As Object, ByVal groupGapTolerance As Double, ByVal rowTolerance As Double, _
        ByVal diameterMode As String, ByVal radiusMode As String, ByVal centralTension As Boolean)
    AssertWritable
    ResetData
    If section Is Nothing Then InvalidData "Не передана модель сечения для СП 35."
    If steelMaterial Is Nothing Then InvalidData "Не передана диаграмма арматуры текущего НДС."
    If groupGapTolerance < 0# Or rowTolerance < 0# Then InvalidData "Допуски групп и рядов СП 35 должны быть неотрицательными."
    Select Case diameterMode
        Case "Max", "Min", "Average"
        Case Else: InvalidData "RadiusDiameterMode СП 35 должен быть Max, Min или Average."
    End Select
    Dim formula As CCrackWidthCalculator, multiplier As Double
    Set formula = New CCrackWidthCalculator: multiplier = formula.SP35InteractionMultiplierFromData(radiusMode)
    Set mSection = section: mSectionRevision = section.Revision
    Set mQuery = New CSectionGeometryQuery: mQuery.Initialize section
    Set mDomain = mQuery.ConcreteDomain
    mGapTolerance = groupGapTolerance: mRowTolerance = rowTolerance
    mDiameterMode = diameterMode: mRadiusMode = radiusMode: mCentralTension = centralTension
    If section.RebarCount = 0 Then Exit Sub
    Dim gradient As Double
    gradient = Sqr(kappaX * kappaX + kappaY * kappaY)
    If Not centralTension Then
        If gradient <= STRAIN_TOL Then Exit Sub
        mNx = kappaY / gradient: mNy = kappaX / gradient: mNeutralN = -epsilon0 / gradient
    ElseIf epsilon0 <= STRAIN_TOL Then
        InvalidData "Для центрального растяжения СП 35 должна быть положительная деформация текущего НДС."
    End If
    BuildGroups epsilon0, kappaX, kappaY, steelMaterial
    If mGroupCount = 0 Then Exit Sub
    Dim included() As Boolean, rows() As Long, rowCount As Long, i As Long
    ReDim included(1 To mGroupCount): ReDim mGroupRow(1 To mGroupCount)
    ReDim mGroupBeta(1 To mGroupCount)
    For i = 1 To mGroupCount
        included(i) = True: mGroupBeta(i) = formula.SP35GroupBetaFromData(mGroupBars(i))
    Next i
    AssignRows included, rows, rowCount
    For i = 1 To mGroupCount: mGroupRow(i) = rows(i): Next i
    mCandidateCount = mGroupCount
    If centralTension Then mCandidateCount = 1
    ReDim mZones(1 To mCandidateCount): ReDim mRows(1 To mCandidateCount)
    ReDim mAnchorGroup(1 To mCandidateCount): ReDim mStressBar(1 To mCandidateCount)
    ReDim mBetaDiameterSum(1 To mCandidateCount)
    ReDim mSideDiameter(1 To mCandidateCount): ReDim mRowDiameter(1 To mCandidateCount)
    ReDim mSideRadius(1 To mCandidateCount): ReDim mNormalRadius(1 To mCandidateCount)
    ReDim mReferenceRow(1 To mCandidateCount): ReDim mRowCount(1 To mCandidateCount)
    ReDim mCandidateRadius(1 To mCandidateCount): ReDim mCandidatePsi(1 To mCandidateCount): ReDim mCandidateWidth(1 To mCandidateCount)
    If centralTension Then
        Set mZones(1) = mDomain.Clone
        mStressBar(1) = MaximumStressBar(0): mAnchorGroup(1) = mBarGroup(mStressBar(1))
        StoreIncludedGroups 1, included, rows, rowCount, formula
    Else
        ' [NDM] Надежный общий fallback: каждый растянутый кластер - кандидат.
        ' Не создаем отдельный классификатор наружного ряда для каждой формы.
        For i = 1 To mCandidateCount: PrepareAnchor i, multiplier, formula: Next i
    End If
End Sub

' Сбрасывает прежнее сочетание до валидации нового, чтобы ошибочный запрос
' не мог выдать область или стержни прошлого расчета как актуальный результат.
Private Sub ResetData()
    Set mSection = Nothing: Set mQuery = Nothing: Set mDomain = Nothing
    mGroupCount = 0: mCandidateCount = 0
    mSectionRevision = 0
    Erase mCandidateRadius: Erase mCandidatePsi: Erase mCandidateWidth
    Erase mGroupBeta
    Erase mBarGroup: Erase mBarStress: Erase mBarStrain: Erase mBarDiameter: Erase mBarArea: Erase mBarID
    Erase mGroupX: Erase mGroupY
    Erase mGroupN: Erase mGroupT: Erase mGroupArea: Erase mGroupBars: Erase mGroupRow
    Erase mZones: Erase mRows: Erase mAnchorGroup: Erase mStressBar: Erase mBetaDiameterSum
    Erase mSideDiameter: Erase mRowDiameter: Erase mSideRadius: Erase mNormalRadius
    Erase mReferenceRow: Erase mRowCount
    mNx = 0#: mNy = 0#: mNeutralN = 0#: mCentralTension = False
End Sub

' Строит связные компоненты по чистому зазору, включая транзитивную A-B-C
' связь. Сжатые/нулевые стержни не соединяют растянутые группы. Идентичность
' группы определяется стабильным порядком элементов модели, а не CAD handle.
Private Sub BuildGroups(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, ByVal steelMaterial As Object)
    Dim count As Long, parent() As Long, i As Long, j As Long, root As Long, otherRoot As Long
    Dim strain As Double, gap As Double, groups As Object, group As Long
    count = mSection.RebarCount
    ReDim parent(1 To count): ReDim mBarGroup(1 To count): ReDim mBarStress(1 To count)
    ReDim mBarStrain(1 To count): ReDim mBarDiameter(1 To count): ReDim mBarArea(1 To count): ReDim mBarID(1 To count)
    For i = 1 To count
        mBarID(i) = mSection.RebarID(i): mBarDiameter(i) = mSection.RebarDiameter(i): mBarArea(i) = mSection.RebarArea(i)
        strain = epsilon0 + kappaX * mSection.RebarY(i) + kappaY * mSection.RebarX(i)
        mBarStrain(i) = strain
        If strain > STRAIN_TOL Then mBarStress(i) = steelMaterial.GetStress(strain)
        If mBarStress(i) > STRESS_TOL Then
            If Not mQuery.ContainsPoint(mDomain, mSection.RebarX(i), mSection.RebarY(i)) Then _
                InvalidData "Растянутый стержень " & mSection.RebarID(i) & " находится вне бетонной области или в проеме."
            parent(i) = i
        End If
    Next i
    For i = 1 To count
        If parent(i) > 0 Then
            For j = i + 1 To count
                If parent(j) > 0 Then
                    gap = Sqr((mSection.RebarX(i) - mSection.RebarX(j)) ^ 2 + (mSection.RebarY(i) - mSection.RebarY(j)) ^ 2) - _
                        (mSection.RebarDiameter(i) + mSection.RebarDiameter(j)) / 2#
                    If gap <= mGapTolerance + LENGTH_TOL Then
                        root = GroupRoot(parent, i): otherRoot = GroupRoot(parent, j)
                        If root < otherRoot Then parent(otherRoot) = root Else parent(root) = otherRoot
                    End If
                End If
            Next j
        End If
    Next i
    Set groups = CreateObject("Scripting.Dictionary")
    For i = 1 To count
        If parent(i) > 0 Then
            root = GroupRoot(parent, i)
            If Not groups.Exists(CStr(root)) Then groups.Add CStr(root), groups.Count + 1
            mBarGroup(i) = CLng(groups(CStr(root)))
        End If
    Next i
    mGroupCount = groups.Count
    If mGroupCount = 0 Then Exit Sub
    ReDim mGroupX(1 To mGroupCount): ReDim mGroupY(1 To mGroupCount)
    ReDim mGroupN(1 To mGroupCount): ReDim mGroupT(1 To mGroupCount)
    ReDim mGroupArea(1 To mGroupCount): ReDim mGroupBars(1 To mGroupCount)
    For i = 1 To count
        group = mBarGroup(i)
        If group > 0 Then
            mGroupArea(group) = mGroupArea(group) + mSection.RebarArea(i)
            mGroupBars(group) = mGroupBars(group) + 1
            mGroupX(group) = mGroupX(group) + mSection.RebarArea(i) * mSection.RebarX(i)
            mGroupY(group) = mGroupY(group) + mSection.RebarArea(i) * mSection.RebarY(i)
        End If
    Next i
    For group = 1 To mGroupCount
        mGroupX(group) = mGroupX(group) / mGroupArea(group): mGroupY(group) = mGroupY(group) / mGroupArea(group)
        mGroupN(group) = mNx * mGroupX(group) + mNy * mGroupY(group)
        mGroupT(group) = -mNy * mGroupX(group) + mNx * mGroupY(group)
        If Not mQuery.ContainsPoint(mDomain, mGroupX(group), mGroupY(group)) Then _
            InvalidData "Центр растянутой группы G" & CStr(group) & " попал в проем или вне бетона. Проверьте геометрию и GroupGapTolerance СП 35."
    Next group
End Sub

' Находит представителя связной компоненты и сокращает цепочку parent.
' Этот численный граф относится к арматурной подготовке, не к solver retries.
Private Function GroupRoot(ByRef parent() As Long, ByVal index As Long) As Long
    Dim root As Long, nextIndex As Long
    root = index
    Do While parent(root) <> root: root = parent(root): Loop
    Do While parent(index) <> index
        nextIndex = parent(index): parent(index) = root: index = nextIndex
    Loop
    GroupRoot = root
End Function

' Назначает ряды от Н.О. наружу. Порядок стабилен по n/t/ID; полный размах
' координат одного ряда не превышает RowTolerance, без цепочного слияния
' нескольких далеких рядов. При центральном растяжении Н.О. нет: один ряд.
Private Sub AssignRows(ByRef included() As Boolean, ByRef rows() As Long, ByRef rowCount As Long)
    Dim order() As Long, count As Long, i As Long, j As Long, group As Long, firstN As Double
    ReDim rows(1 To mGroupCount): ReDim order(1 To mGroupCount): rowCount = 0
    For i = 1 To mGroupCount
        If included(i) Then
            count = count + 1: order(count) = i
            For j = count To 2 Step -1
                If Not GroupBefore(order(j), order(j - 1)) Then Exit For
                group = order(j - 1): order(j - 1) = order(j): order(j) = group
            Next j
        End If
    Next i
    For i = 1 To count
        group = order(i)
        If rowCount = 0 Then
            rowCount = 1: firstN = mGroupN(group)
        ElseIf Not mCentralTension And mGroupN(group) - firstN > mRowTolerance + LENGTH_TOL Then
            rowCount = rowCount + 1: firstN = mGroupN(group)
        End If
        rows(group) = rowCount
    Next i
End Sub

' Сравнивает центры для воспроизводимого порядка рядов и групп одного ряда.
Private Function GroupBefore(ByVal first As Long, ByVal second As Long) As Boolean
    If Abs(mGroupN(first) - mGroupN(second)) > LENGTH_TOL Then
        GroupBefore = (mGroupN(first) < mGroupN(second))
    ElseIf Abs(mGroupT(first) - mGroupT(second)) > LENGTH_TOL Then
        GroupBefore = (mGroupT(first) < mGroupT(second))
    Else
        GroupBefore = (first < second)
    End If
End Function

' Задает нормативное окно вокруг якоря и запрашивает общую геометрию.
' Предварительно исключает арматуру противоположной стенки, затем выбирает
' опорный ряд по реальным As и повторно проверяет центры в конечной области.
Private Sub PrepareAnchor(ByVal anchor As Long, ByVal multiplier As Double, ByVal formula As CCrackWidthCalculator)
    Dim wall As CConcreteRegion, region As CConcreteRegion, i As Long, left As Double, right As Double, distance As Double
    Dim included() As Boolean, rows() As Long, rowCount As Long, rowAreas() As Double, rowN() As Double
    Dim referenceRow As Long, referenceGroup As Long, incomplete As Boolean, lowerN As Double, minN As Double, maxN As Double
    mAnchorGroup(anchor) = anchor: mStressBar(anchor) = MaximumStressBar(anchor)
    Set wall = mQuery.MaterialRegionThroughPlane(mDomain, mNx, mNy, mGroupN(anchor))
    mSideDiameter(anchor) = ReferenceDiameterForGroup(anchor)
    mSideRadius(anchor) = multiplier * mSideDiameter(anchor)
    left = mSideRadius(anchor): right = left
    For i = 1 To mGroupCount
        If i <> anchor And mGroupRow(i) = mGroupRow(anchor) Then
            If mQuery.ContainsPoint(wall, mGroupX(i), mGroupY(i)) Then
                distance = mGroupT(i) - mGroupT(anchor)
                If distance < -LENGTH_TOL Then left = MinValue(left, -distance / 2#)
                If distance > LENGTH_TOL Then right = MinValue(right, distance / 2#)
            End If
        End If
    Next i
    ReDim included(1 To mGroupCount)
    For i = 1 To mGroupCount
        included(i) = (mGroupT(i) >= mGroupT(anchor) - left - LENGTH_TOL And _
            mGroupT(i) <= mGroupT(anchor) + right + LENGTH_TOL And _
            mQuery.ContainsPoint(wall, mGroupX(i), mGroupY(i)))
    Next i
    AssignRows included, rows, rowCount
    ReDim rowAreas(1 To rowCount): ReDim rowN(1 To rowCount)
    For i = 1 To mGroupCount
        If included(i) Then
            rowAreas(rows(i)) = rowAreas(rows(i)) + mGroupArea(i)
            rowN(rows(i)) = rowN(rows(i)) + mGroupArea(i) * mGroupN(i)
        End If
    Next i
    For i = 1 To rowCount: rowN(i) = rowN(i) / rowAreas(i): Next i
    referenceRow = 1: incomplete = (rowCount > 1)
    For i = 2 To rowCount
        If rowAreas(1) >= 0.5 * rowAreas(i) Then incomplete = False
    Next i
    If incomplete Then referenceRow = 2
    For i = 1 To mGroupCount
        If included(i) And rows(i) = referenceRow Then referenceGroup = i: Exit For
    Next i
    mRowDiameter(anchor) = ReferenceDiameterForRow(included, rows, referenceRow)
    mNormalRadius(anchor) = multiplier * mRowDiameter(anchor)
    lowerN = rowN(referenceRow) - mNormalRadius(anchor)
    If lowerN < mNeutralN Then lowerN = mNeutralN
    mQuery.ProjectionBounds mDomain, mNx, mNy, minN, maxN
    Set region = mQuery.ClipWindow(wall, mNx, mNy, lowerN, maxN, mGroupT(anchor) - left, mGroupT(anchor) + right)
    For i = 1 To mGroupCount
        If included(i) Then included(i) = mQuery.ContainsPoint(region, mGroupX(i), mGroupY(i))
    Next i
    ' Перебор всех групп включает внутренние кандидаты, которые после
    ' нормативного выбора полного ряда могут оказаться вне своей зоны.
    ' Такой кандидат не является расчетным якорем и не вызывает формулу;
    ' это не ошибка физического НДС и не причина отменять другие якоря.
    If Not included(anchor) Or region.Area <= LENGTH_TOL Then Exit Sub
    ' Ряд за окончательной внутренней границей больше не участвует ни в
    ' знаменателе, ни в нумерации расшифровки сохраненного результата.
    AssignRows included, rows, rowCount
    mReferenceRow(anchor) = rows(referenceGroup)
    ' Пять базовых пар P/Q сохраняются для проверки происхождения границ.
    ' Между ними область уже обрезана по всем настоящим граням и дугам:
    ' небольшое отверстие между пробами не исчезает и не меняет площадь.
    region.SetBoundaryProbes mQuery.SampleMaterialIntervals(region, mNx, mNy, mGroupN(anchor), _
        mGroupT(anchor) - left, mGroupT(anchor) + right, 5)
    Set mZones(anchor) = region
    StoreIncludedGroups anchor, included, rows, rowCount, formula
End Sub

' Сохраняет разреженную таблицу групп/рядов конкретной конечной области.
' beta применяется к реальной сумме диаметров; один d_ref ее не заменяет.
Private Sub StoreIncludedGroups(ByVal candidate As Long, ByRef included() As Boolean, ByRef rows() As Long, _
        ByVal rowCount As Long, ByVal formula As CCrackWidthCalculator)
    Dim count As Long, i As Long, j As Long, group As Long, data() As Variant
    For i = 1 To mGroupCount: If included(i) Then count = count + 1
    Next i
    ReDim data(1 To count, 1 To 2): count = 0
    For i = 1 To mGroupCount
        If included(i) Then
            count = count + 1: data(count, 1) = i: data(count, 2) = rows(i)
        End If
    Next i
    For i = 1 To mSection.RebarCount
        group = mBarGroup(i)
        If group > 0 Then
            If included(group) Then mBetaDiameterSum(candidate) = mBetaDiameterSum(candidate) + _
                mGroupBeta(group) * mSection.RebarDiameter(i)
        End If
    Next i
    mRows(candidate) = data: mRowCount(candidate) = rowCount
End Sub

' Выбирает d_ref только для построения радиуса отдельной группы.
Private Function ReferenceDiameterForGroup(ByVal group As Long) As Double
    Dim i As Long, count As Long, total As Double, minimum As Double, maximum As Double, diameter As Double
    minimum = 1E+100
    For i = 1 To mSection.RebarCount
        If mBarGroup(i) = group Then
            diameter = mSection.RebarDiameter(i): count = count + 1: total = total + diameter
            minimum = MinValue(minimum, diameter): If diameter > maximum Then maximum = diameter
        End If
    Next i
    ReferenceDiameterForGroup = SelectedDiameter(minimum, maximum, total, count)
End Function

' Выбирает d_ref из всех фактических стержней нормативного опорного ряда.
' Режим Max/Min/Average тот же, что для бокового радиуса якорной группы.
Private Function ReferenceDiameterForRow(ByRef included() As Boolean, ByRef rows() As Long, ByVal row As Long) As Double
    Dim i As Long, group As Long, count As Long, total As Double, minimum As Double, maximum As Double, diameter As Double
    minimum = 1E+100
    For i = 1 To mSection.RebarCount
        group = mBarGroup(i)
        If group > 0 Then
            If included(group) And rows(group) = row Then
                diameter = mSection.RebarDiameter(i): count = count + 1: total = total + diameter
                minimum = MinValue(minimum, diameter): If diameter > maximum Then maximum = diameter
            End If
        End If
    Next i
    ReferenceDiameterForRow = SelectedDiameter(minimum, maximum, total, count)
End Function

' Применяет согласованное программное правило выбора d_ref; только Average
' усредняет диаметры арифметически, не по площади и не через эквивалентный d.
Private Function SelectedDiameter(ByVal minimum As Double, ByVal maximum As Double, ByVal total As Double, ByVal count As Long) As Double
    If count = 0 Then InvalidData "В опорном ряду СП 35 отсутствует арматура."
    Select Case mDiameterMode
        Case "Min": SelectedDiameter = minimum
        Case "Max": SelectedDiameter = maximum
        Case "Average": SelectedDiameter = total / count
    End Select
End Function

' Возвращает фактический стержень с максимальным растяжением якорной группы.
' group=0 используется только для всего сечения при центральном растяжении.
Private Function MaximumStressBar(ByVal group As Long) As Long
    Dim i As Long, maximum As Double
    For i = 1 To mSection.RebarCount
        If mBarGroup(i) > 0 And (group = 0 Or mBarGroup(i) = group) Then
            If mBarStress(i) > maximum Then maximum = mBarStress(i): MaximumStressBar = i
        End If
    Next i
End Function

' Простые read-only свойства возвращают подготовленные данные без расчетов.
Public Property Get CandidateCount() As Long
    CandidateCount = mCandidateCount
End Property
Public Property Get GroupCount() As Long
    GroupCount = mGroupCount
End Property
Public Property Get CandidateRegion(ByVal index As Long) As CConcreteRegion
    CheckCandidate index
    If Not mZones(index) Is Nothing Then Set CandidateRegion = mZones(index).Clone
End Property
Public Property Get CandidateAvailable(ByVal index As Long) As Boolean
    CheckCandidate index: CandidateAvailable = Not mZones(index) Is Nothing
End Property
Public Property Get CandidateRows(ByVal index As Long) As Variant
    CheckCandidate index: CandidateRows = mRows(index)
End Property
Public Property Get CandidateAnchorGroup(ByVal index As Long) As Long
    CheckCandidate index: CandidateAnchorGroup = mAnchorGroup(index)
End Property
Public Property Get CandidateStressBar(ByVal index As Long) As Long
    CheckCandidate index: CandidateStressBar = mStressBar(index)
End Property
Public Property Get CandidateSigmaS(ByVal index As Long) As Double
    CheckCandidate index: CandidateSigmaS = mBarStress(mStressBar(index))
End Property
Public Property Get CandidateBetaDiameterSum(ByVal index As Long) As Double
    CheckCandidate index: CandidateBetaDiameterSum = mBetaDiameterSum(index)
End Property
Public Property Get CandidateSideDiameter(ByVal index As Long) As Double
    CheckCandidate index: CandidateSideDiameter = mSideDiameter(index)
End Property
Public Property Get CandidateRowDiameter(ByVal index As Long) As Double
    CheckCandidate index: CandidateRowDiameter = mRowDiameter(index)
End Property
Public Property Get CandidateSideRadius(ByVal index As Long) As Double
    CheckCandidate index: CandidateSideRadius = mSideRadius(index)
End Property
Public Property Get CandidateNormalRadius(ByVal index As Long) As Double
    CheckCandidate index: CandidateNormalRadius = mNormalRadius(index)
End Property
Public Property Get CandidateReferenceRow(ByVal index As Long) As Long
    CheckCandidate index: CandidateReferenceRow = mReferenceRow(index)
End Property
Public Property Get CandidateRowCount(ByVal index As Long) As Long
    CheckCandidate index: CandidateRowCount = mRowCount(index)
End Property
Public Property Get BarGroup(ByVal index As Long) As Long
    BarGroup = mBarGroup(index)
End Property
Public Property Get GroupX(ByVal index As Long) As Double
    GroupX = mGroupX(index)
End Property
Public Property Get GroupY(ByVal index As Long) As Double
    GroupY = mGroupY(index)
End Property
Public Property Get GroupBarCount(ByVal index As Long) As Long
    GroupBarCount = mGroupBars(index)
End Property
Public Property Get GroupArea(ByVal index As Long) As Double
    GroupArea = mGroupArea(index)
End Property
Public Property Get GroupBeta(ByVal index As Long) As Double
    GroupBeta = mGroupBeta(index)
End Property
Public Property Get GroupRow(ByVal index As Long) As Long
    GroupRow = mGroupRow(index)
End Property
Public Property Get BarID(ByVal index As Long) As String
    BarID = mBarID(index)
End Property
Public Property Get BarDiameter(ByVal index As Long) As Double
    BarDiameter = mBarDiameter(index)
End Property
Public Property Get BarArea(ByVal index As Long) As Double
    BarArea = mBarArea(index)
End Property
Public Property Get BarStrain(ByVal index As Long) As Double
    BarStrain = mBarStrain(index)
End Property
Public Property Get NormalX() As Double
    NormalX = mNx
End Property
Public Property Get NormalY() As Double
    NormalY = mNy
End Property
Public Property Get NeutralProjection() As Double
    NeutralProjection = mNeutralN
End Property
Public Property Get GroupGapTolerance() As Double
    GroupGapTolerance = mGapTolerance
End Property
Public Property Get RowTolerance() As Double
    RowTolerance = mRowTolerance
End Property
Public Property Get RadiusDiameterMode() As String
    RadiusDiameterMode = mDiameterMode
End Property
Public Property Get InteractionRadiusMode() As String
    InteractionRadiusMode = mRadiusMode
End Property
Public Property Get CentralTension() As Boolean
    CentralTension = mCentralTension
End Property
Public Property Get SectionRevision() As Long
    SectionRevision = mSectionRevision
End Property
Public Property Get CandidateReinforcementRadius(ByVal index As Long) As Double
    CheckCandidate index: CandidateReinforcementRadius = mCandidateRadius(index)
End Property
Public Property Get CandidatePsiCm(ByVal index As Long) As Double
    CheckCandidate index: CandidatePsiCm = mCandidatePsi(index)
End Property
Public Property Get CandidateCrackWidth(ByVal index As Long) As Double
    CheckCandidate index: CandidateCrackWidth = mCandidateWidth(index)
End Property

' Принимает уже вычисленные чистой формулой числа для диагностики всех якорей.
' Класс не вычисляет раскрытие и не назначает инженерный статус результата.
Friend Sub StoreFormulaResult(ByVal index As Long, ByVal radius As Double, ByVal psiCm As Double, ByVal width As Double)
    AssertWritable
    CheckCandidate index
    mCandidateRadius(index) = radius: mCandidatePsi(index) = psiCm: mCandidateWidth(index) = width
End Sub

' Закрывает заполнение снимка перед передачей в result-tree. Геометрия при
' чтении по-прежнему возвращается копией; стержни и настройки уже скопированы.
Friend Sub Freeze()
    mFrozen = True
End Sub

' Запрещает изменение подготовленных данных, уже использованных в результате.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 5511, "CSP35CrackData", _
        "Опубликованные данные СП 35 нельзя изменить; для нового сочетания создайте новый объект."
End Sub

' Возвращает копию списка реальных стержней неделимой группы. Writer может
' расшифровать ряды, beta/n/d/rebars, не теряя смешанные номинальные диаметры.
Public Function GroupMembers(ByVal group As Long) As Variant
    If group < 1 Or group > mGroupCount Then InvalidData "Номер растянутой группы СП 35 вне допустимого диапазона."
    Dim members() As Long, i As Long, count As Long
    ReDim members(1 To mGroupBars(group))
    For i = 1 To UBound(mBarGroup)
        If mBarGroup(i) = group Then count = count + 1: members(count) = i
    Next i
    GroupMembers = members
End Function

' Проверяет чтение готового кандидата, чтобы не вернуть числа соседнего LC.
Private Sub CheckCandidate(ByVal index As Long)
    If index < 1 Or index > mCandidateCount Then InvalidData "Номер якорного расчета СП 35 вне допустимого диапазона."
End Sub

' Ошибка подготовленных входных данных передается существующему coordinator;
' внутренний/внешний статус и ResultComment назначает его штатный контракт.
Private Sub InvalidData(ByVal reason As String)
    Err.Raise vbObjectError + 5510, "CSP35CrackData", reason
End Sub

' Минимум двух геометрических ограничений без зависимости от Excel.
Private Function MinValue(ByVal first As Double, ByVal second As Double) As Double
    If first < second Then MinValue = first Else MinValue = second
End Function
