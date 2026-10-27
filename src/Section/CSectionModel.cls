VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Единая расчётная модель сечения
' ==========================================================================
' Хранит бетонные элементы, арматуру и точные наружные границы и отверстия.
' Генератор и импорт AutoCAD заполняют эту же модель, поэтому расчётное ядро
' получает одинаковые данные независимо от способа задания геометрии.
' Расчётные имена C1/R1 создаются только здесь. Размеры и подписи
' хранятся отдельно и не определяют расчётную геометрию.

Private mSourceType As String ' Диагностическая метка источника модели: Generated, AutoCADImport и т.п.
Private mConcreteMeshSource As String ' Способ получения бетонных элементов, отдельно от источника арматуры/контуров.

Private mConcreteElements() As CConcreteElement ' Канонические неизменяемые сущности; расчётные массивы ниже являются только числовой проекцией.
Private mConcreteX() As Double ' X центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteY() As Double ' Y центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteArea() As Double ' Площадь бетонного элемента, мм2.
Private mConcreteCount As Long ' Фактическое число бетонных элементов в модели.
Private mConcreteCapacity As Long ' Вместимость внутренних массивов; чтение и сохранение используют только заполненные элементы.

Private mBoundaryCacheValid() As Boolean ' True, если интерпретация оболочки элемента уже рассчитана и не требует повторения.
Private mBoundaryCacheStatus() As String ' Кэшированный GeometryInterpretationStatus: Rectangle, Equivalent rectangle или Equivalent square.
Private mBoundaryCacheWidth() As Double ' Кэшированная ширина интерпретированной геометрической оболочки, мм.
Private mBoundaryCacheHeight() As Double ' Кэшированная высота интерпретированной геометрической оболочки, мм.
Private mBoundaryCacheRotation() As Double ' Кэшированный угол интерпретированной геометрической оболочки, рад.

Private mRebarID() As String ' Расчетные имена стержней: R1, R2, ...
Private mRebarSourceName() As String ' Необязательное исходное имя арматурной области.
Private mRebarSourceHandle() As String ' Необязательный AutoCAD handle арматурного Region.
Private mRebarX() As Double ' X центра стержня, мм.
Private mRebarY() As Double ' Y центра стержня, мм.
Private mRebarArea() As Double ' Площадь стержня, мм2.
Private mRebarDiameter() As Double ' Диаметр стержня, мм.
Private mRebarSteelClass() As String ' Название материала, переданное при добавлении стержня.
Private mRebarMaterialID() As Long ' Номер материала стержня; в текущем НДС используется одна подготовленная диаграмма арматуры.
Private mRebarComment() As String ' Диагностическая заметка о грани, ряде или источнике стержня.
Private mRebarRecipeTrace() As String ' Необязательные сведения о правилах и исходных позициях, из которых получен стержень.
Private mGeometryDiagnostics As String ' Пояснения и предупреждения, полученные при подготовке геометрии.
Private mHasGeometryWarnings As Boolean ' Значимое событие текущего построения; пояснительный текст снимка его не восстанавливает.
Private mGeometryWarnings As String ' Готовая краткая сводка, независимая от подробного текста диагностики.
Private mRebarCount As Long ' Фактическое число стержней в модели.
Private mRebarCapacity As Long ' Резерв арматуры живет только вместе с этой моделью.

Private mAnnotations As CSectionAnnotations ' Готовые размеры, подписи арматуры и другие аннотации схемы.
Private mContours As CSectionContours ' Достоверные outer/opening, принадлежащие исходной геометрии.
Private mRevision As Long ' Версия геометрии: после её изменения ранее найденное НДС для этой модели использовать нельзя.
Private mObservedContourRevision As Long ' Последняя учтенная версия достоверного outer/opening.
Private mConcreteCenterValid As Boolean, mConcreteCenterRevision As Long ' Признак готовности и версия геометрии для сохранённого центра бетонной сетки.
Private mConcreteCenterX As Double, mConcreteCenterY As Double, mConcreteCenterArea As Double ' Координаты центра бетонной сетки в мм и её суммарная площадь в мм2.
Private mElementRadiusValid As Boolean, mElementRadiusRevision As Long, mElementRadius As Double ' Готовность, версия геометрии и сохранённый максимальный радиус размещения элементов, мм.
Private mForceIntegrator As CDiscreteForceIntegrator ' Индекс только этой модели; точная ревизия и центр проверяются при каждой пробе усилий.
Private mTestDisableForceIntegration As Boolean ' ДЛЯ ТЕСТОВ: прямой обход элементов для независимого сравнения того же solve.
Private mForceProbeCount As Long ' Число реальных проб усилий: короткий одиночный solve не оплачивает построение индекса.
Private mForceBatchCount As Long, mForceProbeRevision As Long ' Объём текущего запуска и ревизия счётчика проб.

Private Const LOCAL_INERTIA_SQUARE_REL_TOLERANCE As Double = 0.000001 ' Допуск, при котором главные локальные I считаются изотропными.
Private Const LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE As Double = 0.01 ' Допуск проверки, что реальные A/I совместимы с прямоугольником.

' Очищает модель перед новым построением сечения.
' После Clear объект снова готов принимать бетонные элементы, арматуру и
' аннотации без следов предыдущего расчета.
Public Sub Clear()
    Set mForceIntegrator = Nothing
    mForceProbeCount = 0
    mRevision = mRevision + 1
    mConcreteCenterValid = False
    mElementRadiusValid = False
    mSourceType = vbNullString
    mConcreteMeshSource = vbNullString

    Erase mConcreteElements
    Erase mConcreteX
    Erase mConcreteY
    Erase mConcreteArea
    mConcreteCount = 0
    mConcreteCapacity = 0
    ClearBoundaryCache

    Erase mRebarID
    Erase mRebarSourceName
    Erase mRebarSourceHandle
    Erase mRebarX
    Erase mRebarY
    Erase mRebarArea
    Erase mRebarDiameter
    Erase mRebarSteelClass
    Erase mRebarMaterialID
    Erase mRebarComment
    Erase mRebarRecipeTrace: mGeometryDiagnostics = vbNullString
    mHasGeometryWarnings = False: mGeometryWarnings = vbNullString
    mRebarCount = 0
    mRebarCapacity = 0

    Set mAnnotations = New CSectionAnnotations
    Set mContours = New CSectionContours
End Sub

' Для большой сетки переиспользует геометрический индекс только при точном совпадении
' ревизии и центра. Малые сетки не выделяют массивов и обходятся напрямую.
Friend Function TryConcreteForces(ByVal diagram As CMaterialDiagram, ByVal eps As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal originX As Double, ByVal originY As Double, ByRef n As Double, ByRef mx As Double, ByRef my As Double) As Boolean
    If mConcreteCount <= 5000 Then Exit Function
    Dim revision As Long: revision = Me.Revision
    If revision <> mForceProbeRevision Then
        mForceProbeRevision = revision: mForceProbeCount = 0: Set mForceIntegrator = Nothing
    End If
    mForceProbeCount = mForceProbeCount + 1
    If mTestDisableForceIntegration Then Exit Function
    If mForceBatchCount >= 4 Then
        If mForceProbeCount < 16 Then Exit Function
    Else
        If mForceProbeCount < 64 Then Exit Function ' Короткий одиночный Newton не оплачивает построение дерева в самом конце поиска.
    End If
    If mForceIntegrator Is Nothing Then Set mForceIntegrator = New CDiscreteForceIntegrator
    If Not mForceIntegrator.Matches(Me.Revision, originX, originY) Then mForceIntegrator.Initialize Me, originX, originY
    TryConcreteForces = mForceIntegrator.TryForces(diagram, eps, kx, ky, n, mx, my)
End Function

' Сообщает объём одного запуска до расчёта. Геометрический индекс переиспользуется
' только при совпадении ревизии; неготовая подготовка предыдущего запуска не навязывается одиночному solve.
Friend Sub BeginForceIntegrationRun(ByVal combinationCount As Long)
    mForceBatchCount = combinationCount
    If mForceIntegrator Is Nothing Then mForceProbeCount = 0
End Sub

' Диагностическая метка источника модели: Generated, AutoCADImport и т.п.
Public Property Let SourceType(ByVal value As String)
    mSourceType = Trim$(value)
End Property

' Диагностическая метка источника модели: Generated, AutoCADImport и т.п.
Public Property Get SourceType() As String
    SourceType = mSourceType
End Property

' Сведения об источнике сетки для Results; на решение равновесия они не влияют.
Public Property Let ConcreteMeshSource(ByVal value As String)
    Select Case value
        Case vbNullString, "Generated", "AutoCADRegions", "AutoCADContours"
            mConcreteMeshSource = value
        Case Else
            Err.Raise vbObjectError + 5203, "CSectionModel", "Неизвестный источник бетонной сетки: " & value & ". Повторите импорт или расчет."
    End Select
End Property

' Способ получения бетонных элементов, отдельно от источника арматуры/контуров.
Public Property Get ConcreteMeshSource() As String
    ConcreteMeshSource = mConcreteMeshSource
End Property

' Сохраняет сообщения проверки геометрии для записи в Results.
Public Property Let GeometryDiagnosticsText(ByVal value As String)
    mGeometryDiagnostics = value
End Property

' Возвращает сообщения проверки геометрии для записи в Results.
Public Property Get GeometryDiagnosticsText() As String
    GeometryDiagnosticsText = mGeometryDiagnostics
End Property

' Возвращает факт значимого события текущего построения, независимо от пояснения.
' Текстовое пояснение Results не является машинной классификацией события.
Public Property Get HasGeometryWarnings() As Boolean
    HasGeometryWarnings = mHasGeometryWarnings
End Property

' Возвращает готовую краткую сводку; подробная Geometry.RecipeDiagnostics сохраняется отдельно.
Public Property Get GeometryWarningText() As String
    GeometryWarningText = mGeometryWarnings
End Property

' Сохраняет итог исполнителя правил без разбора фраз; не меняет арматуру и ревизию геометрии.
Friend Sub SetGeometryWarnings(ByVal hasWarnings As Boolean, ByVal warningText As String)
    mHasGeometryWarnings = hasWarnings
    mGeometryWarnings = warningText
End Sub

' Добавляет связь стержня с правилом, линией и исходной позицией.
' Все связи одного стержня сохраняются без изменения его площади, инерций
' и версии геометрии. Разделители внутри имён кодируются, чтобы различать записи.
Public Sub AppendRebarRecipeTrace(ByVal index As Long, ByVal ruleID As String, ByVal pathID As String, ByVal sourcePositionID As String)
    CheckRebarIndex index
    If Len(mRebarRecipeTrace(index)) > 0 Then mRebarRecipeTrace(index) = mRebarRecipeTrace(index) & ";"
    mRebarRecipeTrace(index) = mRebarRecipeTrace(index) & "RuleId=" & TraceToken(ruleID) & _
        "|PathId=" & TraceToken(pathID) & "|SourcePositionId=" & TraceToken(sourcePositionID)
End Sub

' Возвращает служебные связи всех исходных позиций физического стержня.
' Эти сведения не заменяют его координаты, площадь или материал.
Public Property Get RebarRecipeTrace(ByVal index As Long) As String
    CheckRebarIndex index
    RebarRecipeTrace = mRebarRecipeTrace(index)
End Property

' Восстанавливает уже сохранённые связи физического стержня без повторного
' выполнения рецепта и без изменения его координат или площади.
Friend Sub RestoreRebarRecipeTrace(ByVal index As Long, ByVal value As String)
    CheckRebarIndex index
    mRebarRecipeTrace(index) = value
End Sub

' Заменяет служебные разделители в сведениях о происхождении стержня, чтобы отличать их от текста данных.
Private Function TraceToken(ByVal value As String) As String
    value = Replace$(value, "%", "%25"): value = Replace$(value, "|", "%7C"): value = Replace$(value, ";", "%3B")
    value = Replace$(value, "=", "%3D"): value = Replace$(value, vbCr, "%0D"): value = Replace$(value, vbLf, "%0A")
    TraceToken = value
End Function


' Возвращает неизменяемую сущность бетонного элемента; изменения модели выполняются её владельцем.
Public Function ConcreteElement(ByVal index As Long) As CConcreteElement
    CheckConcreteIndex index
    Set ConcreteElement = mConcreteElements(index)
End Function

' Защищает чтение точного тензора без размерного нижнего порога.
' Нормированное произведение избегает переполнения и сохраняет тонкие сегменты.
Private Function ExactPositiveInertia(ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double) As Boolean
    If localIx <= 0# Or localIy <= 0# Then Exit Function
    ExactPositiveInertia = Abs(localIxy / Sqr(localIx) / Sqr(localIy)) < 1#
End Function

' Добавляет одно бетонное волокно или импортированную бетонную область.
' Координаты, площадь и инерции заданы во внутренних единицах проекта.
' Contour требует границу и согласованные A/c/I до публикации сущности Cn:
' при отказе список, имена, ревизия и подготовленные кэши не изменяются.
' rotationKnown отличает известный нулевой угол грани от отсутствующего направления.
Public Sub AddConcreteElement(ByVal xCoord As Double, ByVal yCoord As Double, ByVal area As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal shapeType As String = "Rectangle", _
        Optional ByVal width As Double = 0#, Optional ByVal height As Double = 0#, _
        Optional ByVal rotation As Double = 0#, Optional ByVal comment As String = vbNullString, _
        Optional ByVal localIx As Double = 0#, Optional ByVal localIy As Double = 0#, _
        Optional ByVal localIxy As Double = 0#, Optional ByVal rotationKnown As Boolean = False, _
        Optional ByVal exactRegion As CConcreteRegion = Nothing)
    If area <= 0# Then Err.Raise vbObjectError + 5200, "CSectionModel", "Площадь бетонного элемента должна быть положительной."
    If Abs(rotation) > 2# * GEOM_PI Then rotation = NormalizeAnglePi(rotation)
    Dim explicitRectangle As Boolean
    explicitRectangle = StrComp(Trim$(shapeType), "Rectangle", vbTextCompare) = 0 And width > GEOM_TOLERANCE And height > GEOM_TOLERANCE And area > GEOM_TOLERANCE
    If explicitRectangle Then explicitRectangle = Abs(width * height - area) / area <= LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE
    Dim exactBoundary As Boolean
    exactBoundary = StrComp(Trim$(shapeType), "Contour", vbTextCompare) = 0 Or Not exactRegion Is Nothing
    Dim query As CSectionGeometryQuery
    If exactBoundary Then
        Set query = New CSectionGeometryQuery
        ValidateConcreteBoundary query, exactRegion, xCoord, yCoord, area, localIx, localIy, localIxy, "C" & CStr(mConcreteCount + 1)
        shapeType = "Contour": explicitRectangle = False ' Фактическая граница определяет тип независимо от переданной подсказки формы.
    ElseIf Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        If explicitRectangle Then RectangleLocalInertiaComponents area, width, height, rotation, localIx, localIy, localIxy
    End If
    Dim displayWidth As Double, displayHeight As Double, displayRotation As Double, statusText As String
    displayWidth = width: displayHeight = height: displayRotation = rotation
    If exactBoundary Then
        ConcreteRegionDisplay query, exactRegion, displayWidth, displayHeight, displayRotation
        rotationKnown = True
    ElseIf Not explicitRectangle Then
        Dim viewIx As Double, viewIy As Double, viewIxy As Double
        viewIx = localIx: viewIy = localIy: viewIxy = localIxy
        If Not IsPositiveDefiniteInertia(viewIx, viewIy, viewIxy) Then EquivalentSquareLocalInertiaComponents area, viewIx, viewIy, viewIxy
        If LocalInertiaIsNearlySquare(viewIx, viewIy, viewIxy) Then
            displayWidth = Sqr(area): displayHeight = displayWidth
        ElseIf Not TryInterpretAnisotropicBoundary(area, viewIx, viewIy, viewIxy, statusText, displayWidth, displayHeight, displayRotation) Then
            displayWidth = Sqr(area): displayHeight = displayWidth
        End If
    End If
    Dim nextIndex As Long, element As CConcreteElement
    nextIndex = mConcreteCount + 1
    Set element = New CConcreteElement
    element.Initialize "C" & CStr(nextIndex), xCoord, yCoord, area, materialID, sourceName, sourceHandle, shapeType, _
        displayWidth, displayHeight, displayRotation, rotationKnown Or Abs(rotation) > GEOM_TOLERANCE Or explicitRectangle, comment, localIx, localIy, localIxy, exactRegion, explicitRectangle
    EnsureConcreteCapacity nextIndex
    Set mConcreteElements(nextIndex) = element
    mConcreteX(nextIndex) = element.X: mConcreteY(nextIndex) = element.Y: mConcreteArea(nextIndex) = element.Area ' Неизменяемая числовая проекция для горячих циклов; источник данных — элемент.
    mConcreteCount = nextIndex
    mRevision = mRevision + 1
End Sub

' Заранее выделяет места для известного числа элементов, не создавая их
' и не меняя версию геометрии. Генератор передаёт точный размер сетки;
' при импорте массивы могут расширяться по мере чтения. Незанятые места не участвуют в расчёте.
Friend Sub ReserveElements(ByVal concreteCount As Long, ByVal rebarCount As Long)
    If concreteCount < 0 Or rebarCount < 0 Then Err.Raise vbObjectError + 5202, "CSectionModel", "Размер резерва элементов не может быть отрицательным."
    EnsureConcreteCapacity concreteCount, True
    EnsureRebarCapacity rebarCount, True
End Sub

' Увеличивает массив сущностей и числовую проекцию бетона вместе. При неизвестном размере
' рост удвоением устраняет копирование всей модели после каждого Add.
' Clear освобождает резерв; емкость фиксируется лишь после успешного выделения.
Private Sub EnsureConcreteCapacity(ByVal required As Long, Optional ByVal exactSize As Boolean = False)
    If required <= mConcreteCapacity Then Exit Sub
    Dim capacity As Long
    capacity = required
    If Not exactSize Then
        If mConcreteCapacity = 0 Then
            If capacity < 64 Then capacity = 64
        ElseIf mConcreteCapacity <= 1073741823 Then
            If capacity < mConcreteCapacity * 2 Then capacity = mConcreteCapacity * 2
        End If
    End If
    ReDim Preserve mConcreteElements(1 To capacity)
    ReDim Preserve mConcreteX(1 To capacity)
    ReDim Preserve mConcreteY(1 To capacity)
    ReDim Preserve mConcreteArea(1 To capacity)
    ReDim Preserve mBoundaryCacheValid(1 To capacity)
    ReDim Preserve mBoundaryCacheStatus(1 To capacity)
    ReDim Preserve mBoundaryCacheWidth(1 To capacity)
    ReDim Preserve mBoundaryCacheHeight(1 To capacity)
    ReDim Preserve mBoundaryCacheRotation(1 To capacity)
    mConcreteCapacity = capacity
End Sub

' Арматура нумеруется в порядке R1/R2; проверяется фактическое число элементов.
' Резерв выделяется только при добавлении/явном Reserve и уничтожается Clear.
Private Sub EnsureRebarCapacity(ByVal required As Long, Optional ByVal exactSize As Boolean = False)
    If required <= mRebarCapacity Then Exit Sub
    Dim capacity As Long
    capacity = required
    If Not exactSize Then
        If mRebarCapacity = 0 Then
            If capacity < 32 Then capacity = 32
        ElseIf mRebarCapacity <= 1073741823 Then
            If capacity < mRebarCapacity * 2 Then capacity = mRebarCapacity * 2
        End If
    End If
    ReDim Preserve mRebarID(1 To capacity)
    ReDim Preserve mRebarSourceName(1 To capacity)
    ReDim Preserve mRebarSourceHandle(1 To capacity)
    ReDim Preserve mRebarX(1 To capacity)
    ReDim Preserve mRebarY(1 To capacity)
    ReDim Preserve mRebarArea(1 To capacity)
    ReDim Preserve mRebarDiameter(1 To capacity)
    ReDim Preserve mRebarSteelClass(1 To capacity)
    ReDim Preserve mRebarMaterialID(1 To capacity)
    ReDim Preserve mRebarComment(1 To capacity)
    ReDim Preserve mRebarRecipeTrace(1 To capacity)
    mRebarCapacity = capacity
End Sub

' Возвращает версию геометрии без копирования волокон. Она меняется только
' при изменении модели, а не при чтении границ или расчетных характеристик.
Public Property Get Revision() As Long
    If Not mContours Is Nothing Then
        If mContours.Revision <> mObservedContourRevision Then
            mObservedContourRevision = mContours.Revision
            mRevision = mRevision + 1
        End If
    End If
    Revision = mRevision
End Property

' Находит центр бетонной сетки без арматуры и сохраняет его для текущей версии геометрии.
' Суммирует координаты относительно первого элемента, чтобы уменьшить ошибку
' при большом смещении чертежа. Исходные координаты элементов не меняются.
Public Sub ConcreteMeshCentroid(ByRef centerX As Double, ByRef centerY As Double, ByRef totalArea As Double)
    CheckConcreteIndex 1
    Dim currentRevision As Long
    currentRevision = Revision
    If Not mConcreteCenterValid Or mConcreteCenterRevision <> currentRevision Then
        Dim x0 As Double, y0 As Double, sumArea As Double, sumX As Double, sumY As Double, i As Long
        x0 = mConcreteX(1): y0 = mConcreteY(1)
        For i = 1 To mConcreteCount
            sumArea = sumArea + mConcreteArea(i)
            sumX = sumX + mConcreteArea(i) * (mConcreteX(i) - x0)
            sumY = sumY + mConcreteArea(i) * (mConcreteY(i) - y0)
        Next i
        mConcreteCenterX = x0 + sumX / sumArea
        mConcreteCenterY = y0 + sumY / sumArea
        mConcreteCenterArea = sumArea
        mConcreteCenterRevision = currentRevision: mConcreteCenterValid = True
    End If
    centerX = mConcreteCenterX: centerY = mConcreteCenterY: totalArea = mConcreteCenterArea
End Sub

' Находит наибольшее расстояние центра бетонного элемента или стержня от центра бетонной сетки.
' Результат сохраняется для текущей версии геометрии и используется
' во всех пробных полях деформаций без повторного обхода элементов.
Friend Function MaximumElementRadiusFromConcreteCenter() As Double
    Dim currentRevision As Long
    currentRevision = Revision
    If Not mElementRadiusValid Or mElementRadiusRevision <> currentRevision Then
        Dim cx As Double, cy As Double, area As Double, x As Double, y As Double, radius As Double, i As Long
        ConcreteMeshCentroid cx, cy, area
        mElementRadius = 0#
        For i = 1 To mConcreteCount + mRebarCount
            If i <= mConcreteCount Then
                x = mConcreteX(i): y = mConcreteY(i)
            Else
                x = mRebarX(i - mConcreteCount): y = mRebarY(i - mConcreteCount)
            End If
            x = x - cx: y = y - cy
            radius = Sqr(x * x + y * y)
            If radius > mElementRadius Then mElementRadius = radius
        Next i
        mElementRadiusRevision = currentRevision: mElementRadiusValid = True
    End If
    MaximumElementRadiusFromConcreteCenter = mElementRadius
End Function

' Назначает среднюю ориентацию эквивалентным ячейкам, для которых источник
' не задал направления. Их границы затем подготавливаются заново,
' а реальные площади и моменты инерции элементов сохраняются.
Public Sub ApplyAverageRotationToEquivalentAreaFallbacks()
    Dim fallbackRotation As Double
    fallbackRotation = AverageKnownConcreteElementRotation()

    Dim i As Long
    For i = 1 To mConcreteCount
        If IsConcreteEquivalentAreaFallback(i) And Not mConcreteElements(i).RotationKnown Then
            If Abs(mConcreteElements(i).Rotation - fallbackRotation) > GEOM_TOLERANCE Then
                Set mConcreteElements(i) = mConcreteElements(i).WithRotation(fallbackRotation)
                InvalidateBoundaryCache i
                mRevision = mRevision + 1
            End If
        End If
    Next i
End Sub

' Находит средний угол бетонных элементов, для которых оболочка уже имеет
' определенную ориентацию. Учитываются не только явно переданные Rectangle,
' но и Rectangle/Equivalent rectangle, интерпретированные из реальных A/I.
' Изотропный Equivalent square тоже является источником, если importer получил
' угол по прямой грани. В том числе учитывается известный нулевой угол;
' оболочки кругов с назначенным средним направлением источниками не становятся.
Public Function AverageKnownConcreteElementRotation() As Double
    Dim sumCos As Double
    Dim sumSin As Double
    Dim count As Long
    Dim i As Long
    For i = 1 To mConcreteCount
        Dim statusText As String
        Dim widthValue As Double
        Dim heightValue As Double
        Dim rotationValue As Double
        InterpretConcreteBoundary i, statusText, widthValue, heightValue, rotationValue

        If ElementContributesToAverageRotation(i, statusText, widthValue, heightValue) Then
            sumCos = sumCos + Cos(2# * rotationValue) ' Удвоенный угол учитывает беззнаковую ось: направления, отличающиеся на pi, равнозначны.
            sumSin = sumSin + Sin(2# * rotationValue)
            count = count + 1
        End If
    Next i

    If count > 0 And (Abs(sumCos) > GEOM_TOLERANCE Or Abs(sumSin) > GEOM_TOLERANCE) Then
        AverageKnownConcreteElementRotation = 0.5 * Atn2Local(sumSin, sumCos)
    End If
End Function

' Решает, можно ли использовать элемент как источник средней ориентации сетки.
' У Equivalent square происхождение направления хранится независимо от угла:
' известная горизонтальная грань участвует, а среднее круга не участвует.
Private Function ElementContributesToAverageRotation(ByVal index As Long, ByVal statusText As String, _
        ByVal widthValue As Double, _
        ByVal heightValue As Double) As Boolean
    If widthValue <= GEOM_TOLERANCE Or heightValue <= GEOM_TOLERANCE Then Exit Function

    If StrComp(statusText, "Equivalent square", vbTextCompare) = 0 Then
        ElementContributesToAverageRotation = mConcreteElements(index).RotationKnown
    Else
        ElementContributesToAverageRotation = True
    End If
End Function

' Проверяет, используется ли для элемента эквивалентный квадрат только по площади.
' Вид определяется общим расчётом его границ, а не исходной подсказкой импорта.
Public Function IsConcreteEquivalentAreaFallback(ByVal index As Long) As Boolean
    CheckConcreteIndex index
    IsConcreteEquivalentAreaFallback = _
        (StrComp(ConcreteBoundaryShapeType(index), "Equivalent square", vbTextCompare) = 0)
End Function

' Принимает фактическую границу без подмены уже подготовленных расчётных A/c/I.
' Несогласованные данные отклоняются до публикации; новый точный элемент никогда
' не восстанавливается эквивалентным прямоугольником при потере его границы.
Public Sub SetConcreteGeometry(ByVal index As Long, ByVal region As CConcreteRegion)
    CheckConcreteIndex index
    Dim query As CSectionGeometryQuery
    Set query = New CSectionGeometryQuery
    ValidateConcreteBoundary query, region, mConcreteX(index), mConcreteY(index), mConcreteArea(index), _
        mConcreteElements(index).LocalIx, mConcreteElements(index).LocalIy, mConcreteElements(index).LocalIxy, mConcreteElements(index).ElementID
    Dim widthValue As Double, heightValue As Double, rotationValue As Double
    ConcreteRegionDisplay query, region, widthValue, heightValue, rotationValue
    Set mConcreteElements(index) = mConcreteElements(index).WithGeometry(region, widthValue, heightValue, rotationValue)
    InvalidateBoundaryCache index
    mRevision = mRevision + 1
End Sub

' Добавляет расчётное имя модели к общей диагностике границы; сама проверка
' принадлежит геометрическому ядру и одинакова для импорта, модели и снимка.
Private Sub ValidateConcreteBoundary(ByVal query As CSectionGeometryQuery, ByVal region As CConcreteRegion, _
        ByVal x As Double, ByVal y As Double, ByVal area As Double, ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double, ByVal elementID As String)
    On Error GoTo InvalidBoundary
    query.ValidateBoundaryProperties region, x, y, area, localIx, localIy, localIxy
    Exit Sub
InvalidBoundary:
    Err.Raise vbObjectError + 5223, "CSectionModel", "Элемент " & elementID & ": " & Err.Description
End Sub

' Подготавливает отображаемые размеры по фактической области. Для произвольной
' формы это габариты в WCS; точная область остаётся источником геометрии.
Private Sub ConcreteRegionDisplay(ByVal query As CSectionGeometryQuery, ByVal region As CConcreteRegion, _
        ByRef widthValue As Double, ByRef heightValue As Double, ByRef rotationValue As Double)
    Dim centerX As Double, centerY As Double
    If Not query.TryRectangle(region, centerX, centerY, widthValue, heightValue, rotationValue) Then
        Dim minX As Double, maxX As Double, minY As Double, maxY As Double
        query.ProjectionBounds region, 1#, 0#, minX, maxX
        query.ProjectionBounds region, 0#, 1#, minY, maxY
        widthValue = maxX - minX: heightValue = maxY - minY: rotationValue = 0# ' Для произвольной формы это только габарит рисунка; геометрические потребители читают саму область.
    End If
End Sub

' Возвращает единое геометрическое представление элемента. Эквивалентная
' оболочка описывает только компактные записи без маркера точной границы.
Public Function ConcreteGeometry(ByVal index As Long) As CConcreteRegion
    CheckConcreteIndex index
    Set ConcreteGeometry = mConcreteElements(index).Geometry
End Function

' Отличает сохранённую фактическую область от компактного прямоугольного описания.
' Результат не определяет способ построения: импорт и оба генератора используют один контракт.
Public Function ConcreteHasBoundary(ByVal index As Long) As Boolean
    CheckConcreteIndex index
    ConcreteHasBoundary = mConcreteElements(index).HasExactGeometry
End Function

' Возвращает настоящий прямоугольник точного элемента либо компактную оболочку.
' Произвольная фактическая форма возвращает False. Только компактные записи
' интерпретируют A/I как Width/Height/Rotation, без обратного пересчёта инерций.
Public Function ConcreteBoundaryRectangle(ByVal index As Long, ByRef widthValue As Double, _
        ByRef heightValue As Double, ByRef rotationValue As Double) As Boolean
    CheckConcreteIndex index
    If mConcreteElements(index).HasExactGeometry Then
        Dim query As CSectionGeometryQuery, centerX As Double, centerY As Double
        Set query = New CSectionGeometryQuery
        ConcreteBoundaryRectangle = query.TryRectangle(mConcreteElements(index).Geometry, centerX, centerY, widthValue, heightValue, rotationValue)
        Exit Function
    End If
    Dim statusText As String
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteBoundaryRectangle = (widthValue > GEOM_TOLERANCE And heightValue > GEOM_TOLERANCE)
End Function

' Возвращает статус интерпретации геометрической оболочки бетонного элемента.
' Возможные значения: Contour, Rectangle, Equivalent rectangle, Equivalent square.
Public Function ConcreteBoundaryShapeType(ByVal index As Long) As String
    Dim statusText As String
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteBoundaryShapeType = statusText
End Function

' Возвращает реальные крайние проекции области; стороны произвольной ячейки
' могут находиться на разных расстояниях от её центра тяжести.
Public Sub ConcreteProjectionBounds(ByVal index As Long, ByVal axisX As Double, ByVal axisY As Double, _
        ByRef minimum As Double, ByRef maximum As Double)
    CheckConcreteIndex index
    If mConcreteElements(index).HasExactGeometry Then
        Dim query As CSectionGeometryQuery, axisNorm As Double, axisScale As Double
        Set query = New CSectionGeometryQuery
        axisScale = GeomMax(Abs(axisX), Abs(axisY))
        If axisScale <= 0# Then Err.Raise vbObjectError + 5223, "CSectionModel", "Направление проекции бетонной ячейки имеет нулевую длину."
        axisNorm = axisScale * Sqr((axisX / axisScale) ^ 2 + (axisY / axisScale) ^ 2) ' Линейные коэффициенты ядра могут быть много меньше геометрического допуска, оставаясь ненулевыми.
        query.ProjectionBounds mConcreteElements(index).Geometry, axisX / axisNorm, axisY / axisNorm, minimum, maximum
        minimum = minimum * axisNorm: maximum = maximum * axisNorm ' Общий запрос нормирует направление; линейная проекция ядра сохраняет его исходный масштаб.
    Else
        Dim centerProjection As Double, halfProjection As Double
        centerProjection = axisX * mConcreteX(index) + axisY * mConcreteY(index)
        halfProjection = ConcreteBoundaryHalfProjection(index, axisX, axisY)
        minimum = centerProjection - halfProjection: maximum = centerProjection + halfProjection
    End If
End Sub

' Возвращает максимальное расстояние от центра до проекции границы.
' Для точных несимметричных габаритов следует читать обе ConcreteProjectionBounds.
Public Function ConcreteBoundaryHalfProjection(ByVal index As Long, ByVal axisX As Double, _
        ByVal axisY As Double) As Double
    CheckConcreteIndex index
    If mConcreteElements(index).HasExactGeometry Then
        Dim minimum As Double, maximum As Double, centerProjection As Double
        ConcreteProjectionBounds index, axisX, axisY, minimum, maximum
        centerProjection = axisX * mConcreteX(index) + axisY * mConcreteY(index)
        ConcreteBoundaryHalfProjection = GeomMax(Abs(minimum - centerProjection), Abs(maximum - centerProjection))
        Exit Function
    End If
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    If ConcreteBoundaryRectangle(index, widthValue, heightValue, rotationValue) Then
        ConcreteBoundaryHalfProjection = RectangleHalfProjection(widthValue, heightValue, rotationValue, axisX, axisY)
    End If
End Function

' Возвращает собственные моменты инерции элемента. Если источник не передал
' пригодных значений, рассчитывает их для квадрата той же площади.
Public Sub ConcreteLocalInertiaComponents(ByVal index As Long, ByRef localIx As Double, _
        ByRef localIy As Double, ByRef localIxy As Double)
    CheckConcreteIndex index

    localIx = mConcreteElements(index).LocalIx
    localIy = mConcreteElements(index).LocalIy
    localIxy = mConcreteElements(index).LocalIxy
    If mConcreteElements(index).HasExactGeometry Then
        If Not ExactPositiveInertia(localIx, localIy, localIxy) Then Err.Raise vbObjectError + 5223, "CSectionModel", "Повреждены собственные инерции фактической бетонной ячейки."
        Exit Sub
    End If
    If IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then Exit Sub

    EquivalentSquareLocalInertiaComponents mConcreteArea(index), localIx, localIy, localIxy
End Sub

' Проверяет, различаются ли собственные моменты инерции достаточно, чтобы определить
' направление главных осей элемента. При почти равных инерциях импорт AutoCAD
' может уточнить угол по фактической прямой грани Region.
Public Function LocalInertiaIsNearlySquare(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double, Optional ByVal relativeTolerance As Double = LOCAL_INERTIA_SQUARE_REL_TOLERANCE) As Boolean
    If Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then Exit Function

    Dim i1 As Double
    Dim i2 As Double
    Dim angle1 As Double
    PrincipalLocalInertia localIx, localIy, localIxy, i1, i2, angle1
    If i1 <= GEOM_TOLERANCE Then Exit Function
    LocalInertiaIsNearlySquare = (Abs(i1 - i2) / i1 <= relativeTolerance)
End Function

' Интерпретирует оболочку бетонного элемента по единому каскаду:
' 1) явно известный прямоугольник источника;
' 2) прямоугольник, восстановленный по реальным A/I;
' 3) эквивалентный квадрат по площади.
Private Sub InterpretConcreteBoundary(ByVal index As Long, ByRef statusText As String, _
        ByRef widthValue As Double, ByRef heightValue As Double, ByRef rotationValue As Double)
    CheckConcreteIndex index

    If mBoundaryCacheValid(index) Then
        statusText = mBoundaryCacheStatus(index)
        widthValue = mBoundaryCacheWidth(index)
        heightValue = mBoundaryCacheHeight(index)
        rotationValue = mBoundaryCacheRotation(index)
        Exit Sub
    End If

    CalculateConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    StoreBoundaryCache index, statusText, widthValue, heightValue, rotationValue
End Sub

' Выполняет фактическую интерпретацию геометрической оболочки элемента.
' Этот расчет может включать собственные значения тензора инерции, поэтому
' CSectionModel кэширует результат и не повторяет его при каждом Width/Height.
Private Sub CalculateConcreteBoundary(ByVal index As Long, ByRef statusText As String, _
        ByRef widthValue As Double, ByRef heightValue As Double, ByRef rotationValue As Double)
    If StrComp(mConcreteElements(index).ShapeType, "Contour", vbTextCompare) = 0 Then
        If Not mConcreteElements(index).HasExactGeometry Then Err.Raise vbObjectError + 5223, "CSectionModel", "Потеряна обязательная фактическая граница элемента " & mConcreteElements(index).ElementID & "."
        statusText = "Contour"
        widthValue = mConcreteElements(index).Width: heightValue = mConcreteElements(index).Height: rotationValue = mConcreteElements(index).Rotation
        Exit Sub
    End If
    If HasExplicitRectangleGeometry(index) Then
        statusText = "Rectangle"
        widthValue = mConcreteElements(index).Width
        heightValue = mConcreteElements(index).Height
        rotationValue = NormalizeAnglePi(mConcreteElements(index).Rotation)
        Exit Sub
    End If

    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy

    If LocalInertiaIsNearlySquare(localIx, localIy, localIxy) Then
        statusText = "Equivalent square"
        widthValue = Sqr(mConcreteArea(index))
        heightValue = widthValue
        rotationValue = NormalizeAnglePi(mConcreteElements(index).Rotation)
        Exit Sub
    End If

    If TryInterpretAnisotropicBoundary(mConcreteArea(index), localIx, localIy, localIxy, _
            statusText, widthValue, heightValue, rotationValue) Then Exit Sub

    statusText = "Equivalent square"
    widthValue = Sqr(mConcreteArea(index))
    heightValue = widthValue
    rotationValue = NormalizeAnglePi(mConcreteElements(index).Rotation)
End Sub

' Сохраняет рассчитанные размеры и угол границы элемента для геометрии, схемы
' и AutoCAD. Реальные площадь и инерции для расчёта усилий не меняются.
Private Sub StoreBoundaryCache(ByVal index As Long, ByVal statusText As String, _
        ByVal widthValue As Double, ByVal heightValue As Double, _
        ByVal rotationValue As Double)
    mBoundaryCacheStatus(index) = statusText
    mBoundaryCacheWidth(index) = widthValue
    mBoundaryCacheHeight(index) = heightValue
    mBoundaryCacheRotation(index) = rotationValue
    mBoundaryCacheValid(index) = True
End Sub

' Сбрасывает один элемент кэша после изменения данных, влияющих на оболочку.
Private Sub InvalidateBoundaryCache(ByVal index As Long)
    If index < 1 Or index > mConcreteCount Then Exit Sub
    mBoundaryCacheValid(index) = False
    mBoundaryCacheStatus(index) = vbNullString
    mBoundaryCacheWidth(index) = 0#
    mBoundaryCacheHeight(index) = 0#
    mBoundaryCacheRotation(index) = 0#
End Sub

' Полностью очищает кэш интерпретации оболочек.
Private Sub ClearBoundaryCache()
    Erase mBoundaryCacheValid
    Erase mBoundaryCacheStatus
    Erase mBoundaryCacheWidth
    Erase mBoundaryCacheHeight
    Erase mBoundaryCacheRotation
End Sub

' Восстанавливает прямоугольные границы элемента по его площади и главным инерциям.
' Если они соответствуют настоящему прямоугольнику, ставит Rectangle; иначе
' использует Equivalent rectangle с той же площадью и отношением главных инерций.
Private Function TryInterpretAnisotropicBoundary(ByVal areaValue As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByRef statusText As String, ByRef widthValue As Double, _
        ByRef heightValue As Double, ByRef rotationValue As Double) As Boolean
    If areaValue <= GEOM_TOLERANCE Then Exit Function

    Dim i1 As Double
    Dim i2 As Double
    Dim angle1 As Double
    PrincipalLocalInertia localIx, localIy, localIxy, i1, i2, angle1
    If i1 <= GEOM_TOLERANCE Or i2 <= GEOM_TOLERANCE Then Exit Function

    Dim rectangleWidth As Double
    Dim rectangleHeight As Double
    rectangleWidth = Sqr(12# * i2 / areaValue) ' В главных осях Iy = A*B^2/12; эта сторона соответствует меньшей инерции i2.
    rectangleHeight = Sqr(12# * i1 / areaValue)

    If rectangleWidth > GEOM_TOLERANCE And rectangleHeight > GEOM_TOLERANCE Then
        If Abs(rectangleWidth * rectangleHeight - areaValue) / areaValue <= _
                LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE Then
            statusText = "Rectangle"
            widthValue = rectangleWidth
            heightValue = rectangleHeight
            rotationValue = NormalizeAnglePi(angle1)
            TryInterpretAnisotropicBoundary = True
            Exit Function
        End If
    End If

    statusText = "Equivalent rectangle"
    heightValue = Sqr(areaValue) * (i1 / i2) ^ 0.25 ' Эквивалентный прямоугольник сохраняет площадь и отношение главных инерций, а не их абсолютные значения.
    widthValue = Sqr(areaValue) * (i2 / i1) ^ 0.25
    rotationValue = NormalizeAnglePi(angle1)
    TryInterpretAnisotropicBoundary = (widthValue > GEOM_TOLERANCE And heightValue > GEOM_TOLERANCE)
End Function

' Возвращает главные локальные инерции и направление оси I1.
' Угол считается от глобальной +X против часовой стрелки; ось физически
' беззнаковая, поэтому +180 градусов означает то же направление.
Public Sub PrincipalLocalInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double, ByRef i1 As Double, ByRef i2 As Double, _
        ByRef angle1 As Double)
    Dim averageI As Double
    Dim radiusI As Double
    averageI = (localIx + localIy) / 2#
    radiusI = Sqr(((localIx - localIy) / 2#) ^ 2 + localIxy ^ 2)
    i1 = averageI + radiusI
    i2 = averageI - radiusI

    If LocalInertiaIsNumericallyIsotropic(localIx, localIy, localIxy) Then
        angle1 = 0#
    Else
        angle1 = NormalizeAnglePi(0.5 * Atn2Local(-2# * localIxy, localIx - localIy))
    End If
End Sub

' Явно известная прямоугольная оболочка допустима только как геометрический
' факт источника. Она не отменяет реальные Ix/Iy/Ixy, если они уже переданы.
Private Function HasExplicitRectangleGeometry(ByVal index As Long) As Boolean
    HasExplicitRectangleGeometry = mConcreteElements(index).HasKnownRectangle
End Function

' Считает Ix/Iy/Ixy прямоугольной ячейки вокруг собственного центра.
' Используется только когда прямоугольная форма известна до расчета A/I,
' например у встроенного генератора сетки.
Private Sub RectangleLocalInertiaComponents(ByVal areaValue As Double, _
        ByVal widthValue As Double, ByVal heightValue As Double, ByVal angleValue As Double, _
        ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    Dim ixLocal As Double
    Dim iyLocal As Double
    Dim c As Double
    Dim s As Double
    ixLocal = areaValue * heightValue * heightValue / 12#
    iyLocal = areaValue * widthValue * widthValue / 12#
    c = Cos(angleValue)
    s = Sin(angleValue)

    localIx = ixLocal * c * c + iyLocal * s * s
    localIy = ixLocal * s * s + iyLocal * c * c
    localIxy = (iyLocal - ixLocal) * s * c
End Sub

' Вычисляет собственные инерции эквивалентного квадрата той же площади,
' когда более точные характеристики элемента отсутствуют.
Private Sub EquivalentSquareLocalInertiaComponents(ByVal areaValue As Double, _
        ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    localIx = areaValue * areaValue / 12#
    localIy = localIx
    localIxy = 0#
End Sub

' Возвращает половину проекции прямоугольной оболочки на произвольную ось.
Private Function RectangleHalfProjection(ByVal widthValue As Double, ByVal heightValue As Double, _
        ByVal angleValue As Double, ByVal axisX As Double, ByVal axisY As Double) As Double
    Dim ux As Double
    Dim uy As Double
    Dim vx As Double
    Dim vy As Double
    ux = Cos(angleValue)
    uy = Sin(angleValue)
    vx = -uy
    vy = ux
    RectangleHalfProjection = 0.5 * (widthValue * Abs(axisX * ux + axisY * uy) + _
        heightValue * Abs(axisX * vx + axisY * vy))
End Function

' Проверяет положительную определенность локальной матрицы инерций.
Private Function IsPositiveDefiniteInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double) As Boolean
    If localIx <= GEOM_TOLERANCE Or localIy <= GEOM_TOLERANCE Then Exit Function

    Dim scaleValue As Double
    scaleValue = GeomMax(Abs(localIx * localIy), 1#)
    IsPositiveDefiniteInertia = ((localIx * localIy - localIxy * localIxy) / scaleValue > GEOM_TOLERANCE)
End Function

' Проверяет, можно ли считать локальный тензор инерции изотропным для выбора оболочки элемента.
Private Function LocalInertiaIsNumericallyIsotropic(ByVal localIx As Double, _
        ByVal localIy As Double, ByVal localIxy As Double) As Boolean
    Dim scaleValue As Double
    scaleValue = GeomMax(GeomMax(Abs(localIx), Abs(localIy)), 1#)
    LocalInertiaIsNumericallyIsotropic = _
        Abs(localIx - localIy) / scaleValue <= LOCAL_INERTIA_SQUARE_REL_TOLERANCE And _
        Abs(localIxy) / scaleValue <= LOCAL_INERTIA_SQUARE_REL_TOLERANCE
End Function

' Нормализует угол беззнаковой оси в диапазон (-pi/2; pi/2], радианы.
Private Function NormalizeAnglePi(ByVal angleValue As Double) As Double
    If Abs(angleValue) > 2# * GEOM_PI Then
        On Error GoTo InvalidAngle
        angleValue = Atn2Local(Sin(angleValue), Cos(angleValue)) ' Большой угол сводим за постоянное время: при 1e20 вычитание pi уже не изменяет Double.
        On Error GoTo 0
    End If
    Do While angleValue > GEOM_PI / 2#
        angleValue = angleValue - GEOM_PI
    Loop
    Do While angleValue <= -GEOM_PI / 2#
        angleValue = angleValue + GEOM_PI
    Loop
    NormalizeAnglePi = angleValue
    Exit Function
InvalidAngle:
    Err.Raise vbObjectError + 5222, "CSectionModel", "Не удалось определить направление бетонного элемента по заданному углу. Укажите конечный угол поворота в радианах."
End Function

' Добавляет стержень арматуры. Если площадь не задана, вычисляет её по диаметру.
' Импорт AutoCAD может передать фактическую площадь Region,
' если она известна точнее диаметра.
Public Sub AddRebarElement(ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal comment As String = vbNullString)
    If diameter <= 0# Then Err.Raise vbObjectError + 5210, "CSectionModel", "Диаметр арматуры должен быть положительным."

    Dim actualArea As Double ' Площадь, которая реально попадет в расчетную модель.
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 5211, "CSectionModel", "Площадь арматурного элемента должна быть положительной."

    EnsureRebarCapacity mRebarCount + 1 ' Нумерация R1/R2/... создается здесь, чтобы у всех источников геометрии был одинаковый формат расчетных имен.
    mRebarCount = mRebarCount + 1

    mRebarID(mRebarCount) = "R" & CStr(mRebarCount) ' Единое расчётное имя создаётся здесь; имя источника хранится отдельным полем.
    mRebarSourceName(mRebarCount) = sourceName
    mRebarSourceHandle(mRebarCount) = sourceHandle
    mRebarX(mRebarCount) = xCoord
    mRebarY(mRebarCount) = yCoord
    mRebarArea(mRebarCount) = actualArea
    mRebarDiameter(mRebarCount) = diameter
    mRebarSteelClass(mRebarCount) = steelClass
    mRebarMaterialID(mRebarCount) = materialID
    mRebarComment(mRebarCount) = comment
    mRevision = mRevision + 1
End Sub

' Фактическое число бетонных элементов в модели.
Public Property Get ConcreteCount() As Long
    ConcreteCount = mConcreteCount
End Property

' Фактическое число стержней в модели.
Public Property Get RebarCount() As Long
    RebarCount = mRebarCount
End Property

' Расчетные имена бетонных элементов: C1, C2, ...
Public Property Get ConcreteID(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteID = mConcreteElements(index).ElementID
End Property

' Необязательное имя исходной бетонной области для диагностики.
Public Property Get ConcreteSourceName(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceName = mConcreteElements(index).SourceName
End Property

' Необязательный AutoCAD handle для поиска исходного Region.
Public Property Get ConcreteSourceHandle(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceHandle = mConcreteElements(index).SourceHandle
End Property

' X центра тяжести бетонного элемента во внутренней системе координат, мм.
Public Property Get ConcreteX(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteX = mConcreteX(index)
End Property

' Y центра тяжести бетонного элемента во внутренней системе координат, мм.
Public Property Get ConcreteY(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteY = mConcreteY(index)
End Property

' Площадь бетонного элемента, мм2.
Public Property Get ConcreteArea(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteArea = mConcreteArea(index)
End Property

' Возвращает координаты и площадь бетонного элемента одним обращением.
' Внутренний цикл решения НДС проверяет номер элемента один раз
' вместо отдельных проверок для каждой координаты и площади.
Public Sub ConcreteElementForSolver(ByVal index As Long, ByRef xCoord As Double, _
        ByRef yCoord As Double, ByRef area As Double)
    CheckConcreteIndex index
    xCoord = mConcreteX(index)
    yCoord = mConcreteY(index)
    area = mConcreteArea(index)
End Sub

' Возвращает номер материала элемента; текущий расчёт использует одну подготовленную диаграмму
' бетона.
Public Property Get ConcreteMaterialID(ByVal index As Long) As Long
    CheckConcreteIndex index
    ConcreteMaterialID = mConcreteElements(index).MaterialID
End Property

' Геометрическое представление выбранной бетонной ячейки.
Public Property Get ConcreteShapeType(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteShapeType = ConcreteBoundaryShapeType(index)
End Property

' Читает ширину представления границы элемента. Для импортированного волокна
' оно может быть эквивалентным; реальная площадь и инерции не заменяются.
Public Property Get ConcreteWidth(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    Dim statusText As String
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteWidth = widthValue
End Property

' Читает высоту той же границы/эквивалентного отображения, не меняя расчетные A/I.
Public Property Get ConcreteHeight(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    Dim statusText As String
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteHeight = heightValue
End Property

' Возвращает поворот отображаемого прямоугольника относительно расчетных осей.
Public Property Get ConcreteRotation(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    Dim statusText As String
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteRotation = rotationValue
End Property

' Читает собственную Ix элемента из единого блока реальных геометрических характеристик.
Public Property Get ConcreteLocalIx(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIx = localIx
End Property

' Читает собственную Iy элемента, а не инерцию его эквивалентного изображения.
Public Property Get ConcreteLocalIy(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIy = localIy
End Property

' Читает собственную центробежную инерцию элемента в расчетных координатах.
Public Property Get ConcreteLocalIxy(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIxy = localIxy
End Property

' Короткая диагностическая заметка для Results.
Public Property Get ConcreteComment(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteComment = mConcreteElements(index).Comment
End Property

' Расчетные имена стержней: R1, R2, ...
Public Property Get RebarID(ByVal index As Long) As String
    CheckRebarIndex index
    RebarID = mRebarID(index)
End Property

' Необязательный AutoCAD handle арматурного Region.
Public Property Get RebarSourceHandle(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceHandle = mRebarSourceHandle(index)
End Property

' X центра стержня, мм.
Public Property Get RebarX(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarX = mRebarX(index)
End Property

' Y центра стержня, мм.
Public Property Get RebarY(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarY = mRebarY(index)
End Property

' Площадь стержня, мм2.
Public Property Get RebarArea(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarArea = mRebarArea(index)
End Property

' Отдает координаты и площадь стержня одним обращением для горячего цикла
' равновесия. Расчетная механика не меняется: возвращаются те же данные, что
' и через RebarX/RebarY/RebarArea.
Public Sub RebarElementForSolver(ByVal index As Long, ByRef xCoord As Double, _
        ByRef yCoord As Double, ByRef area As Double)
    CheckRebarIndex index
    xCoord = mRebarX(index)
    yCoord = mRebarY(index)
    area = mRebarArea(index)
End Sub

' Диаметр стержня, мм.
Public Property Get RebarDiameter(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarDiameter = mRebarDiameter(index)
End Property

' Возвращает номер материала стержня; текущий расчёт использует одну подготовленную диаграмму
' арматуры.
Public Property Get RebarMaterialID(ByVal index As Long) As Long
    CheckRebarIndex index
    RebarMaterialID = mRebarMaterialID(index)
End Property

' Диагностическая заметка о грани, ряде или источнике стержня.
Public Property Get RebarComment(ByVal index As Long) As String
    CheckRebarIndex index
    RebarComment = mRebarComment(index)
End Property

' Возвращает общий список размеров, подписей и расчётных областей модели.
' Если списка ещё нет, создаёт пустой; добавленные аннотации относятся к этой же модели.
Public Property Get Annotations() As CSectionAnnotations
    If mAnnotations Is Nothing Then Set mAnnotations = New CSectionAnnotations
    Set Annotations = mAnnotations
End Property

' Возвращает исходные границы материала отдельно от аннотаций и расчетных областей.
Public Property Get Contours() As CSectionContours
    If mContours Is Nothing Then Set mContours = New CSectionContours
    Set Contours = mContours
End Property

' Возвращает число сохраненных аннотаций без создания коллекции.
' У модели без аннотаций результат равен нулю, расчетная геометрия не меняется.
Public Property Get AnnotationCount() As Long
    If mAnnotations Is Nothing Then
        AnnotationCount = 0
    Else
        AnnotationCount = mAnnotations.Count
    End If
End Property

' Проверяет индекс бетонного элемента перед доступом к массивам модели.
Private Sub CheckConcreteIndex(ByVal index As Long)
    If index < 1 Or index > mConcreteCount Then Err.Raise vbObjectError + 5220, "CSectionModel", "Индекс бетонного элемента вне допустимого диапазона." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mConcreteCount) & "."
End Sub

' Проверяет индекс арматурного элемента перед доступом к массивам модели.
Private Sub CheckRebarIndex(ByVal index As Long)
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 5221, "CSectionModel", "Индекс арматурного элемента вне допустимого диапазона." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mRebarCount) & "."
End Sub

' Возвращает atan2(y, x) без зависимости от внешних модулей.
' Используется для главных локальных осей элемента и усреднения направлений импортированной сетки.
Private Function Atn2Local(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        If y > 0# Then
            Atn2Local = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2Local = -GEOM_PI / 2#
        Else
            Atn2Local = 0#
        End If
    Else
        Atn2Local = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then
                Atn2Local = Atn2Local + GEOM_PI
            Else
                Atn2Local = Atn2Local - GEOM_PI
            End If
        End If
    End If
End Function

' Возвращает исходное имя стержня для сохранения происхождения в Results.
' Расчётное имя R1/R2/... хранится отдельно; исходное имя не влияет на равновесие.
Public Property Get RebarSourceName(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceName = mRebarSourceName(index)
End Property

' Возвращает название материала, переданное при добавлении стержня и сохраняемое в Results.
Public Property Get RebarSteelClass(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSteelClass = mRebarSteelClass(index)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Позволяет сравнить полный поиск с прямым обходом без изменения сетки и материалов.
Public Sub TestSetForceIntegrationEnabled(ByVal enabled As Boolean)
    mTestDisableForceIntegration = Not enabled
    mForceProbeCount = 0: Set mForceIntegrator = Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Показывает число настоящих проб для проверки холодного одиночного и пакетного пути.
Public Property Get TestForceProbeCount() As Long
    TestForceProbeCount = mForceProbeCount
End Property
