VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionTypeRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Реестр поддерживаемых типов сечений
' ==========================================================================
' Registry - единственная точка, где строка Geometry.Type из Config
' сопоставляется с конкретным набором классов формы: CGeometry*,
' *RebarLayoutBuilder и *AnnotationBuilder. Он не должен знать внутреннюю
' структуру таблиц граней и рядов конкретного сечения: такие детали остаются
' внутри специализированных builders. Отдельные provider-классы здесь
' сознательно не используются: обязательны геометрия с точными outer/opening
' и арматурный builder, а annotation builder необязателен. Новый тип
' подключается только здесь; ниже по цепочке используются общие контракты.

Private Const TYPE_CIRCLE As String = "circle" ' Ключ круглого сечения в Geometry.Type.
Private Const TYPE_RECTSET As String = "rectset" ' Ключ сечения из одного или двух прямоугольников в Geometry.Type.
Private Const TYPE_ROUNDED_RECTANGLE As String = "roundedrectangle" ' Ключ скругленного прямоугольника.
Private Const TYPE_HOLLOW_RECTANGLE As String = "hollowrectangle" ' Ключ пустотелого скругленного прямоугольника.

' Возвращает нормализованный тип встроенной геометрии из Config.
' Все публичные методы registry пользуются одним и тем же значением, поэтому
' новая форма регистрируется в одном месте и не расползается по workbook-коду.
Private Function ActiveGeometryType(ByVal settings As CSystemSettingsReader) As String
    If settings Is Nothing Then Err.Raise vbObjectError + 4135, _
        "CSectionTypeRegistry", "Настройки Config не переданы."

    Dim geometryType As String
    geometryType = LCase$(settings.GetRequiredChoice("Geometry.Type", _
        Array("Circle", "RectSet", "RoundedRectangle", "HollowRectangle")))

    Select Case geometryType
        Case TYPE_CIRCLE, TYPE_RECTSET, TYPE_ROUNDED_RECTANGLE, TYPE_HOLLOW_RECTANGLE
            ActiveGeometryType = geometryType
        Case Else
            Err.Raise vbObjectError + 4134, "CSectionTypeRegistry", _
                "Неподдерживаемый тип сечения Geometry.Type: " & settings.GetString("Geometry.Type", "RectSet")
    End Select
End Function

' Создает математическую форму сечения для Geometry.Source = Generated.
' Здесь допустим shape-specific Select Case: это единственная точка регистрации
' встроенных форм. Сетка и проверки арматуры получают общий CGeometryRegion.
Public Function CreateShape(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionShape
    If units Is Nothing Then Err.Raise vbObjectError + 4135, "CSectionTypeRegistry", "Система единиц геометрии не передана."
    Dim circleGeometry As CGeometryCircle ' Конкретная математическая форма круга.
    Dim rectSetGeometry As CGeometryRectSet ' Математическая форма из одного/двух прямоугольников.
    Dim roundedGeometry As CGeometryRoundedRectangle ' Конкретная математическая форма скругленного прямоугольника.
    Dim hollowGeometry As CGeometryHollowRectangle ' Пустотелая форма: внешний RoundedRectangle минус Opening.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleGeometry = New CGeometryCircle
            circleGeometry.InitializeByDiameter units.InputLengthToInternal(settings.GetRequiredDouble("Circle.Diameter")), 0#, 0#
            Set CreateShape = circleGeometry

        Case TYPE_RECTSET
            Set rectSetGeometry = New CGeometryRectSet
            Dim sectionType As String, lowerB As Double, lowerH As Double, upperOffset As Double
            sectionType = settings.GetRequiredChoice("RectSet.SectionType", _
                Array("Rectangle", "LSection", "TwoRectangles", "RectSet", "TwoRectangle", "TwoRectanglesOffset", "OffsetTwoRectangles"))
            If sectionType <> "Rectangle" Then
                lowerB = units.InputLengthToInternal(settings.GetRequiredDouble("RectSet.B2"))
                lowerH = units.InputLengthToInternal(settings.GetRequiredDouble("RectSet.H2"))
            End If
            If sectionType <> "Rectangle" And sectionType <> "LSection" And sectionType <> "RectSet" Then _
                upperOffset = units.InputLengthToInternal(settings.GetRequiredDouble("RectSet.UpperOffsetX"))
            rectSetGeometry.Initialize units.InputLengthToInternal(settings.GetRequiredDouble("RectSet.B1")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("RectSet.H1")), _
                lowerB, lowerH, 0#, 0#, upperOffset, sectionType
            Set CreateShape = rectSetGeometry

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedGeometry = New CGeometryRoundedRectangle
            Dim leftType As String, rightType As String, leftW As Double, rightW As Double, leftR2 As Double, rightR2 As Double
            leftType = settings.GetRequiredChoice("RoundedRectangle.Left.Type", Array("Simple", "Tapered"))
            rightType = settings.GetRequiredChoice("RoundedRectangle.Right.Type", Array("Simple", "Tapered"))
            If leftType = "Tapered" Then
                leftW = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.W"))
                leftR2 = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.R2"))
            End If
            If rightType = "Tapered" Then
                rightW = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.W"))
                rightR2 = units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.R2"))
            End If
            roundedGeometry.InitializeSides _
                units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.B")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.H")), _
                leftType, rightType, leftW, rightW, _
                units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Left.R1")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("RoundedRectangle.Right.R1")), leftR2, rightR2
            Set CreateShape = roundedGeometry

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowGeometry = New CGeometryHollowRectangle
            hollowGeometry.Initialize _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.B")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.H")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.R")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.OpeningB")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.OpeningH")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.OpeningR")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.InnerOffsetX")), _
                units.InputLengthToInternal(settings.GetRequiredDouble("HollowRectangle.InnerOffsetY"))
            Set CreateShape = hollowGeometry
    End Select
End Function

' Общая бетонная область для потребителей, которым не нужны параметры формы.
Public Function CreateGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CGeometryRegion
    Set CreateGeometry = BuildConcreteGeometry(CreateShape(settings, units))
End Function

' Создает раскладку арматуры для активной встроенной формы.
' Registry выбирает нужный builder, а чтение shape-specific ключей Config
' выполняет сам builder. Так таблица граней RectSet и правила дополнительных
' рядов не размазываются по общему registry.
Public Function CreateRebars(ByVal geometry As CGeometryRegion, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Dim circleRebars As CCircleRebarLayoutBuilder ' Генератор раскладки арматуры для кругового сечения.
    Dim rectSetRebars As CRectSetRebarLayoutBuilder ' Генератор раскладки арматуры для RectSet-сечения.
    Dim roundedRebars As CRoundedRectRebarLayoutBuilder ' Генератор арматуры для RoundedRectangle-сечения.
    Dim hollowRebars As CHollowRectRebarLayoutBuilder ' Генератор арматуры для HollowRectangle-сечения.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleRebars = New CCircleRebarLayoutBuilder
            Set CreateRebars = circleRebars.BuildFromSettings(settings, units, concrete:=geometry)

        Case TYPE_RECTSET
            Set rectSetRebars = New CRectSetRebarLayoutBuilder
            Set CreateRebars = rectSetRebars.BuildFromSettings(settings, units, geometry)

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedRebars = New CRoundedRectRebarLayoutBuilder
            Set CreateRebars = roundedRebars.BuildFromSettings(settings, units, geometry)

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowRebars = New CHollowRectRebarLayoutBuilder
            Set CreateRebars = hollowRebars.BuildFromSettings(settings, units, geometry)
    End Select
End Function

' Добавляет semantic-аннотации формы в готовую CSectionModel.
' AnnotationBuilder получает модельные координаты и смысловые группы арматуры;
' Excel/Chart-координаты по-прежнему рассчитывает только layout/plot layer.
' Контуры здесь не создаются и не очищаются. Если зарегистрированная форма
' не имеет annotation builder, модель остается без размеров/подписей граней;
' геометрия и классификация всех отверстий уже подготовлены независимо.
Public Sub BuildAnnotations(ByVal model As CSectionModel, ByVal shape As ISectionShape, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader)
    If model Is Nothing Then Err.Raise vbObjectError + 4136, "CSectionTypeRegistry", "Модель сечения не передана."
    If shape Is Nothing Then Err.Raise vbObjectError + 4137, "CSectionTypeRegistry", "Описание формы сечения не передано."

    Dim circleAnnotations As CCircleAnnotationBuilder ' Формирует semantic-размеры и подписи арматуры для круга.
    Dim rectSetAnnotations As CRectSetAnnotationBuilder ' Формирует semantic-аннотации для граней RectSet-сечения.
    Dim roundedAnnotations As CRoundedRectAnnotationBuilder ' Формирует semantic-аннотации скругленного прямоугольника.
    Dim hollowAnnotations As CHollowRectAnnotationBuilder ' Формирует semantic-аннотации наружного контура и Opening.

    model.Annotations.Clear

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleAnnotations = New CCircleAnnotationBuilder
            Dim circleShape As CGeometryCircle: Set circleShape = shape
            circleAnnotations.Build model, circleShape, rebars

        Case TYPE_RECTSET
            Set rectSetAnnotations = New CRectSetAnnotationBuilder
            Dim rectSet As CGeometryRectSet: Set rectSet = shape
            rectSetAnnotations.Build model, rectSet, rebars

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedAnnotations = New CRoundedRectAnnotationBuilder
            Dim rounded As CGeometryRoundedRectangle: Set rounded = shape
            roundedAnnotations.Build model, rounded, rebars

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowAnnotations = New CHollowRectAnnotationBuilder
            Dim hollow As CGeometryHollowRectangle: Set hollow = shape
            hollowAnnotations.Build model, hollow, rebars
    End Select
End Sub

' Собирает Generated-модель целиком: форму, бетонную сетку, арматуру и
' точные контуры с ролями и необязательные semantic-аннотации. Новый тип
' требует геометрию и арматурный builder, необязательно annotation builder;
' writer/reader/layout/plotter и расчетные потребители не меняются.
Public Function BuildGeneratedModel(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CSectionModel
    Dim shape As ISectionShape, geometry As CGeometryRegion
    Set shape = CreateShape(settings, units)
    Set geometry = BuildConcreteGeometry(shape)

    Dim mesh As CFiberMeshBuilder
    Set mesh = BuildConfiguredConcreteMesh(geometry, settings, units)

    Dim rebars As CRebarLayout
    Set rebars = CreateRebars(geometry, settings, units)
    If rebars Is Nothing Then
        Err.Raise vbObjectError + 4138, "CSectionTypeRegistry", _
            "Для выбранного параметрического сечения не создана раскладка арматуры."
    End If
    If rebars.Count = 0 Then
        Err.Raise vbObjectError + 4138, "CSectionTypeRegistry", _
            "Для выбранного параметрического сечения не задан ни один стержень арматуры. " & _
            "Допустимо оставить отдельные грани без арматуры, но в сечении должен быть хотя бы один стержень."
    End If

    Dim model As CSectionModel
    Dim modelBuilder As CSectionModelBuilder
    Set modelBuilder = New CSectionModelBuilder
    Set model = modelBuilder.BuildFromGenerated(mesh, rebars, "Generated")
    shape.BuildContours model.Contours
    BuildAnnotations model, shape, rebars, settings
    Set BuildGeneratedModel = model
End Function

' Возвращает число подъячеек для уточнения границы бетонной сетки.
' Настройка общая для всех встроенных форм и не относится к конкретной
' geometry/rebar/annotation реализации.
Public Function MeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    MeshBoundarySubdivisions = ReadMeshBoundarySubdivisions(settings)
End Function
