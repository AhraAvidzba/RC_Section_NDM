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
' сознательно не используются, чтобы новый тип сечения требовал не более трех
' новых специализированных классов и одной регистрации в этом registry.

Private Const TYPE_CIRCLE As String = "circle" ' Ключ круглого сечения в Geometry.Type.
Private Const TYPE_LSHAPE As String = "lshape" ' Ключ Г-образного сечения в Geometry.Type.
Private Const TYPE_ROUNDED_RECTANGLE As String = "roundedrectangle" ' Ключ скругленного прямоугольника.

' Возвращает нормализованный тип встроенной геометрии из Config.
' Все публичные методы registry пользуются одним и тем же значением, поэтому
' новая форма регистрируется в одном месте и не расползается по workbook-коду.
Private Function ActiveGeometryType(ByVal settings As CSystemSettingsReader) As String
    If settings Is Nothing Then Err.Raise vbObjectError + 4135, _
        "CSectionTypeRegistry", "Настройки Config не переданы."

    Dim geometryType As String
    geometryType = LCase$(Trim$(settings.GetString("Geometry.Type", "LShape")))
    If Len(Trim$(geometryType)) = 0 Then geometryType = TYPE_LSHAPE

    Select Case geometryType
        Case TYPE_CIRCLE, TYPE_LSHAPE, TYPE_ROUNDED_RECTANGLE
            ActiveGeometryType = geometryType
        Case Else
            Err.Raise vbObjectError + 4134, "CSectionTypeRegistry", _
                "Неподдерживаемый тип сечения Geometry.Type: " & settings.GetString("Geometry.Type", "LShape")
    End Select
End Function

' Создает математическую форму сечения для Geometry.Source = Generated.
' Здесь допустим shape-specific Select Case: это единственная точка регистрации
' встроенных форм. Ниже по цепочке расчетное ядро видит только ISectionGeometry.
Public Function CreateGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim circleGeometry As CGeometryCircle ' Конкретная математическая форма круга.
    Dim lshapeGeometry As CGeometryLShape ' Конкретная математическая форма Г-сечения.
    Dim roundedGeometry As CGeometryRoundedRectangle ' Конкретная математическая форма скругленного прямоугольника.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleGeometry = New CGeometryCircle
            circleGeometry.InitializeByDiameter units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), 0#, 0#
            Set CreateGeometry = circleGeometry

        Case TYPE_LSHAPE
            Set lshapeGeometry = New CGeometryLShape
            lshapeGeometry.Initialize units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), 0#, 0#
            Set CreateGeometry = lshapeGeometry

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedGeometry = New CGeometryRoundedRectangle
            roundedGeometry.Initialize units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Width", 300#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Height", 200#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopLeft", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusTopRight", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomRight", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.RadiusBottomLeft", 0#))
            Set CreateGeometry = roundedGeometry
    End Select
End Function

' Создает раскладку арматуры для активной встроенной формы.
' Registry выбирает нужный builder, а чтение shape-specific ключей Config
' выполняет сам builder. Так таблица граней LShape и правила дополнительных
' рядов не размазываются по общему registry.
Public Function CreateRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Dim circleRebars As CCircleRebarLayoutBuilder ' Генератор раскладки арматуры для кругового сечения.
    Dim lshapeRebars As CLShapeRebarLayoutBuilder ' Генератор раскладки арматуры для Г-сечения.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleRebars = New CCircleRebarLayoutBuilder
            Set CreateRebars = circleRebars.BuildFromSettings(settings, units)

        Case TYPE_LSHAPE
            Set lshapeRebars = New CLShapeRebarLayoutBuilder
            Set CreateRebars = lshapeRebars.BuildFromSettings(settings, units)

        Case TYPE_ROUNDED_RECTANGLE
            ' Для RoundedRectangle пока нет отдельного rebar builder-а. Это не
            ' ошибка геометрии: сечение строится как бетонное, а подписи
            ' арматуры отсутствуют естественным образом.
            Set CreateRebars = New CRebarLayout
    End Select
End Function

' Добавляет semantic-аннотации формы в готовую CSectionModel.
' AnnotationBuilder получает модельные координаты и смысловые группы арматуры;
' Excel/Chart-координаты по-прежнему рассчитывает только layout/plot layer.
' Если для будущего типа сечения annotation builder еще не создан, это не
' считается ошибкой: модель просто останется без размеров/подписей граней,
' а общие подписи схемы CSectionPlotter продолжит выводить сам.
Public Sub BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader)
    If model Is Nothing Then Err.Raise vbObjectError + 4136, "CSectionTypeRegistry", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 4137, "CSectionTypeRegistry", "Геометрия сечения не передана."

    Dim circleAnnotations As CCircleAnnotationBuilder ' Формирует semantic-размеры и подписи арматуры для круга.
    Dim lshapeAnnotations As CLShapeAnnotationBuilder ' Формирует semantic-аннотации для граней Г-сечения.
    Dim roundedAnnotations As CRoundedRectAnnotationBuilder ' Формирует semantic-аннотации скругленного прямоугольника.

    model.Annotations.Clear

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleAnnotations = New CCircleAnnotationBuilder
            circleAnnotations.Build model, geometry, rebars

        Case TYPE_LSHAPE
            Set lshapeAnnotations = New CLShapeAnnotationBuilder
            lshapeAnnotations.Build model, geometry, rebars

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedAnnotations = New CRoundedRectAnnotationBuilder
            roundedAnnotations.Build model, geometry, rebars
    End Select
End Sub

' Собирает Generated-модель целиком: форму, бетонную сетку, арматуру и
' semantic-аннотации. При добавлении нового типа меняется только registry и
' три специализированных класса формы; writer/reader/layout/plotter не трогаются.
Public Function BuildGeneratedModel(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CSectionModel
    Dim geometry As ISectionGeometry
    Set geometry = CreateGeometry(settings, units)

    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder

    Dim meshStepX As Double ' Шаг бетонной сетки по X во внутренних единицах.
    Dim meshStepY As Double ' Шаг бетонной сетки по Y; если ключа нет, сохраняем старое квадратное поведение.
    meshStepX = units.InputLengthToInternal(settings.GetDouble("Mesh.Step", 10#))
    meshStepY = units.InputLengthToInternal(settings.GetDouble("Mesh.StepY", settings.GetDouble("Mesh.Step", 10#)))
    mesh.BuildMesh geometry, meshStepX, meshStepY, 1, MeshBoundarySubdivisions(settings)

    Dim rebars As CRebarLayout
    Set rebars = CreateRebars(geometry, settings, units)

    Dim model As CSectionModel
    Dim modelBuilder As CSectionModelBuilder
    Set modelBuilder = New CSectionModelBuilder
    Set model = modelBuilder.BuildFromGenerated(mesh, rebars, "Generated")
    BuildAnnotations model, geometry, rebars, settings
    Set BuildGeneratedModel = model
End Function

' Возвращает число подъячеек для уточнения границы бетонной сетки.
' Настройка общая для всех встроенных форм и не относится к конкретной
' geometry/rebar/annotation реализации.
Public Function MeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    MeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If MeshBoundarySubdivisions < 1 Then MeshBoundarySubdivisions = 1
End Function
