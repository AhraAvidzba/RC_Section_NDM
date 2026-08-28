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
' *RebarLayoutBuilder и *AnnotationBuilder. Отдельные provider-классы здесь
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
' Логика принадлежности стержней граням остается внутри специализированного
' RebarLayoutBuilder; registry только передает уже нормализованные настройки.
Public Function CreateRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Dim circleRebars As CCircleRebarLayoutBuilder ' Генератор раскладки арматуры для кругового сечения.
    Dim lshapeRebars As CLShapeRebarLayoutBuilder ' Генератор раскладки арматуры для Г-сечения.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleRebars = New CCircleRebarLayoutBuilder
            Set CreateRebars = circleRebars.Build( _
                units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), _
                0#, _
                0#, _
                units.InputLengthToInternal(settings.GetDouble("Rebar.AxisDistance", 40#)), _
                settings.GetLong("Rebar.Count", 8), _
                units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter", 20#)), _
                settings.GetString("Steel.RebarProfile", "Ribbed"), _
                units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter2", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("Rebar.Diameter3", 0#)), _
                settings.GetString("Rebar.Loc2row", "Stacked"), _
                settings.GetString("Rebar.Loc3row", "Stacked"))

        Case TYPE_LSHAPE
            Set lshapeRebars = New CLShapeRebarLayoutBuilder
            Set CreateRebars = lshapeRebars.Build( _
                units.InputLengthToInternal(settings.GetDouble("LShape.B1", 250#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.H1", 550#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.B2", 600#)), _
                units.InputLengthToInternal(settings.GetDouble("LShape.H2", 250#)), _
                0#, _
                0#, _
                LShapeFaceSettings(settings, units, "H1"), _
                LShapeFaceSettings(settings, units, "H2"), _
                LShapeFaceSettings(settings, units, "B1"), _
                LShapeFaceSettings(settings, units, "B2"), _
                settings.GetString("Steel.RebarProfile", "Ribbed"))

        Case TYPE_ROUNDED_RECTANGLE
            Err.Raise vbObjectError + 4103, "CSectionTypeRegistry", _
                "Автоматическая раскладка арматуры поддерживается только для Circle и LShape."
    End Select
End Function

' Добавляет semantic-аннотации формы в готовую CSectionModel.
' AnnotationBuilder получает модельные координаты и смысловые группы арматуры;
' Excel/Chart-координаты по-прежнему рассчитывает только layout/plot layer.
Public Sub BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader)
    Dim circleAnnotations As CCircleAnnotationBuilder ' Формирует semantic-размеры и подписи арматуры для круга.
    Dim lshapeAnnotations As CLShapeAnnotationBuilder ' Формирует semantic-аннотации для граней Г-сечения.
    Dim roundedAnnotations As CRoundedRectAnnotationBuilder ' Формирует semantic-аннотации скругленного прямоугольника.

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

    Dim meshStep As Double
    meshStep = units.InputLengthToInternal(settings.GetDouble("Mesh.Step", 10#))
    mesh.BuildMesh geometry, meshStep, meshStep, 1, MeshBoundarySubdivisions(settings)

    Dim rebars As CRebarLayout
    Set rebars = CreateRebars(geometry, settings, units)

    Dim model As CSectionModel
    Set model = BuildGeneratedSectionModel(mesh, rebars, "Generated")
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

' Читает параметры одной смысловой грани Г-сечения и переводит длины во
' внутренние миллиметры. Дальше CLShapeRebarLayoutBuilder сам решает, какие
' стержни создать и какие semantic-группы арматуры сформировать.
Private Function LShapeFaceSettings(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal faceName As String) As Variant
    LShapeFaceSettings = Array( _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_1", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".as_2", 40#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_1", 32#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2", 32#)), _
        settings.GetLong("LShape." & faceName & ".n_1", 0), _
        settings.GetLong("LShape." & faceName & ".n_2", 0), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset1", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".StartOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".EndOffset2", 80#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_2row_2", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_1", 0#)), _
        units.InputLengthToInternal(settings.GetDouble("LShape." & faceName & ".d_3row_2", 0#)), _
        settings.GetString("LShape." & faceName & ".loc_2row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".loc_3row", "Stacked"), _
        settings.GetString("LShape." & faceName & ".bind_2row", "EachBar"), _
        settings.GetString("LShape." & faceName & ".bind_3row", "EachBar"))
End Function

