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
    geometryType = LCase$(Trim$(settings.GetString("Geometry.Type", "RectSet")))
    If Len(Trim$(geometryType)) = 0 Then geometryType = TYPE_RECTSET

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
' встроенных форм. Ниже по цепочке расчетное ядро видит только ISectionGeometry.
Public Function CreateGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
    Dim circleGeometry As CGeometryCircle ' Конкретная математическая форма круга.
    Dim rectSetGeometry As CGeometryRectSet ' Математическая форма из одного/двух прямоугольников.
    Dim roundedGeometry As CGeometryRoundedRectangle ' Конкретная математическая форма скругленного прямоугольника.
    Dim hollowGeometry As CGeometryHollowRectangle ' Пустотелая форма: внешний RoundedRectangle минус Opening.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleGeometry = New CGeometryCircle
            circleGeometry.InitializeByDiameter units.InputLengthToInternal(settings.GetDouble("Circle.Diameter", 300#)), 0#, 0#
            Set CreateGeometry = circleGeometry

        Case TYPE_RECTSET
            Set rectSetGeometry = New CGeometryRectSet
            rectSetGeometry.Initialize units.InputLengthToInternal(settings.GetDouble("RectSet.B1", 250#)), _
                units.InputLengthToInternal(settings.GetDouble("RectSet.H1", 550#)), _
                units.InputLengthToInternal(settings.GetDouble("RectSet.B2", 600#)), _
                units.InputLengthToInternal(settings.GetDouble("RectSet.H2", 250#)), 0#, 0#, _
                units.InputLengthToInternal(settings.GetDouble("RectSet.UpperOffsetX", 0#)), _
                settings.GetString("RectSet.SectionType", "LSection")
            Set CreateGeometry = rectSetGeometry

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedGeometry = New CGeometryRoundedRectangle
            roundedGeometry.InitializeSides _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.B", 300#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.H", 200#)), _
                settings.GetString("RoundedRectangle.Left.Type", "Simple"), _
                settings.GetString("RoundedRectangle.Right.Type", "Simple"), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.W", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.W", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.R1", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.R1", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Left.R2", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("RoundedRectangle.Right.R2", 0#))
            Set CreateGeometry = roundedGeometry

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowGeometry = New CGeometryHollowRectangle
            hollowGeometry.Initialize _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.B", 500#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.H", 800#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.R", 180#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningB", 200#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningH", 500#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.OpeningR", 30#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.InnerOffsetX", 0#)), _
                units.InputLengthToInternal(settings.GetDouble("HollowRectangle.InnerOffsetY", 0#))
            Set CreateGeometry = hollowGeometry
    End Select
End Function

' Создает раскладку арматуры для активной встроенной формы.
' Registry выбирает нужный builder, а чтение shape-specific ключей Config
' выполняет сам builder. Так таблица граней RectSet и правила дополнительных
' рядов не размазываются по общему registry.
Public Function CreateRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem) As CRebarLayout
    Dim circleRebars As CCircleRebarLayoutBuilder ' Генератор раскладки арматуры для кругового сечения.
    Dim rectSetRebars As CRectSetRebarLayoutBuilder ' Генератор раскладки арматуры для RectSet-сечения.
    Dim roundedRebars As CRoundedRectRebarLayoutBuilder ' Генератор арматуры для RoundedRectangle-сечения.
    Dim hollowRebars As CHollowRectRebarLayoutBuilder ' Генератор арматуры для HollowRectangle-сечения.

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleRebars = New CCircleRebarLayoutBuilder
            Set CreateRebars = circleRebars.BuildFromSettings(settings, units)

        Case TYPE_RECTSET
            Set rectSetRebars = New CRectSetRebarLayoutBuilder
            Set CreateRebars = rectSetRebars.BuildFromSettings(settings, units)

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedRebars = New CRoundedRectRebarLayoutBuilder
            Set CreateRebars = roundedRebars.BuildFromSettings(settings, units)

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowRebars = New CHollowRectRebarLayoutBuilder
            Set CreateRebars = hollowRebars.BuildFromSettings(settings, units)
    End Select
End Function

' Добавляет semantic-аннотации формы в готовую CSectionModel.
' AnnotationBuilder получает модельные координаты и смысловые группы арматуры;
' Excel/Chart-координаты по-прежнему рассчитывает только layout/plot layer.
' Если зарегистрированная форма не имеет annotation builder, это не
' считается ошибкой: модель остается без размеров/подписей граней,
' а общие подписи схемы CSectionPlotter продолжит выводить сам.
Public Sub BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, _
        ByVal rebars As CRebarLayout, ByVal settings As CSystemSettingsReader)
    If model Is Nothing Then Err.Raise vbObjectError + 4136, "CSectionTypeRegistry", "Модель сечения не передана."
    If geometry Is Nothing Then Err.Raise vbObjectError + 4137, "CSectionTypeRegistry", "Геометрия сечения не передана."

    Dim circleAnnotations As CCircleAnnotationBuilder ' Формирует semantic-размеры и подписи арматуры для круга.
    Dim rectSetAnnotations As CRectSetAnnotationBuilder ' Формирует semantic-аннотации для граней RectSet-сечения.
    Dim roundedAnnotations As CRoundedRectAnnotationBuilder ' Формирует semantic-аннотации скругленного прямоугольника.
    Dim hollowAnnotations As CHollowRectAnnotationBuilder ' Формирует semantic-аннотации наружного контура и Opening.

    model.Annotations.Clear

    Select Case ActiveGeometryType(settings)
        Case TYPE_CIRCLE
            Set circleAnnotations = New CCircleAnnotationBuilder
            circleAnnotations.Build model, geometry, rebars

        Case TYPE_RECTSET
            Set rectSetAnnotations = New CRectSetAnnotationBuilder
            rectSetAnnotations.Build model, geometry, rebars

        Case TYPE_ROUNDED_RECTANGLE
            Set roundedAnnotations = New CRoundedRectAnnotationBuilder
            roundedAnnotations.Build model, geometry, rebars

        Case TYPE_HOLLOW_RECTANGLE
            Set hollowAnnotations = New CHollowRectAnnotationBuilder
            hollowAnnotations.Build model, geometry, rebars
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

    Dim meshStepX As Double ' Шаг бетонной сетки по X во внутренних единицах; старый Mesh.Step читается только для совместимости.
    Dim meshStepY As Double ' Шаг бетонной сетки по Y; если ключа нет, сохраняем старое квадратное поведение.
    meshStepX = units.InputLengthToInternal(settings.GetDouble("Mesh.StepX", settings.GetDouble("Mesh.Step", 10#)))
    meshStepY = units.InputLengthToInternal(settings.GetDouble("Mesh.StepY", settings.GetDouble("Mesh.StepX", settings.GetDouble("Mesh.Step", 10#))))
    mesh.BuildMesh geometry, meshStepX, meshStepY, 1, MeshBoundarySubdivisions(settings)

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
    BuildAnnotations model, geometry, rebars, settings
    Set BuildGeneratedModel = model
End Function

' Возвращает число подъячеек для уточнения границы бетонной сетки.
' Настройка общая для всех встроенных форм и не относится к конкретной
' geometry/rebar/annotation реализации.
Public Function MeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    If settings Is Nothing Then Err.Raise vbObjectError + 4139, "CSectionTypeRegistry", "Настройки разбиения бетонной сетки не переданы."
    If settings.HasKey("Mesh.BoundarySubdivisions") Then
        Dim rawValue As String
        rawValue = Trim$(settings.GetRawString("Mesh.BoundarySubdivisions"))
        If Len(rawValue) = 0 Or UCase$(rawValue) = "TODO" Then Err.Raise vbObjectError + 4139, "CSectionTypeRegistry", _
            "Mesh.BoundarySubdivisions не задан: введите целое число не меньше 1 на листе Config."
    End If
    MeshBoundarySubdivisions = settings.GetLong("Mesh.BoundarySubdivisions", 1)
    If MeshBoundarySubdivisions < 1 Then Err.Raise vbObjectError + 4139, "CSectionTypeRegistry", _
        "Mesh.BoundarySubdivisions должен быть целым числом не меньше 1. Проверьте значение на листе Config."
End Function
