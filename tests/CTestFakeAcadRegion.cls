VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestFakeAcadRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: имитация области AutoCAD Region
' ==========================================================================
' Предоставляет площадь, центр, моменты инерции, слой, идентификатор
' и операции Copy/Explode/Delete для проверки импорта без запуска AutoCAD.

Private mArea As Double ' Реальная площадь тестового Region.
Private mX As Double ' Координата X центра тяжести Region.
Private mY As Double ' Координата Y центра тяжести Region.
Private mLocalIx As Double ' Центральный Ix Region, который должен попасть в CSectionModel.
Private mLocalIy As Double ' Центральный Iy Region, который должен попасть в CSectionModel.
Private mLocalIxy As Double ' Центральный Ixy Region, который должен попасть в CSectionModel.
Private mLayer As String ' Слой, по которому importer относит Region к бетону или арматуре.
Private mHandle As String ' Тестовый handle для трассировки.
Private mEdgeRotation As Double ' Угол простой грани, отдаваемой Explode.
Private mHasEdge As Boolean ' Есть ли у тестового Region пригодная прямая грань.
Private mDeleted As Boolean ' ДЛЯ ТЕСТОВ: был ли Delete вызван для этой сущности, а не ее Copy.
Private mGeometry As CConcreteRegion ' ДЛЯ ТЕСТОВ: полная аналитическая граница для нового контурного адаптера.
Private mUnsupportedCurve As String ' ДЛЯ ТЕСТОВ: явно неподдержанная кривая вместо подмены ее линиями.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Заполняет фактические свойства тестового Region.
Public Sub TestInitialize(ByVal area As Double, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal layerName As String, ByVal handleText As String, _
        Optional ByVal edgeRotation As Double = 0#, Optional ByVal hasEdge As Boolean = False)
    mArea = area
    mX = xCoord
    mY = yCoord
    mLocalIx = localIx
    mLocalIy = localIy
    mLocalIxy = localIxy
    mLayer = layerName
    mHandle = handleText
    mEdgeRotation = edgeRotation
    mHasEdge = hasEdge
    mDeleted = False
    Set mGeometry = Nothing: mUnsupportedCurve = vbNullString
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает точные loops независимо от COM-массовых характеристик fixture.
Public Sub TestSetGeometry(ByVal geometry As CConcreteRegion)
    Set mGeometry = geometry.Clone: mUnsupportedCurve = vbNullString
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет неподдержанный тип для проверки запрета эквивалентного fallback нового адаптера.
Public Sub TestSetUnsupportedCurve(ByVal kind As String)
    Set mGeometry = Nothing: mUnsupportedCurve = kind
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Тестовый Region лежит в WCS XY и не имеет пользовательского документа.
Public Property Get Normal() As Variant
    Normal = Array(0#, 0#, 1#)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Импорт без настоящего DWG не проверяет блокировку реальных слоев.
Public Property Get Document() As Object
    Set Document = Nothing
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает имя AutoCAD-типа, по которому importer распознает Region.
Public Property Get ObjectName() As String
    ObjectName = "AcDbRegion"
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Слой, по которому importer относит Region к бетону или арматуре.
Public Property Get Layer() As String
    Layer = mLayer
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Тестовый handle для трассировки.
Public Property Get Handle() As String
    Handle = mHandle
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Реальная площадь тестового Region.
Public Property Get Area() As Double
    Area = mArea
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает центр тяжести в формате, похожем на AutoCAD COM Variant-array.
Public Property Get Centroid() As Variant
    Dim result(0 To 2) As Double
    result(0) = mX
    result(1) = mY
    result(2) = 0#
    Centroid = result
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' AutoCAD отдает моменты относительно начала координат; importer сам снимает
' перенос Штейнера и должен получить сохраненные центральные I.
Public Property Get MomentOfInertia() As Variant
    Dim result(0 To 1) As Double
    result(0) = mLocalIx + mArea * mY * mY
    result(1) = mLocalIy + mArea * mX * mX
    MomentOfInertia = result
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет XY-компоненту тензора настоящего Autodesk Region: отрицательный
' интеграл x*y относительно начала координат. После смены знака и снятия
' переноса importer должен получить сохраненный центральный mLocalIxy.
Public Property Get ProductOfInertia() As Variant
    Dim result(0 To 0) As Double
    result(0) = -(mLocalIxy + mArea * mX * mY)
    ProductOfInertia = result
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Имитирует AutoCAD Copy перед Explode.
Public Function Copy() As CTestFakeAcadRegion
    Dim clone As CTestFakeAcadRegion
    Set clone = New CTestFakeAcadRegion
    clone.TestInitialize mArea, mX, mY, mLocalIx, mLocalIy, mLocalIxy, _
        mLayer, mHandle, mEdgeRotation, mHasEdge
    If Not mGeometry Is Nothing Then clone.TestSetGeometry mGeometry
    If Len(mUnsupportedCurve) > 0 Then clone.TestSetUnsupportedCurve mUnsupportedCurve
    Set Copy = clone
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает первую прямую грань с заданным углом.
Public Function Explode() As Variant
    If Not mGeometry Is Nothing Then
        Dim edges() As TRegionEdge, count As Long, i As Long, radius As Double, parts() As Variant, curve As CTestFakeAcadContour
        mGeometry.CopyEdges edges, count: ReDim parts(0 To count - 1)
        For i = 1 To count
            Set curve = New CTestFakeAcadContour
            If Abs(edges(i).Sweep) <= 0.000000000001 Then
                curve.TestInitializeLine edges(i).X1, edges(i).Y1, edges(i).X2, edges(i).Y2
            Else
                radius = Sqr((edges(i).X1 - edges(i).CenterX) ^ 2 + (edges(i).Y1 - edges(i).CenterY) ^ 2)
                curve.TestInitializeArc edges(i).CenterX, edges(i).CenterY, radius, _
                    modGeometryPrimitives.AngleOf(edges(i).X1 - edges(i).CenterX, edges(i).Y1 - edges(i).CenterY), edges(i).Sweep
            End If
            Set parts(i - 1) = curve
        Next i
        Explode = parts: Exit Function
    End If
    If Len(mUnsupportedCurve) > 0 Then
        Dim unsupported As CTestFakeAcadContour: Set unsupported = New CTestFakeAcadContour
        Dim noCurves As Collection: Set noCurves = New Collection
        unsupported.TestInitializeLoop mUnsupportedCurve, mLayer, mHandle & "_CURVE", noCurves
        Explode = Array(unsupported): Exit Function
    End If
    If Not mHasEdge Then
        Dim emptyResult() As Variant
        Explode = emptyResult
        Exit Function
    End If

    Dim line As CTestFakeAcadLine
    Set line = New CTestFakeAcadLine
    line.TestInitialize mX, mY, mX + Cos(mEdgeRotation), mY + Sin(mEdgeRotation)

    Dim result(0 To 0) As Variant
    Set result(0) = line
    Explode = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Имитирует удаление временной копии Region.
Public Sub Delete()
    mDeleted = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' ============================== ДЛЯ ТЕСТОВ ==============================
' Позволяет проверке очистки отличить сохраненный Region от удаленного,
' не создавая AutoCAD application или дополнительный тестовый класс.
Public Property Get TestDeleted() As Boolean
    TestDeleted = mDeleted
End Property
