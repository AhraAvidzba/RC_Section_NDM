VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestFakeAcadContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: имитация точного контура AutoCAD
' ==========================================================================
' Воспроизводит Region/Polyline с неупорядоченными фрагментами Explode,
' а также Line/Arc/Circle. Учитывает удаление копии и временных фрагментов
' для проверки сохранности исходника при успехе и ошибке импорта.
' Проверку интеграции с установленным AutoCAD выполняют отдельно.

Private mKind As String ' Имя вида подставного объекта AutoCAD.
Private mLayer As String ' Имя слоя тестового объекта.
Private mHandle As String ' Уникальный номер тестового объекта.
Private mCurves As Collection ' Отрезки и дуги, возвращаемые при разбиении тестового контура.
Private mStart(0 To 2) As Double ' Три мировые координаты начала отрезка, мм.
Private mEnd(0 To 2) As Double ' Три мировые координаты конца отрезка, мм.
Private mCenter(0 To 2) As Double ' Три мировые координаты центра дуги или окружности, мм.
Private mNormal(0 To 2) As Double ' Три компоненты нормали к плоскости объекта.
Private mRadius As Double ' Радиус тестовой дуги или окружности, мм.
Private mAngle As Double ' Угол тестовой дуги, рад.
Private mClosed As Boolean ' Признак замкнутого тестового контура.
Private mElevation As Double ' Высотная отметка объекта, мм.
Private mDeleted As Boolean ' Был ли вызван Delete для тестового объекта.
Private mLastCopy As CTestFakeAcadContour ' Последняя копия, созданная тестовым вызовом Copy.
Private mExploded As Collection ' Объекты, полученные последним тестовым разбиением контура.
Private mLegacyContourTag As Boolean ' ДЛЯ ТЕСТОВ: метки экспортированного контура не влияют на импорт.
Private mXDataReads As Long ' Число чтений расширенных данных объекта для проверки импорта.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Моделирует объект DWG с экспортной меткой, которую импорт не использует для фильтрации.
Public Sub TestSetLegacyContourTag(ByVal value As Boolean)
    mLegacyContourTag = value
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' число чтений XData объекта AutoCAD.
Public Property Get TestXDataReads() As Long
    TestXDataReads = mXDataReads
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' источник в памяти не имеет владельца DWG и его слоев.
' Блокировка настоящего слоя проверяется отдельно в native AutoCAD.
Public Property Get Document() As Object
    Set Document = Nothing
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает XData в формате Autodesk либо Empty у исходника.
Public Sub GetXData(ByVal appName As String, ByRef codes As Variant, ByRef values As Variant)
    codes = Empty: values = Empty
    mXDataReads = mXDataReads + 1
    If mLegacyContourTag And appName = "RC_SECTION_NDM" Then
        codes = Array(1001, 1000)
        values = Array("RC_SECTION_NDM", "MATERIAL_CONTOUR_OUTPUT")
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает составной Region/Polyline либо unsupported entity.
Public Sub TestInitializeLoop(ByVal kind As String, ByVal layerName As String, ByVal handleText As String, _
        ByVal curves As Collection, Optional ByVal closed As Boolean = True)
    mKind = kind: mLayer = layerName: mHandle = handleText
    Set mCurves = curves: mClosed = closed: mNormal(2) = 1#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает WCS-сегмент, в том числе неплоский отрицательный случай.
Public Sub TestInitializeLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, Optional ByVal z As Double = 0#)
    mKind = "AcDbLine": mNormal(2) = 1#
    mStart(0) = x1: mStart(1) = y1: mStart(2) = z
    mEnd(0) = x2: mEnd(1) = y2: mEnd(2) = z
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает точную круговую дугу со signed направлением нормали.
Public Sub TestInitializeArc(ByVal cx As Double, ByVal cy As Double, ByVal radius As Double, _
        ByVal startAngle As Double, ByVal sweep As Double)
    mKind = "AcDbArc": mCenter(0) = cx: mCenter(1) = cy: mRadius = radius
    mAngle = Abs(sweep): mNormal(2) = Sgn(sweep)
    mStart(0) = cx + radius * Cos(startAngle): mStart(1) = cy + radius * Sin(startAngle)
    mEnd(0) = cx + radius * Cos(startAngle + sweep): mEnd(1) = cy + radius * Sin(startAngle + sweep)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает Circle, которым Explode Region может представить отверстие.
Public Sub TestInitializeCircle(ByVal cx As Double, ByVal cy As Double, ByVal radius As Double)
    mKind = "AcDbCircle": mNormal(2) = 1#
    mCenter(0) = cx: mCenter(1) = cy: mRadius = radius
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' меняет нормаль/высоту для проверки запрета молчаливой проекции.
Public Sub TestSetPlane(ByVal nx As Double, ByVal ny As Double, ByVal nz As Double, ByVal elevation As Double)
    mNormal(0) = nx: mNormal(1) = ny: mNormal(2) = nz: mElevation = elevation
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Простые COM-свойства тестового объекта возвращают заданные значения.
Public Property Get ObjectName() As String
    ObjectName = mKind
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' имя слоя имитируемого объекта AutoCAD.
Public Property Get Layer() As String
    Layer = mLayer
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' устойчивый идентификатор имитируемого объекта AutoCAD.
Public Property Get Handle() As String
    Handle = mHandle
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак фактической замкнутости имитируемого объекта.
Public Property Get Closed() As Boolean
    Closed = mClosed
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' высота объекта относительно WCS, мм.
Public Property Get Elevation() As Double
    Elevation = mElevation
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' нормаль плоскости имитируемого объекта AutoCAD.
Public Property Get Normal() As Variant
    Normal = mNormal
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' координаты начала имитируемого сегмента в WCS.
Public Property Get StartPoint() As Variant
    StartPoint = mStart
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' координаты конца имитируемого сегмента в WCS.
Public Property Get EndPoint() As Variant
    EndPoint = mEnd
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' координаты центра имитируемой дуги в WCS.
Public Property Get Center() As Variant
    Center = mCenter
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' радиус имитируемой дуги, мм.
Public Property Get Radius() As Double
    Radius = mRadius
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' угол обхода имитируемой дуги, радианы.
Public Property Get TotalAngle() As Double
    TotalAngle = mAngle
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак удаления только имитируемого временного объекта.
Public Property Get TestDeleted() As Boolean
    TestDeleted = mDeleted
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' последняя временная копия имитируемого объекта.
Public Property Get TestLastCopy() As CTestFakeAcadContour
    Set TestLastCopy = mLastCopy
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' сегменты, возвращаемые разбиением имитируемого объекта.
Public Property Get TestExplodedPieces() As Collection
    Set TestExplodedPieces = mExploded
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' копирует фактические поля без совместного владения фрагментами.
Public Function Copy() As CTestFakeAcadContour
    Dim result As CTestFakeAcadContour
    Set result = New CTestFakeAcadContour
    result.TestInitializeLoop mKind, mLayer, mHandle, mCurves, mClosed
    result.TestCopyValues mStart, mEnd, mCenter, mNormal, mRadius, mAngle, mElevation
    Set mLastCopy = result
    Set Copy = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' принимает численные поля Copy через объявленный Friend API,
' не обращаясь к private-полям другого экземпляра VBA-класса.
Friend Sub TestCopyValues(ByVal startPoint As Variant, ByVal endPoint As Variant, ByVal center As Variant, _
        ByVal normal As Variant, ByVal radius As Double, ByVal angle As Double, ByVal elevation As Double)
    Dim i As Long
    For i = 0 To 2
        mStart(i) = startPoint(i): mEnd(i) = endPoint(i)
        mCenter(i) = center(i): mNormal(i) = normal(i)
    Next i
    mRadius = radius: mAngle = angle: mElevation = elevation
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создает новые собственные фрагменты, как настоящий Explode.
Public Function Explode() As Variant
    Dim result() As Variant, i As Long, curve As CTestFakeAcadContour
    If mCurves Is Nothing Then Err.Raise vbObjectError + 5501, "CTestFakeAcadContour", "Фрагменты Explode не заданы."
    ReDim result(0 To mCurves.Count - 1)
    Set mExploded = New Collection
    For i = 1 To mCurves.Count
        Set curve = mCurves(i)
        Set result(i - 1) = curve.Copy
        mExploded.Add result(i - 1)
    Next i
    Explode = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует удаление без обращения к живому чертежу.
Public Sub Delete()
    mDeleted = True
End Sub
