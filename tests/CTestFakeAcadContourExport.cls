VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestFakeAcadContourExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: имитация создания контуров в AutoCAD
' ==========================================================================
' Воспроизводит операции экспорта с ModelSpace, полилиниями и Region.
' Запрещает просмотр существующих объектов через Count/Item, чтобы проверить
' экспорт без поиска совпадений и очистки. Сохраняет свойства созданных объектов
' для проверок без запуска AutoCAD.

Private mEntities As Collection ' Объекты, созданные тестовым экспортом.
Private mCoordinates As Variant ' Координаты вершин созданной полилинии.
Private mBulges As Object ' Коэффициенты дуг по номерам сегментов полилинии.
Private mLayer As String, mColor As Long, mClosed As Boolean, mDeleted As Boolean ' Слой, цвет, замкнутость и признак удаления тестового объекта.
Private mCodes As Variant, mValues As Variant ' Коды и значения расширенных данных тестового объекта.
Private mForbiddenReads As Long ' Число запрещённых обращений к исходным объектам чертежа.
Private mObjectName As String, mArea As Double ' Имя вида тестового объекта и его площадь в мм2.
Private mRegionCalls As Long, mFailRegionAt As Long ' Число попыток создания Region и номер попытки, на которой нужно имитировать отказ.
Private mNormal As Variant, mElevation As Double ' Плоскость новой полилинии, заданная экспортером в WCS.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создает пустой ModelSpace или отдельную полилинию.
Private Sub Class_Initialize()
    Set mEntities = New Collection: Set mBulges = CreateObject("Scripting.Dictionary")
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' добавляет новую полилинию, не читая существующие объекты.
Public Function AddLightWeightPolyline(ByVal coordinates As Variant) As CTestFakeAcadContourExport
    Dim entity As CTestFakeAcadContourExport: Set entity = New CTestFakeAcadContourExport
    entity.TestInitializePolyline coordinates
    mEntities.Add entity
    Set AddLightWeightPolyline = entity
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет фактический массив вершин экспортера.
Public Sub TestInitializePolyline(ByVal coordinates As Variant)
    mCoordinates = coordinates
    mObjectName = "AcDbPolyline"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Region получает геометрию своей новой полилинии, без обращения к существующим объектам.
Public Function AddRegion(ByVal curves As Variant) As Variant
    mRegionCalls = mRegionCalls + 1
    If mFailRegionAt = mRegionCalls Then Err.Raise vbObjectError + 9103, "CTestFakeAcadContourExport", "Injected AddRegion failure."
    Dim entity As CTestFakeAcadContourExport: Set entity = New CTestFakeAcadContourExport
    entity.TestInitializeRegion curves(LBound(curves))
    mEntities.Add entity
    AddRegion = Array(entity)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' независимая площадь полилинии учитывает круговые сегменты bulge.
Public Sub TestInitializeRegion(ByVal curve As CTestFakeAcadContourExport)
    mObjectName = "AcDbRegion": mCoordinates = curve.Coordinates
    Dim i As Long, j As Long, n As Long, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim bulge As Double, angle As Double, chordSquared As Double, radiusSquared As Double
    n = (UBound(mCoordinates) + 1) \ 2
    For i = 0 To n - 1
        j = (i + 1) Mod n
        x1 = CDbl(mCoordinates(2 * i)): y1 = CDbl(mCoordinates(2 * i + 1))
        x2 = CDbl(mCoordinates(2 * j)): y2 = CDbl(mCoordinates(2 * j + 1))
        mArea = mArea + (x1 * y2 - x2 * y1) / 2#
        bulge = curve.TestBulgeAt(i): mBulges(CStr(i)) = bulge
        If Abs(bulge) > 0.000000000001 Then
            angle = 4# * Atn(bulge): chordSquared = (x2 - x1) ^ 2 + (y2 - y1) ^ 2
            radiusSquared = chordSquared / (4# * Sin(angle / 2#) ^ 2)
            mArea = mArea + radiusSquared * (angle - Sin(angle)) / 2#
        End If
    Next i
    mArea = Abs(mArea)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Площадь выбранного элемента, мм^2.
Public Property Get Area() As Double
    Area = mArea
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет нормаль новой полилинии без обращения к пользовательскому DWG.
Public Property Let Normal(ByVal value As Variant)
    mNormal = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает фактическую нормаль, заданную рабочим экспортером.
Public Property Get Normal() As Variant
    Normal = mNormal
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет WCS-высоту новой полилинии для проверки независимости от чертежа.
Public Property Let Elevation(ByVal value As Double)
    mElevation = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает WCS-высоту новой полилинии, мм.
Public Property Get Elevation() As Double
    Elevation = mElevation
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' тип имитируемого объекта AutoCAD.
Public Property Get ObjectName() As String
    ObjectName = mObjectName
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' номер операции создания Region с принудительной ошибкой.
Public Property Let TestFailRegionAt(ByVal value As Long)
    mFailRegionAt = value
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' позволяет oracle читать записи без обращения к ModelSpace.Count.
Public Property Get TestCreatedEntities() As Collection
    Set TestCreatedEntities = mEntities
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' поиск или очистка существующих объектов является нарушением.
Public Property Get Count() As Long
    mForbiddenReads = mForbiddenReads + 1
    Err.Raise vbObjectError + 9100, "CTestFakeAcadContourExport", "Export must not scan ModelSpace.Count."
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запрет чтения существующих объектов дополняет запрет Count.
Public Function Item(ByVal index As Variant) As Object
    mForbiddenReads = mForbiddenReads + 1
    Err.Raise vbObjectError + 9101, "CTestFakeAcadContourExport", "Export must not read existing entities."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сообщает фактическое число запрещенных обращений.
Public Property Get TestForbiddenReads() As Long
    TestForbiddenReads = mForbiddenReads
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' свойства новой полилинии получают значения рабочего экспортера.
Public Property Let Layer(ByVal value As String)
    mLayer = value
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' имя слоя имитируемого объекта AutoCAD.
Public Property Get Layer() As String
    Layer = mLayer
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' цвет имитируемого экспортируемого объекта.
Public Property Let Color(ByVal value As Long)
    mColor = value
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' цвет имитируемого экспортируемого объекта.
Public Property Get Color() As Long
    Color = mColor
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак фактической замкнутости имитируемого объекта.
Public Property Let Closed(ByVal value As Boolean)
    mClosed = value
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак фактической замкнутости имитируемого объекта.
Public Property Get Closed() As Boolean
    Closed = mClosed
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' массив координат вершин имитируемой полилинии, мм.
Public Property Get Coordinates() As Variant
    Coordinates = mCoordinates
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' признак удаления только имитируемого временного объекта.
Public Property Get TestDeleted() As Boolean
    TestDeleted = mDeleted
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' фиксирует дуговой параметр и предоставляет независимое чтение.
Public Sub SetBulge(ByVal index As Long, ByVal value As Double)
    mBulges(CStr(index)) = value
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает заданный bulge сегмента полилинии для проверки экспорта аналитической
' дуги.
Public Function TestBulgeAt(ByVal index As Long) As Double
    If mBulges.Exists(CStr(index)) Then TestBulgeAt = mBulges(CStr(index))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Update не изменяет геометрию; Delete отмечает недопустимую очистку.
Public Sub Update()
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' отмечает удаление временного объекта без воздействия на настоящий AutoCAD.
Public Sub Delete()
    mDeleted = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' зарегистрированное приложение доступно для реальной записи XData.
Public Property Get Document() As Object
    Set Document = Me
End Property
' -------------------- ДЛЯ ТЕСТОВ --------------------
' зарегистрированные приложения XData имитируемого документа.
Public Property Get RegisteredApplications() As Object
    Dim apps As Object: Set apps = CreateObject("Scripting.Dictionary")
    apps.Add "RC_SECTION_NDM", Me
    Set RegisteredApplications = apps
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' сохраняет и возвращает данные XData с теми же признаками объектов, что при обычном
' экспорте.
Public Sub SetXData(ByVal codes As Variant, ByVal values As Variant)
    mCodes = codes: mValues = values
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает сохранённые XData и учитывает чтение для проверки полного обхода
' объектов.
Public Sub GetXData(ByVal appName As String, ByRef codes As Variant, ByRef values As Variant)
    codes = mCodes: values = mValues
End Sub
