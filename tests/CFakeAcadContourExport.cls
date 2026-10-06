VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeAcadContourExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: добавляющий ModelSpace и созданные контурные полилинии
' ==========================================================================
' Исполняет минимальный COM-контракт production exporter без AutoCAD.
' Чтение Count/Item запрещено: экспорт не должен искать существующие контуры
' или удалять прежний вывод. Созданные объекты сохраняют цвет, слой, дуги,
' замкнутость и XData для проверки. Это не приемка живой CAD-интеграции.

Private mEntities As Collection
Private mCoordinates As Variant
Private mBulges As Object
Private mLayer As String, mColor As Long, mClosed As Boolean, mDeleted As Boolean
Private mCodes As Variant, mValues As Variant
Private mForbiddenReads As Long

' ДЛЯ ТЕСТОВ: создает пустой ModelSpace или отдельную полилинию.
Private Sub Class_Initialize()
    Set mEntities = New Collection: Set mBulges = CreateObject("Scripting.Dictionary")
End Sub

' ДЛЯ ТЕСТОВ: добавляет новую полилинию, не читая прежние объекты.
Public Function AddLightWeightPolyline(ByVal coordinates As Variant) As CFakeAcadContourExport
    Dim entity As CFakeAcadContourExport: Set entity = New CFakeAcadContourExport
    entity.InitializePolyline coordinates
    mEntities.Add entity
    Set AddLightWeightPolyline = entity
End Function

' ДЛЯ ТЕСТОВ: сохраняет фактический массив вершин экспортера.
Public Sub InitializePolyline(ByVal coordinates As Variant)
    mCoordinates = coordinates
End Sub

' ДЛЯ ТЕСТОВ: позволяет oracle читать записи без обращения к ModelSpace.Count.
Public Property Get CreatedEntities() As Collection
    Set CreatedEntities = mEntities
End Property

' ДЛЯ ТЕСТОВ: поиск или очистка существующих объектов является нарушением.
Public Property Get Count() As Long
    mForbiddenReads = mForbiddenReads + 1
    Err.Raise vbObjectError + 9100, "CFakeAcadContourExport", "Export must not scan ModelSpace.Count."
End Property

' ДЛЯ ТЕСТОВ: запрет чтения существующих объектов дополняет запрет Count.
Public Function Item(ByVal index As Variant) As Object
    mForbiddenReads = mForbiddenReads + 1
    Err.Raise vbObjectError + 9101, "CFakeAcadContourExport", "Export must not read existing entities."
End Function

' ДЛЯ ТЕСТОВ: сообщает фактическое число запрещенных обращений.
Public Property Get ForbiddenReads() As Long
    ForbiddenReads = mForbiddenReads
End Property

' ДЛЯ ТЕСТОВ: свойства новой полилинии получают значения рабочего экспортера.
Public Property Let Layer(ByVal value As String)
    mLayer = value
End Property
Public Property Get Layer() As String
    Layer = mLayer
End Property
Public Property Let Color(ByVal value As Long)
    mColor = value
End Property
Public Property Get Color() As Long
    Color = mColor
End Property
Public Property Let Closed(ByVal value As Boolean)
    mClosed = value
End Property
Public Property Get Closed() As Boolean
    Closed = mClosed
End Property
Public Property Get Coordinates() As Variant
    Coordinates = mCoordinates
End Property
Public Property Get Deleted() As Boolean
    Deleted = mDeleted
End Property

' ДЛЯ ТЕСТОВ: фиксирует дуговой параметр и предоставляет независимое чтение.
Public Sub SetBulge(ByVal index As Long, ByVal value As Double)
    mBulges(CStr(index)) = value
End Sub
Public Function BulgeAt(ByVal index As Long) As Double
    If mBulges.Exists(CStr(index)) Then BulgeAt = mBulges(CStr(index))
End Function

' ДЛЯ ТЕСТОВ: Update не изменяет геометрию; Delete отмечает недопустимую очистку.
Public Sub Update()
End Sub
Public Sub Delete()
    mDeleted = True
End Sub

' ДЛЯ ТЕСТОВ: зарегистрированное приложение доступно для реальной записи XData.
Public Property Get Document() As Object
    Set Document = Me
End Property
Public Property Get RegisteredApplications() As Object
    Dim apps As Object: Set apps = CreateObject("Scripting.Dictionary")
    apps.Add "RC_SECTION_NDM", Me
    Set RegisteredApplications = apps
End Property

' ДЛЯ ТЕСТОВ: сохраняет и возвращает XData без подмены production маркировки.
Public Sub SetXData(ByVal codes As Variant, ByVal values As Variant)
    mCodes = codes: mValues = values
End Sub
Public Sub GetXData(ByVal appName As String, ByRef codes As Variant, ByRef values As Variant)
    codes = mCodes: values = mValues
End Sub
