VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeAcadRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Тестовая имитация AutoCAD Region
' ==========================================================================
' Класс нужен только unit-тестам importer-а. Он отдает минимальный набор
' свойств Region: площадь, центр, моменты инерции, произведение инерции, слой,
' handle, Copy/Explode/Delete. Так можно проверить ветвление importer-а без
' запуска AutoCAD.

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

' Заполняет фактические свойства тестового Region.
Public Sub Initialize(ByVal area As Double, ByVal xCoord As Double, ByVal yCoord As Double, _
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
End Sub

' Возвращает имя AutoCAD-типа, по которому importer распознает Region.
Public Property Get ObjectName() As String
    ObjectName = "AcDbRegion"
End Property

Public Property Get Layer() As String
    Layer = mLayer
End Property

Public Property Get Handle() As String
    Handle = mHandle
End Property

Public Property Get Area() As Double
    Area = mArea
End Property

' Возвращает центр тяжести в формате, похожем на AutoCAD COM Variant-array.
Public Property Get Centroid() As Variant
    Dim result(0 To 2) As Double
    result(0) = mX
    result(1) = mY
    result(2) = 0#
    Centroid = result
End Property

' AutoCAD отдает моменты относительно начала координат; importer сам снимает
' перенос Штейнера и должен получить сохраненные центральные I.
Public Property Get MomentOfInertia() As Variant
    Dim result(0 To 1) As Double
    result(0) = mLocalIx + mArea * mY * mY
    result(1) = mLocalIy + mArea * mX * mX
    MomentOfInertia = result
End Property

' Возвращает произведение инерции в варианте, который после снятия переноса
' дает mLocalIxy.
Public Property Get ProductOfInertia() As Variant
    Dim result(0 To 0) As Double
    result(0) = mLocalIxy + mArea * mX * mY
    ProductOfInertia = result
End Property

' Имитирует AutoCAD Copy перед Explode.
Public Function Copy() As CFakeAcadRegion
    Dim clone As CFakeAcadRegion
    Set clone = New CFakeAcadRegion
    clone.Initialize mArea, mX, mY, mLocalIx, mLocalIy, mLocalIxy, _
        mLayer, mHandle, mEdgeRotation, mHasEdge
    Set Copy = clone
End Function

' Возвращает первую прямую грань с заданным углом.
Public Function Explode() As Variant
    If Not mHasEdge Then
        Dim emptyResult() As Variant
        Explode = emptyResult
        Exit Function
    End If

    Dim line As CFakeAcadLine
    Set line = New CFakeAcadLine
    line.Initialize mX, mY, mX + Cos(mEdgeRotation), mY + Sin(mEdgeRotation)

    Dim result(0 To 0) As Variant
    Set result(0) = line
    Explode = result
End Function

' Имитирует удаление временной копии Region.
Public Sub Delete()
    mDeleted = True
End Sub

' ============================== ДЛЯ ТЕСТОВ ==============================
' Позволяет проверке очистки отличить сохраненный Region от удаленного,
' не создавая AutoCAD application или дополнительный тестовый класс.
Public Property Get Deleted() As Boolean
    Deleted = mDeleted
End Property
