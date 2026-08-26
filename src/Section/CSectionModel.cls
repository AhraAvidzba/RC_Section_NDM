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
' Унифицированная расчетная модель сечения
' ==========================================================================
' CSectionModel - это единственный формат геометрии, который видит расчетное
' ядро. Неважно, пришла сетка из встроенного генератора или из AutoCAD import:
' ниже по цепочке solver получает один и тот же набор бетонных элементов,
' стержней арматуры и semantic-аннотаций для вывода.
'
' Важно: расчетные имена элементов C1/R1 создаются здесь. Генераторы
' геометрии и импорт из AutoCAD не должны придумывать собственный формат имен
' для расчетного ядра.

Private mSourceType As String ' Диагностическая метка источника модели: Generated, AutoCADImport и т.п.

Private mConcreteID() As String ' Расчетные имена бетонных элементов: C1, C2, ...
Private mConcreteSourceName() As String ' Необязательное исходное имя области для трассировки.
Private mConcreteSourceHandle() As String ' Необязательный AutoCAD handle для поиска исходного Region.
Private mConcreteX() As Double ' X центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteY() As Double ' Y центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteArea() As Double ' Площадь бетонного элемента, мм2.
Private mConcreteMaterialID() As Long ' Номер/тип материала бетона для будущего расширения на разные материалы.
Private mConcreteShapeType() As String ' Упрощенная форма элемента для вывода: Rectangle, Circle и т.п.
Private mConcreteWidth() As Double ' Ширина эквивалентного прямоугольника элемента для схемы/AutoCAD export, мм.
Private mConcreteHeight() As Double ' Высота эквивалентного прямоугольника элемента для схемы/AutoCAD export, мм.
Private mConcreteRotation() As Double ' Поворот эквивалентной области, рад.
Private mConcreteLocalIx() As Double ' Собственный Ix элемента вокруг его центра, мм4.
Private mConcreteLocalIy() As Double ' Собственный Iy элемента вокруг его центра, мм4.
Private mConcreteLocalIxy() As Double ' Собственный центробежный момент элемента, мм4.
Private mConcreteComment() As String ' Короткая диагностическая заметка для Results.
Private mConcreteCount As Long ' Фактическое число бетонных элементов в модели.

Private mRebarID() As String ' Расчетные имена стержней: R1, R2, ...
Private mRebarSourceName() As String ' Необязательное исходное имя арматурной области.
Private mRebarSourceHandle() As String ' Необязательный AutoCAD handle арматурного Region.
Private mRebarX() As Double ' X центра стержня, мм.
Private mRebarY() As Double ' Y центра стержня, мм.
Private mRebarArea() As Double ' Площадь стержня, мм2.
Private mRebarDiameter() As Double ' Диаметр стержня, мм.
Private mRebarSteelClass() As String ' Расчетная подпись арматуры; сейчас хранит профиль стержней для трассировки.
Private mRebarMaterialID() As Long ' Номер/тип материала арматуры для будущего расширения.
Private mRebarComment() As String ' Диагностическая заметка о грани, ряде или источнике стержня.
Private mRebarCount As Long ' Фактическое число стержней в модели.

Private mAnnotations As CSectionAnnotations ' Semantic-аннотации, сформированные генератором сечения.

' Очищает модель перед новым построением сечения.
' После Clear объект снова готов принимать бетонные элементы, арматуру и
' аннотации без следов предыдущего расчета.
Public Sub Clear()
    mSourceType = vbNullString

    Erase mConcreteID
    Erase mConcreteSourceName
    Erase mConcreteSourceHandle
    Erase mConcreteX
    Erase mConcreteY
    Erase mConcreteArea
    Erase mConcreteMaterialID
    Erase mConcreteShapeType
    Erase mConcreteWidth
    Erase mConcreteHeight
    Erase mConcreteRotation
    Erase mConcreteLocalIx
    Erase mConcreteLocalIy
    Erase mConcreteLocalIxy
    Erase mConcreteComment
    mConcreteCount = 0

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
    mRebarCount = 0

    Set mAnnotations = New CSectionAnnotations
End Sub

Public Property Let SourceType(ByVal value As String)
    mSourceType = Trim$(value)
End Property

Public Property Get SourceType() As String
    SourceType = mSourceType
End Property

' Добавляет одно бетонное волокно или импортированную бетонную область.
' Координаты и площадь уже должны быть во внутренних единицах проекта. Метод
' только проверяет минимальную физическую корректность и присваивает элементу
' стабильное расчетное имя Cn.
Public Sub AddConcreteElement(ByVal xCoord As Double, ByVal yCoord As Double, ByVal area As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal shapeType As String = "Rectangle", _
        Optional ByVal width As Double = 0#, Optional ByVal height As Double = 0#, _
        Optional ByVal rotation As Double = 0#, Optional ByVal comment As String = vbNullString, _
        Optional ByVal localIx As Double = 0#, Optional ByVal localIy As Double = 0#, _
        Optional ByVal localIxy As Double = 0#)
    If area <= 0# Then Err.Raise vbObjectError + 5200, "CSectionModel", "Площадь бетонного элемента должна быть положительной."

    ' Массивы хранятся параллельно, чтобы горячие расчетные циклы работали без
    ' коллекций и COM-объектов на каждом элементе.
    mConcreteCount = mConcreteCount + 1
    ReDim Preserve mConcreteID(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceName(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceHandle(1 To mConcreteCount)
    ReDim Preserve mConcreteX(1 To mConcreteCount)
    ReDim Preserve mConcreteY(1 To mConcreteCount)
    ReDim Preserve mConcreteArea(1 To mConcreteCount)
    ReDim Preserve mConcreteMaterialID(1 To mConcreteCount)
    ReDim Preserve mConcreteShapeType(1 To mConcreteCount)
    ReDim Preserve mConcreteWidth(1 To mConcreteCount)
    ReDim Preserve mConcreteHeight(1 To mConcreteCount)
    ReDim Preserve mConcreteRotation(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIx(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIy(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIxy(1 To mConcreteCount)
    ReDim Preserve mConcreteComment(1 To mConcreteCount)

    mConcreteID(mConcreteCount) = "C" & CStr(mConcreteCount)
    mConcreteSourceName(mConcreteCount) = sourceName
    mConcreteSourceHandle(mConcreteCount) = sourceHandle
    mConcreteX(mConcreteCount) = xCoord
    mConcreteY(mConcreteCount) = yCoord
    mConcreteArea(mConcreteCount) = area
    mConcreteMaterialID(mConcreteCount) = materialID
    mConcreteShapeType(mConcreteCount) = shapeType
    mConcreteWidth(mConcreteCount) = width
    mConcreteHeight(mConcreteCount) = height
    mConcreteRotation(mConcreteCount) = rotation
    mConcreteLocalIx(mConcreteCount) = localIx
    mConcreteLocalIy(mConcreteCount) = localIy
    mConcreteLocalIxy(mConcreteCount) = localIxy
    mConcreteComment(mConcreteCount) = comment
End Sub

' Добавляет один стержень арматуры.
' Если площадь не передана явно, она считается по диаметру. Это удобно для
' встроенных раскладчиков, а AutoCAD import может передать площадь Region
' напрямую, если она известна надежнее диаметра.
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

    ' Нумерация R1/R2/... создается здесь, чтобы у всех источников геометрии
    ' был одинаковый формат расчетных имен.
    mRebarCount = mRebarCount + 1
    ReDim Preserve mRebarID(1 To mRebarCount)
    ReDim Preserve mRebarSourceName(1 To mRebarCount)
    ReDim Preserve mRebarSourceHandle(1 To mRebarCount)
    ReDim Preserve mRebarX(1 To mRebarCount)
    ReDim Preserve mRebarY(1 To mRebarCount)
    ReDim Preserve mRebarArea(1 To mRebarCount)
    ReDim Preserve mRebarDiameter(1 To mRebarCount)
    ReDim Preserve mRebarSteelClass(1 To mRebarCount)
    ReDim Preserve mRebarMaterialID(1 To mRebarCount)
    ReDim Preserve mRebarComment(1 To mRebarCount)

    mRebarID(mRebarCount) = "R" & CStr(mRebarCount)
    mRebarSourceName(mRebarCount) = sourceName
    mRebarSourceHandle(mRebarCount) = sourceHandle
    mRebarX(mRebarCount) = xCoord
    mRebarY(mRebarCount) = yCoord
    mRebarArea(mRebarCount) = actualArea
    mRebarDiameter(mRebarCount) = diameter
    mRebarSteelClass(mRebarCount) = steelClass
    mRebarMaterialID(mRebarCount) = materialID
    mRebarComment(mRebarCount) = comment
End Sub

Public Property Get ConcreteCount() As Long
    ConcreteCount = mConcreteCount
End Property

Public Property Get RebarCount() As Long
    RebarCount = mRebarCount
End Property

Public Property Get ConcreteID(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteID = mConcreteID(index)
End Property

Public Property Get ConcreteSourceName(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceName = mConcreteSourceName(index)
End Property

Public Property Get ConcreteSourceHandle(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceHandle = mConcreteSourceHandle(index)
End Property

Public Property Get ConcreteX(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteX = mConcreteX(index)
End Property

Public Property Get ConcreteY(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteY = mConcreteY(index)
End Property

Public Property Get ConcreteArea(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteArea = mConcreteArea(index)
End Property

Public Property Get ConcreteMaterialID(ByVal index As Long) As Long
    CheckConcreteIndex index
    ConcreteMaterialID = mConcreteMaterialID(index)
End Property

Public Property Get ConcreteShapeType(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteShapeType = mConcreteShapeType(index)
End Property

Public Property Get ConcreteWidth(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteWidth = mConcreteWidth(index)
End Property

Public Property Get ConcreteHeight(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteHeight = mConcreteHeight(index)
End Property

Public Property Get ConcreteRotation(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteRotation = mConcreteRotation(index)
End Property

Public Property Get ConcreteLocalIx(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIx = mConcreteLocalIx(index)
End Property

Public Property Get ConcreteLocalIy(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIy = mConcreteLocalIy(index)
End Property

Public Property Get ConcreteLocalIxy(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteLocalIxy = mConcreteLocalIxy(index)
End Property

Public Property Get ConcreteComment(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteComment = mConcreteComment(index)
End Property

Public Property Get RebarID(ByVal index As Long) As String
    CheckRebarIndex index
    RebarID = mRebarID(index)
End Property

Public Property Get RebarSourceHandle(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceHandle = mRebarSourceHandle(index)
End Property

Public Property Get RebarX(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarX = mRebarX(index)
End Property

Public Property Get RebarY(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarY = mRebarY(index)
End Property

Public Property Get RebarArea(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarArea = mRebarArea(index)
End Property

Public Property Get RebarDiameter(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarDiameter = mRebarDiameter(index)
End Property

Public Property Get RebarMaterialID(ByVal index As Long) As Long
    CheckRebarIndex index
    RebarMaterialID = mRebarMaterialID(index)
End Property

Public Property Get RebarComment(ByVal index As Long) As String
    CheckRebarIndex index
    RebarComment = mRebarComment(index)
End Property

Public Property Get Annotations() As CSectionAnnotations
    If mAnnotations Is Nothing Then Set mAnnotations = New CSectionAnnotations
    Set Annotations = mAnnotations
End Property

Public Property Get AnnotationCount() As Long
    If mAnnotations Is Nothing Then
        AnnotationCount = 0
    Else
        AnnotationCount = mAnnotations.Count
    End If
End Property

Private Sub CheckConcreteIndex(ByVal index As Long)
    If index < 1 Or index > mConcreteCount Then Err.Raise vbObjectError + 5220, "CSectionModel", "Индекс бетонного элемента вне допустимого диапазона."
End Sub

Private Sub CheckRebarIndex(ByVal index As Long)
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 5221, "CSectionModel", "Индекс арматурного элемента вне допустимого диапазона."
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти трассировочные свойства нужны regression-тестам AutoCAD/import и старых
' сценариев, где проверяется перенос исходных имен и профиля арматуры внутрь
' CSectionModel. Обычные расчеты берут рабочие свойства стержня через ID,
' координаты, площадь, диаметр, материал и комментарий.

Public Property Get RebarSourceName(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceName = mRebarSourceName(index)
End Property

Public Property Get RebarSteelClass(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSteelClass = mRebarSteelClass(index)
End Property
