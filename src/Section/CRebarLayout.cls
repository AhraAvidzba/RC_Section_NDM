VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Промежуточный список стержней до CSectionModel
' ==========================================================================
' CRebarLayout хранит координаты, диаметры и диагностические подписи стержней,
' сформированных раскладчиком. Затем CSectionModelBuilder переносит эти данные
' в единую расчетную модель и присваивает расчетные имена R1/R2/....

Private mBarID() As String
Private mX() As Double
Private mY() As Double
Private mDiameter() As Double
Private mArea() As Double
Private mSteelClass() As String
Private mComment() As String
Private mBarAnnotationGroupName() As String
Private mCount As Long
Private mCapacity As Long
Private mIDs As Object
Private mPositionBuckets As Object
Private mMemberships As Object

Private mAnnotationGroupName() As String
Private mAnnotationStartX() As Double
Private mAnnotationStartY() As Double
Private mAnnotationEndX() As Double
Private mAnnotationEndY() As Double
Private mAnnotationNormalX() As Double
Private mAnnotationNormalY() As Double
Private mAnnotationAxisDistance() As Double
Private mAnnotationCount As Long

' Удаляет промежуточные стержни и аннотации линий раскладки целиком.
' Оба счетчика сбрасываются вместе с массивами, чтобы следующая генерация
' не смешивала координаты или подписи с предыдущей геометрией.
Public Sub Clear()
    Erase mBarID
    Erase mX
    Erase mY
    Erase mDiameter
    Erase mArea
    Erase mSteelClass
    Erase mComment
    Erase mBarAnnotationGroupName
    mCount = 0
    mCapacity = 0
    Set mIDs = Nothing
    Set mPositionBuckets = Nothing
    Set mMemberships = Nothing

    Erase mAnnotationGroupName
    Erase mAnnotationStartX
    Erase mAnnotationStartY
    Erase mAnnotationEndX
    Erase mAnnotationEndY
    Erase mAnnotationNormalX
    Erase mAnnotationNormalY
    Erase mAnnotationAxisDistance
    mAnnotationCount = 0
End Sub

' Добавляет стержень в промежуточную раскладку с проверкой ID, площади, диаметра и положения в бетоне.
Public Sub AddBar(ByVal barName As String, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal comment As String = vbNullString, Optional ByVal geometry As CGeometryRegion = Nothing, _
        Optional ByVal annotationGroupName As String = vbNullString)

    If Len(Trim$(barName)) = 0 Then Err.Raise vbObjectError + 3100, "CRebarLayout", "ID стержня не заполнен."
    If diameter <= 0# Then Err.Raise vbObjectError + 3101, "CRebarLayout", "Диаметр арматуры должен быть положительным."
    ' Общая защита фактического количества, в том числе для проекций Opening
    ' и импорта; builders дополнительно проверяют свой запрос до размещения.
    If CDbl(mCount) >= GEOM_MAX_REBAR_COUNT Then Err.Raise vbObjectError + 3108, "CRebarLayout", _
        "Количество стержней превышает диапазон счетчиков программы. Уменьшите активную раскладку в Config или число импортируемых элементов."

    Dim actualArea As Double
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 3102, "CRebarLayout", "Площадь арматуры должна быть положительной."

    EnsureIndexes
    If mIDs.Exists(barName) Then Err.Raise vbObjectError + 3103, "CRebarLayout", "Дублирующийся ID стержня: " & barName
    If FindBarAt(xCoord, yCoord) > 0 Then Err.Raise vbObjectError + 3104, "CRebarLayout", "Дублирующиеся координаты стержней."

    If Not geometry Is Nothing Then
        If Not geometry.ContainsPoint(xCoord, yCoord) Then Err.Raise vbObjectError + 3105, "CRebarLayout", "Центр стержня находится вне бетонного сечения."
    End If

    EnsureCapacity
    mCount = mCount + 1

    mBarID(mCount) = barName
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mDiameter(mCount) = diameter
    mArea(mCount) = actualArea
    mSteelClass(mCount) = steelClass
    mComment(mCount) = comment
    mBarAnnotationGroupName(mCount) = Trim$(annotationGroupName)
    mIDs.Add barName, mCount
    IndexPosition mCount
    AddMembership mCount, annotationGroupName
End Sub

' Создаёт индексы координат и смысловых групп для поиска совпавших физических стержней.
Private Sub EnsureIndexes()
    If Not mIDs Is Nothing Then Exit Sub
    Set mIDs = CreateObject("Scripting.Dictionary")
    mIDs.CompareMode = vbTextCompare
    Set mPositionBuckets = CreateObject("Scripting.Dictionary")
    Set mMemberships = CreateObject("Scripting.Dictionary")
End Sub

' Расширяет массивы арматуры блоками, сохраняя опубликованные стержни.
Private Sub EnsureCapacity()
    If mCount < mCapacity Then Exit Sub
    mCapacity = CLng(GeomMin(GEOM_MAX_REBAR_COUNT, GeomMax(16#, CDbl(mCapacity) * 2#)))
    ReDim Preserve mBarID(1 To mCapacity)
    ReDim Preserve mX(1 To mCapacity)
    ReDim Preserve mY(1 To mCapacity)
    ReDim Preserve mDiameter(1 To mCapacity)
    ReDim Preserve mArea(1 To mCapacity)
    ReDim Preserve mSteelClass(1 To mCapacity)
    ReDim Preserve mComment(1 To mCapacity)
    ReDim Preserve mBarAnnotationGroupName(1 To mCapacity)
End Sub

' Квантование служит только индексом кандидатов. Решение о совпадении
' всегда принимается по исходным координатам с прежним GEOM_TOLERANCE.
Private Function PositionBin(ByVal coordinate As Double) As Double
    If Abs(coordinate) > 1E+290 Then
        PositionBin = coordinate
    Else
        PositionBin = Int(coordinate / (4# * GEOM_TOLERANCE))
    End If
End Function

' Строит ключ пространственного индекса арматуры для поиска ближайших совпадений.
Private Function PositionKey(ByVal binX As Double, ByVal binY As Double) As String
    PositionKey = Format$(binX, "0") & "|" & Format$(binY, "0")
End Function

' Включает физический стержень в индекс координат без повторного добавления в раскладку.
Private Sub IndexPosition(ByVal index As Long)
    Dim key As String, bucket As Collection
    key = PositionKey(PositionBin(mX(index)), PositionBin(mY(index)))
    If mPositionBuckets.Exists(key) Then
        Set bucket = mPositionBuckets(key)
    Else
        Set bucket = New Collection
        mPositionBuckets.Add key, bucket
    End If
    bucket.Add index
End Sub

' Поиск в соседних ячейках покрывает точки по обе стороны границы bucket.
' При нескольких кандидатах владелец - первый опубликованный стержень.
Public Function FindBarAt(ByVal xCoord As Double, ByVal yCoord As Double) As Long
    If mCount = 0 Then Exit Function
    Dim bx As Double, by As Double, dx As Long, dy As Long, key As String
    Dim bucket As Collection, candidate As Variant, index As Long
    bx = PositionBin(xCoord): by = PositionBin(yCoord)
    For dx = -1 To 1
        For dy = -1 To 1
            key = PositionKey(bx + dx, by + dy)
            If mPositionBuckets.Exists(key) Then
                Set bucket = mPositionBuckets(key)
                For Each candidate In bucket
                    index = CLng(candidate)
                    If Abs(mX(index) - xCoord) <= GEOM_TOLERANCE And Abs(mY(index) - yCoord) <= GEOM_TOLERANCE Then
                        If FindBarAt = 0 Or index < FindBarAt Then FindBarAt = index
                    End If
                Next candidate
            End If
        Next dy
    Next dx
End Function

' Membership не создает второй физический стержень и не меняет его площадь.
Public Sub AddMembership(ByVal index As Long, ByVal groupName As String)
    CheckIndex index
    groupName = Trim$(groupName)
    If Len(groupName) = 0 Then Exit Sub
    EnsureIndexes
    Dim groups As Object, key As String
    key = CStr(index)
    If mMemberships.Exists(key) Then
        Set groups = mMemberships(key)
    Else
        Set groups = CreateObject("Scripting.Dictionary")
        groups.CompareMode = vbTextCompare
        mMemberships.Add key, groups
    End If
    If Not groups.Exists(groupName) Then groups.Add groupName, True
End Sub

' Все смысловые принадлежности физического стержня; они не создают дополнительные физические
' стержни.
Public Property Get BarMemberships(ByVal index As Long) As Variant
    CheckIndex index
    EnsureIndexes
    If mMemberships.Exists(CStr(index)) Then
        BarMemberships = mMemberships(CStr(index)).Keys
    Else
        BarMemberships = Array()
    End If
End Property

' Количество записей, сохранённых в контейнере этого класса.
Public Property Get Count() As Long
    Count = mCount
End Property

' Общий поиск уже созданного стержня для сопряжений граней и проекций.
' AddBar по-прежнему отклоняет дубли; решение пропустить позицию принимает builder.
Public Function HasBarAt(ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    HasBarAt = (FindBarAt(xCoord, yCoord) > 0)
End Function

' Промежуточный список стержней до CSectionModel CRebarLayout хранит координаты, диаметры и
' диагностические подписи стержней, сформированных раскладчиком. Затем CSectionModelBuilder
' переносит эти данные в единую расчетную модель и присваивает расчетные имена R1/R2/....
Public Property Get BarID(ByVal index As Long) As String
    CheckIndex index
    BarID = mBarID(index)
End Property

' Координата X выбранного элемента модели, мм.
Public Property Get X(ByVal index As Long) As Double
    CheckIndex index
    X = mX(index)
End Property

' Координата Y выбранного элемента модели, мм.
Public Property Get Y(ByVal index As Long) As Double
    CheckIndex index
    Y = mY(index)
End Property

' Диаметр выбранного физического стержня, мм.
Public Property Get Diameter(ByVal index As Long) As Double
    CheckIndex index
    Diameter = mDiameter(index)
End Property

' Площадь выбранного элемента, мм^2.
Public Property Get Area(ByVal index As Long) As Double
    CheckIndex index
    Area = mArea(index)
End Property

' Принадлежность выбранного стержня к материалу арматуры.
Public Property Get SteelClass(ByVal index As Long) As String
    CheckIndex index
    SteelClass = mSteelClass(index)
End Property

' Сохранённое пояснение выбранной записи без нового расчёта.
Public Property Get Comment(ByVal index As Long) As String
    CheckIndex index
    Comment = mComment(index)
End Property

' Смысловая группа владельца физического стержня, используемая для подписи количества и
' диаметра.
Public Property Get BarAnnotationGroupName(ByVal index As Long) As String
    CheckIndex index
    BarAnnotationGroupName = mBarAnnotationGroupName(index)
End Property

' Добавляет опорную линию группы арматуры, по которой потом строится групповая подпись на схеме.
Public Sub AddAnnotationAnchor(ByVal groupName As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal axisDistance As Double = 0#)
    If Len(Trim$(groupName)) = 0 Then Exit Sub

    mAnnotationCount = mAnnotationCount + 1
    ReDim Preserve mAnnotationGroupName(1 To mAnnotationCount)
    ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
    ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
    ReDim Preserve mAnnotationAxisDistance(1 To mAnnotationCount)

    mAnnotationGroupName(mAnnotationCount) = Trim$(groupName)
    mAnnotationStartX(mAnnotationCount) = startX
    mAnnotationStartY(mAnnotationCount) = startY
    mAnnotationEndX(mAnnotationCount) = endX
    mAnnotationEndY(mAnnotationCount) = endY
    mAnnotationNormalX(mAnnotationCount) = normalX
    mAnnotationNormalY(mAnnotationCount) = normalY
    mAnnotationAxisDistance(mAnnotationCount) = axisDistance
End Sub

' Количество привязок арматурных подписей в раскладке.
Public Property Get AnnotationCount() As Long
    AnnotationCount = mAnnotationCount
End Property

' Смысловая группа выбранной привязки арматурной подписи.
Public Property Get AnnotationGroupName(ByVal index As Long) As String
    CheckAnnotationIndex index
    AnnotationGroupName = mAnnotationGroupName(index)
End Property

' Координата X начала выбранной привязки арматурной подписи, мм.
Public Property Get AnnotationStartX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationStartX = mAnnotationStartX(index)
End Property

' Координата Y начала выбранной привязки арматурной подписи, мм.
Public Property Get AnnotationStartY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationStartY = mAnnotationStartY(index)
End Property

' Координата X конца выбранной привязки арматурной подписи, мм.
Public Property Get AnnotationEndX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationEndX = mAnnotationEndX(index)
End Property

' Координата Y конца выбранной привязки арматурной подписи, мм.
Public Property Get AnnotationEndY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationEndY = mAnnotationEndY(index)
End Property

' Компонента X направления вынесения выбранной арматурной подписи.
Public Property Get AnnotationNormalX(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationNormalX = mAnnotationNormalX(index)
End Property

' Компонента Y направления вынесения выбранной арматурной подписи.
Public Property Get AnnotationNormalY(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationNormalY = mAnnotationNormalY(index)
End Property

' Расстояние от исходной границы до оси ряда выбранной привязки, мм.
Public Property Get AnnotationAxisDistance(ByVal index As Long) As Double
    CheckAnnotationIndex index
    AnnotationAxisDistance = mAnnotationAxisDistance(index)
End Property

' Проверяет индекс стержня перед доступом к массивам раскладки.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3106, "CRebarLayout", "Индекс стержня вне допустимого диапазона."
End Sub

' Проверяет индекс аннотационной линии перед доступом к массивам раскладки.
Private Sub CheckAnnotationIndex(ByVal index As Long)
    If index < 1 Or index > mAnnotationCount Then Err.Raise vbObjectError + 3107, "CRebarLayout", "Индекс аннотации арматуры вне допустимого диапазона."
End Sub







