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
' Список стержней перед добавлением в расчётную модель
' ==========================================================================
' CRebarLayout хранит координаты, диаметры и диагностические подписи стержней,
' сформированных раскладчиком. Затем CSectionModelBuilder переносит эти данные
' в единую расчётную модель; расчётные имена R1/R2/... назначает сама CSectionModel.

Private mBarID() As String ' Имена источника физических стержней; расчётные R1/R2 назначает CSectionModel.
Private mX() As Double ' Координаты центров стержней X, мм.
Private mY() As Double ' Координаты центров стержней Y, мм.
Private mDiameter() As Double ' Фактические диаметры стержней, мм.
Private mArea() As Double ' Фактические площади стержней, мм2.
Private mSteelClass() As String ' Материал стержня; сравнивается при объединении совпавших позиций.
Private mComment() As String ' Пояснение происхождения физического стержня.
Private mBarAnnotationGroupName() As String ' Группа, в подписи которой учитываются количество и диаметр этого стержня.
Private mReadOnly As Boolean ' После завершения сборки координаты стержней и состав групп изменять нельзя.
Private mCount As Long ' Количество физических стержней после обработки совпадений.
Private mCapacity As Long ' Число мест в массивах; когда свободные места заканчиваются, размер увеличивается вдвое.
Private mIDs As Object ' Номера стержней по их уникальным именам для быстрого поиска.
Private mPositionBuckets As Object ' Списки стержней с близкими координатами; окончательно сравниваются исходные координаты.
Private mMemberships As Object ' Группы, с которыми связан стержень; эти связи не добавляют арматуру и не увеличивают её подпись.


' Очищает список стержней и связанные с ними группы перед новой расстановкой.
' После завершения сборки очистка запрещена. Аннотации хранятся отдельно.
Public Sub Clear()
    RequireMutable
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

End Sub

' Добавляет стержень с уникальными ID и координатами; проверяет диаметр и площадь.
' Если передана geometry, дополнительно проверяет принадлежность центра бетону;
' проверку полного диска стержня при необходимости выполняет CRebarRuleExecutor.
Public Sub AddBar(ByVal barName As String, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal comment As String = vbNullString, Optional ByVal geometry As CGeometryRegion = Nothing, _
        Optional ByVal annotationGroupName As String = vbNullString)
    RequireMutable

    If Len(Trim$(barName)) = 0 Then Err.Raise vbObjectError + 3100, "CRebarLayout", "ID стержня не заполнен."
    If diameter <= 0# Then Err.Raise vbObjectError + 3101, "CRebarLayout", "Диаметр арматуры должен быть положительным."
    If CDbl(mCount) >= GEOM_MAX_REBAR_COUNT Then Err.Raise vbObjectError + 3108, "CRebarLayout", _
        "Количество стержней превышает диапазон счетчиков программы. Уменьшите активную раскладку в Config или число импортируемых элементов." ' Ограничиваем фактическое число стержней при любом способе создания, включая импорт и проекцию на контур отверстия. Построители заранее проверяют и запрошенное количество.

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

' Создаёт таблицы для быстрого поиска стержня по имени, близким координатам и связанным группам.
Private Sub EnsureIndexes()
    If Not mIDs Is Nothing Then Exit Sub
    Set mIDs = CreateObject("Scripting.Dictionary")
    mIDs.CompareMode = vbTextCompare
    Set mPositionBuckets = CreateObject("Scripting.Dictionary")
    Set mMemberships = CreateObject("Scripting.Dictionary")
End Sub

' При нехватке места увеличивает массивы, сохраняя все уже добавленные стержни.
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

' Вычисляет номер условной ячейки для координаты, чтобы быстрее искать близкие стержни.
' Попадание в одну ячейку ещё не означает совпадение: исходные координаты затем
' сравниваются с допуском GEOM_TOLERANCE без округления.
Private Function PositionBin(ByVal coordinate As Double) As Double
    If Abs(coordinate) > 1E+290 Then
        PositionBin = coordinate
    Else
        PositionBin = Int(coordinate / (4# * GEOM_TOLERANCE))
    End If
End Function

' Соединяет номера ячеек по X и Y в текстовый ключ для быстрого поиска близких стержней.
Private Function PositionKey(ByVal binX As Double, ByVal binY As Double) As String
    PositionKey = Format$(binX, "0") & "|" & Format$(binY, "0")
End Function

' Добавляет номер стержня в список своей координатной ячейки, не создавая новый стержень.
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

' Ищет стержень с заданными координатами в текущей и соседних ячейках поиска.
' Проверка соседей учитывает близкие точки по разные стороны границы ячейки.
' При нескольких совпадениях возвращает номер стержня, добавленного раньше остальных.
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

' Связывает существующий стержень с ещё одной группой.
' Новый стержень не создаётся, его площадь и количество арматуры не меняются.
Public Sub AddMembership(ByVal index As Long, ByVal groupName As String)
    RequireMutable
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

' Возвращает названия всех групп, с которыми связан выбранный стержень.
' Несколько связей означают один стержень, а не несколько копий арматуры.
Public Property Get BarMemberships(ByVal index As Long) As Variant
    CheckIndex index
    EnsureIndexes
    If mMemberships.Exists(CStr(index)) Then
        BarMemberships = mMemberships(CStr(index)).Keys
    Else
        BarMemberships = Array()
    End If
End Property

' Возвращает количество сохранённых записей.
Public Property Get Count() As Long
    Count = mCount
End Property

' Проверяет, есть ли стержень в заданной точке, например на стыке сторон или после проекции.
' Сам AddBar отклоняет дубликат; пропуск повторной позиции определяет общий построитель.
Public Function HasBarAt(ByVal xCoord As Double, ByVal yCoord As Double) As Boolean
    HasBarAt = (FindBarAt(xCoord, yCoord) > 0)
End Function

' Возвращает исходное имя выбранного стержня. Расчётное имя R1/R2/... назначается позднее в
' CSectionModel.
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

' Возвращает группу, в подписи которой учитываются количество и диаметр выбранного стержня.
Public Property Get BarAnnotationGroupName(ByVal index As Long) As String
    CheckIndex index
    BarAnnotationGroupName = mBarAnnotationGroupName(index)
End Property

' Проверяет индекс стержня перед доступом к массивам раскладки.
Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3106, "CRebarLayout", "Индекс стержня вне допустимого диапазона." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mCount) & "."
End Sub

' Завершает заполнение списка стержней и запрещает менять его.
' Массивы сохраняются для чтения без повторного создания и копирования.
Friend Sub Freeze()
    mReadOnly = True
End Sub

' Показывает, запрещено ли дальнейшее изменение координат стержней и состава групп.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mReadOnly
End Property

' Проверяет, можно ли ещё менять список стержней.
' После завершения сборки изменение запрещено; для новой сборки нужен новый объект.
Private Sub RequireMutable()
    If mReadOnly Then Err.Raise vbObjectError + 3110, "CRebarLayout", "Опубликованную арматурную раскладку нельзя изменять. Создайте новую раскладку."
End Sub
