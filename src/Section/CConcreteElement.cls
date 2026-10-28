VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConcreteElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемые данные одного бетонного элемента
' ==========================================================================
' Хранит расчётные характеристики, источник и фактическую границу одной сущности.
' Численные свойства подготавливает общий геометрический слой, а согласованность
' проверяет CSectionModel. Здесь нет Excel, решения НДС и вычисления инерций.

Private mInitialized As Boolean ' Повторная запись уже опубликованной сущности запрещена.
Private mElementID As String ' Единственное расчётное имя Cn, назначенное CSectionModel.
Private mX As Double ' Координата центра тяжести по X, мм.
Private mY As Double ' Координата центра тяжести по Y, мм.
Private mArea As Double ' Фактическая расчётная площадь элемента, мм2.
Private mMaterialID As Long ' Идентификатор подготовленного бетонного материала.
Private mSourceName As String ' Исходное имя элемента из генератора, файла или чертежа.
Private mSourceHandle As String ' Handle исходного CAD-объекта без привязки к живому COM.
Private mShapeType As String ' Contour требует фактическую границу; компактные типы используют подготовленное прямоугольное описание.
Private mWidth As Double ' Настоящая ширина прямоугольника либо только габарит произвольной формы, мм.
Private mHeight As Double ' Настоящая высота прямоугольника либо только габарит произвольной формы, мм.
Private mRotation As Double ' Подготовленный угол прямоугольной геометрии в радианах.
Private mRotationKnown As Boolean ' Отличает заданное направление от неопределённой ориентации почти изотропной области.
Private mKnownRectangle As Boolean ' Прямоугольная форма была явно задана исходником до интерпретации Width/Height; эквивалентная оболочка не меняет этот факт.
Private mComment As String ' Пояснение происхождения и особенностей исходного элемента.
Private mLocalIx As Double ' Собственный момент вокруг X центра элемента, мм4.
Private mLocalIy As Double ' Собственный момент вокруг Y центра элемента, мм4.
Private mLocalIxy As Double ' Подписанный интеграл x*y относительно центра элемента, мм4.
Private mExactGeometry As CConcreteRegion ' Независимая фактическая область; Nothing допустим только для компактной записи.
Private mLegacyGeometry As CConcreteRegion ' Лениво подготовленный прямоугольник компактного элемента; не заменяет геометрию Contour.

' Один раз принимает уже проверенные моделью данные и независимо сохраняет
' точный контур. Размеры и угол компактных элементов уже интерпретированы моделью.
Friend Sub Initialize(ByVal elementID As String, ByVal x As Double, ByVal y As Double, ByVal area As Double, _
        ByVal materialID As Long, ByVal sourceName As String, ByVal sourceHandle As String, ByVal shapeType As String, _
        ByVal width As Double, ByVal height As Double, ByVal rotation As Double, ByVal rotationKnown As Boolean, _
        ByVal comment As String, ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        Optional ByVal exactRegion As CConcreteRegion = Nothing, Optional ByVal knownRectangle As Boolean = False)
    If mInitialized Then Err.Raise vbObjectError + 5224, "CConcreteElement", "Данные опубликованного бетонного элемента нельзя переинициализировать."
    If StrComp(Trim$(shapeType), "Contour", vbTextCompare) = 0 And exactRegion Is Nothing Then _
        Err.Raise vbObjectError + 5223, "CConcreteElement", "Фактическая граница бетонного элемента Contour не передана."
    Dim geometryCopy As CConcreteRegion
    If Not exactRegion Is Nothing Then Set geometryCopy = exactRegion.Clone
    mElementID = elementID: mX = x: mY = y: mArea = area: mMaterialID = materialID
    mSourceName = sourceName: mSourceHandle = sourceHandle: mShapeType = shapeType
    mWidth = width: mHeight = height: mRotation = rotation: mRotationKnown = rotationKnown
    mKnownRectangle = knownRectangle
    mComment = comment: mLocalIx = localIx: mLocalIy = localIy: mLocalIxy = localIxy
    Set mExactGeometry = geometryCopy
    mInitialized = True
End Sub

' Возвращает расчётное имя, назначенное единственным владельцем списка элементов.
Public Property Get ElementID() As String
    ElementID = mElementID
End Property

' Возвращает координату центра элемента по X, мм.
Public Property Get X() As Double
    X = mX
End Property

' Возвращает координату центра элемента по Y, мм.
Public Property Get Y() As Double
    Y = mY
End Property

' Возвращает фактическую расчётную площадь, мм2.
Public Property Get Area() As Double
    Area = mArea
End Property

' Возвращает идентификатор материала, не создавая модель диаграммы.
Public Property Get MaterialID() As Long
    MaterialID = mMaterialID
End Property

' Возвращает исходное имя из генератора, файла или CAD.
Public Property Get SourceName() As String
    SourceName = mSourceName
End Property

' Возвращает исходный CAD-handle как строку, без обращения к чертежу.
Public Property Get SourceHandle() As String
    SourceHandle = mSourceHandle
End Property

' Возвращает сохранённый геометрический контракт, включая обязательный Contour.
Public Property Get ShapeType() As String
    ShapeType = mShapeType
End Property

' Возвращает подготовленную ширину; у произвольной формы она не задаёт её границу.
Public Property Get Width() As Double
    Width = mWidth
End Property

' Возвращает подготовленную высоту; у произвольной формы она не задаёт её границу.
Public Property Get Height() As Double
    Height = mHeight
End Property

' Возвращает подготовленный угол прямоугольника, рад.
Public Property Get Rotation() As Double
    Rotation = mRotation
End Property

' Сообщает, было ли направление элемента достоверно определено источником.
Public Property Get RotationKnown() As Boolean
    RotationKnown = mRotationKnown
End Property

' Отличает достоверный прямоугольник источника от интерпретированной оболочки
' той же площади/инерции; подготовленные размеры сами по себе этого не доказывают.
Public Property Get HasKnownRectangle() As Boolean
    HasKnownRectangle = mKnownRectangle
End Property

' Возвращает исходное пояснение без форматирования диалога или ячейки Excel.
Public Property Get Comment() As String
    Comment = mComment
End Property

' Возвращает подготовленный собственный Ix, мм4; геометрия здесь не интегрируется.
Public Property Get LocalIx() As Double
    LocalIx = mLocalIx
End Property

' Возвращает подготовленный собственный Iy, мм4; размеры не подменяют инерцию.
Public Property Get LocalIy() As Double
    LocalIy = mLocalIy
End Property

' Возвращает подписанный собственный интеграл x*y, мм4.
Public Property Get LocalIxy() As Double
    LocalIxy = mLocalIxy
End Property

' Отличает фактическую сохранённую область от ленивой компактной оболочки.
Public Property Get HasExactGeometry() As Boolean
    HasExactGeometry = Not mExactGeometry Is Nothing
End Property

' Выдаёт независимую копию геометрии, чтобы изменение её назначения или
' границ не меняло опубликованную сущность. Только компактный элемент без
' обязательного Contour получает ленивый прямоугольник через общее ядро.
Public Property Get Geometry() As CConcreteRegion
    If Not mInitialized Then Err.Raise vbObjectError + 5224, "CConcreteElement", "Данные бетонного элемента ещё не подготовлены."
    If Not mExactGeometry Is Nothing Then
        Set Geometry = mExactGeometry.Clone
        Exit Property
    End If
    If StrComp(mShapeType, "Contour", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 5223, "CConcreteElement", "Потеряна обязательная фактическая граница элемента " & mElementID & "."
    If mLegacyGeometry Is Nothing Then
        Dim query As CSectionGeometryQuery
        Set query = New CSectionGeometryQuery
        Set mLegacyGeometry = query.RectangleRegion(mX, mY, mWidth, mHeight, mRotation)
    End If
    Set Geometry = mLegacyGeometry.Clone
End Property

' Создаёт заменяющую сущность с той же расчётной идентичностью и A/c/I.
' Согласованность новой области предварительно проверяет CSectionModel.
Friend Function WithGeometry(ByVal region As CConcreteRegion, ByVal width As Double, _
        ByVal height As Double, ByVal rotation As Double) As CConcreteElement
    If Not mInitialized Then Err.Raise vbObjectError + 5224, "CConcreteElement", "Данные бетонного элемента ещё не подготовлены."
    If region Is Nothing Then Err.Raise vbObjectError + 5223, "CConcreteElement", "Фактическая граница бетонного элемента не передана."
    Dim result As CConcreteElement
    Set result = New CConcreteElement
    result.Initialize mElementID, mX, mY, mArea, mMaterialID, mSourceName, mSourceHandle, "Contour", _
        width, height, rotation, True, mComment, mLocalIx, mLocalIy, mLocalIxy, region, False
    Set WithGeometry = result
End Function

' Создаёт компактную сущность с уточнённой ориентацией её оболочки.
' Фактическую область этим методом поворачивать или заменять запрещено.
Friend Function WithRotation(ByVal rotation As Double) As CConcreteElement
    If Not mInitialized Then Err.Raise vbObjectError + 5224, "CConcreteElement", "Данные бетонного элемента ещё не подготовлены."
    If Not mExactGeometry Is Nothing Or StrComp(mShapeType, "Contour", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 5224, "CConcreteElement", "Нельзя менять фактическую границу элемента через ориентацию эквивалентной оболочки."
    Dim result As CConcreteElement
    Set result = New CConcreteElement
    result.Initialize mElementID, mX, mY, mArea, mMaterialID, mSourceName, mSourceHandle, mShapeType, _
        mWidth, mHeight, rotation, mRotationKnown, mComment, mLocalIx, mLocalIy, mLocalIxy, Nothing, mKnownRectangle
    Set WithRotation = result
End Function


' Подпись уже проверенной геометрической записи принадлежит тому же объекту, что и фактическая граница.
Public Property Get GeometryHash() As String
    If Not mExactGeometry Is Nothing Then GeometryHash = mExactGeometry.GeometryHash
End Property

' Добавляет только производную подпись сериализации; характеристики и границы элемента не меняются.
Friend Sub RememberGeometryHash(ByVal value As String)
    If Not mExactGeometry Is Nothing Then mExactGeometry.RememberGeometryHash value
End Sub
