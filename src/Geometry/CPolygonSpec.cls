VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPolygonSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Исходное описание многоугольника со скруглениями
' ==========================================================================
' Хранит координаты вершин, радиусы углов, имена и сведения об источниках.
' Касания, скруглённые контуры и бетонную область строят общие геометрические классы.
' После Seal исходное описание доступно только для чтения.

Private Type TVertex
    ID As String ' Имя исходной вершины с различением регистра.
    X As Double ' Исходная координата X до скругления, мм.
    Y As Double ' Исходная координата Y до скругления, мм.
    Radius As Double ' Радиус сопряжения данной вершины, мм; ноль оставляет острый угол.
    Source As String ' Адрес или текстовое происхождение данных без ссылки на объект Excel.
    OutgoingEdgeID As String ' Имя прямого участка от этой вершины к следующей между касаниями скруглений.
    CornerID As String ' Имя скругления этой вершины; сохраняется как привязка и при нулевом радиусе.
End Type
Private Type TDoubleValue
    Value As Double ' Число Double для чтения его точного двоичного представления.
End Type
Private Type TDoubleWords
    Low As Long ' Младшие 32 бита числа Double.
    High As Long ' Старшие 32 бита числа Double.
End Type
Private mVertices() As TVertex ' Единственный массив полных записей в исходном порядке.
Private mCount As Long ' Число добавленных вершин; первую вершину повторно в конце не добавляют.
Private mCapacity As Long ' Внутренняя ёмкость массива, не ограниченная конкретным типом сечения.
Private mSealed As Boolean ' Признак завершения ввода вершин; замкнутость, пересечения и вложенность проверяет общее геометрическое ядро.
Private mIDs As Object ' Поиск повторных ID с различением регистра.

' Добавляет вершину с её именем, координатами, радиусом скругления и ссылкой на исходные данные.
' Сам контур на этом этапе не строится.
' Пустые имена участков дают Edge.<vertexID> и Corner.<vertexID>; имя вершины остаётся самостоятельным.
Public Sub AddVertex(ByVal vertexID As String, ByVal x As Double, ByVal y As Double, _
        Optional ByVal filletRadius As Double = 0#, Optional ByVal sourceReference As String = "", _
        Optional ByVal outgoingEdgeID As String = "", Optional ByVal cornerID As String = "")
    If mSealed Then SpecError "Описание вершин уже завершено; создайте новое для других размеров.", sourceReference
    If Len(Trim$(vertexID)) = 0 Then SpecError "Не задано имя исходной вершины.", sourceReference
    If mIDs Is Nothing Then
        Set mIDs = CreateObject("Scripting.Dictionary"): mIDs.CompareMode = vbBinaryCompare
    End If
    If mIDs.Exists(vertexID) Then SpecError "Повторное имя исходной вершины '" & vertexID & "'.", sourceReference
    If Not Finite(x) Or Not Finite(y) Or Not Finite(filletRadius) Then _
        SpecError "Координаты и радиус вершины '" & vertexID & "' должны быть конечными числами.", sourceReference
    If filletRadius < 0# Then SpecError "Радиус вершины '" & vertexID & "' не может быть отрицательным.", sourceReference
    If Len(Trim$(outgoingEdgeID)) = 0 Then outgoingEdgeID = "Edge." & vertexID
    If Len(Trim$(cornerID)) = 0 Then cornerID = "Corner." & vertexID
    If outgoingEdgeID = cornerID Then SpecError "Прямой участок и скругление не могут иметь общее имя '" & cornerID & "'.", sourceReference
    Dim i As Long
    For i = 1 To mCount
        If outgoingEdgeID = mVertices(i).OutgoingEdgeID Or outgoingEdgeID = mVertices(i).CornerID Then _
            SpecError "Повторное имя участка '" & outgoingEdgeID & "'.", sourceReference
        If cornerID = mVertices(i).OutgoingEdgeID Or cornerID = mVertices(i).CornerID Then _
            SpecError "Повторное имя участка '" & cornerID & "'.", sourceReference
    Next i
    If mCount = mCapacity Then
        If mCapacity > 1073741823 Then SpecError "Слишком много исходных вершин.", sourceReference
        If mCapacity = 0 Then mCapacity = 4 Else mCapacity = mCapacity * 2
        ReDim Preserve mVertices(1 To mCapacity)
    End If
    With mVertices(mCount + 1)
        .ID = vertexID: .X = x: .Y = y: .Radius = filletRadius: .Source = sourceReference
        .OutgoingEdgeID = outgoingEdgeID: .CornerID = cornerID
    End With
    mIDs.Add vertexID, mCount + 1
    mCount = mCount + 1
End Sub

' Завершает ввод исходных вершин без их сортировки и без построения скруглений.
' Повторный вызов безопасен; ошибка не закрывает неполное описание от исправления.
Public Sub Seal()
    If mSealed Then Exit Sub
    If mCount < 3 Then SpecError "Для многоугольника нужны не менее трёх исходных вершин."
    ReDim Preserve mVertices(1 To mCount): mCapacity = mCount
    mSealed = True
End Sub

' Число исходных вершин, доступное также во время заполнения.
Public Property Get Count() As Long: Count = mCount: End Property
' Показывает, завершён ли ввод вершин. Команды построения принимают только завершённое описание.
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property

' Возвращает имя исходной вершины по её порядковому номеру от 1 до Count.
Public Function VertexID(ByVal index As Long) As String
    CheckIndex index: VertexID = mVertices(index).ID
End Function
' Координата исходной вершины X, мм; она может не лежать на готовом скруглённом контуре.
Public Function X(ByVal index As Long) As Double
    CheckIndex index: X = mVertices(index).X
End Function
' Координата исходной вершины Y, мм.
Public Function Y(ByVal index As Long) As Double
    CheckIndex index: Y = mVertices(index).Y
End Function
' Назначенный радиус сопряжения, мм; ноль сохраняет острый угол.
Public Function FilletRadius(ByVal index As Long) As Double
    CheckIndex index: FilletRadius = mVertices(index).Radius
End Function
' Текстовое происхождение исходных данных, без сохранения Range/Workbook.
Public Function SourceReference(ByVal index As Long) As String
    CheckIndex index: SourceReference = mVertices(index).Source
End Function

' Имя прямого участка к следующей вершине; действует также для привязки исчезнувшего участка.
Public Function OutgoingEdgeID(ByVal index As Long) As String
    CheckIndex index: OutgoingEdgeID = mVertices(index).OutgoingEdgeID
End Function
' Имя скругления вершины; при нулевом радиусе обозначает только место его привязки.
Public Function CornerID(ByVal index As Long) As String
    CheckIndex index: CornerID = mVertices(index).CornerID
End Function

' Копирует координаты и радиусы во внутренние массивы общего геометрического ядра.
' Рецепт сечения задаёт вершины через объект и не работает с этими массивами.
Friend Sub CopyKernelData(ByRef vertices As Variant, ByRef radii As Variant, ByRef vertexIDs As Variant, _
        Optional ByRef outgoingEdgeIDs As Variant, Optional ByRef cornerIDs As Variant)
    RequireSealed
    Dim points() As Variant, assigned() As Double, names() As String, i As Long
    Dim edges() As String, corners() As String
    ReDim points(0 To mCount - 1): ReDim assigned(0 To mCount - 1): ReDim names(0 To mCount - 1)
    ReDim edges(0 To mCount - 1): ReDim corners(0 To mCount - 1)
    For i = 1 To mCount
        points(i - 1) = Array(mVertices(i).X, mVertices(i).Y)
        assigned(i - 1) = mVertices(i).Radius: names(i - 1) = mVertices(i).ID
        edges(i - 1) = mVertices(i).OutgoingEdgeID: corners(i - 1) = mVertices(i).CornerID
    Next i
    vertices = points: radii = assigned: vertexIDs = names
    If Not IsMissing(outgoingEdgeIDs) Then outgoingEdgeIDs = edges
    If Not IsMissing(cornerIDs) Then cornerIDs = corners
End Sub

' Проверяет индекс исходного описания; чтение геометрических данных возможно после Seal.
Private Sub CheckIndex(ByVal index As Long)
    RequireSealed
    If index < 1 Or index > mCount Then SpecError "Номер исходной вершины вне диапазона 1..Count." & " Запрошен номер " & CStr(index) & "; доступно элементов: " & CStr(mCount) & "."
End Sub
' Отклоняет чтение координат до завершения и проверки описания многоугольника.
Private Sub RequireSealed()
    If Not mSealed Then SpecError "Перед использованием завершите описание вершин методом Seal."
End Sub

' Проверяет по битовому представлению IEEE754, что число конечно.
' NaN и бесконечность отклоняются без арифметических операций и геометрических допусков.
Private Function Finite(ByVal value As Double) As Boolean
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number ' LSet переносит те же 64 бита Double, не выполняя числового преобразования.
    Finite = ((words.High And &H7FF00000) <> &H7FF00000) ' Все единицы в поле экспоненты IEEE754 означают NaN либо бесконечность.
End Function
' Сообщает об ошибке ввода вершины с адресом именно того параметра, который добавлялся.
Private Sub SpecError(ByVal reason As String, Optional ByVal sourceReference As String = "")
    If Len(sourceReference) > 0 Then reason = reason & " Источник: " & sourceReference
    Err.Raise vbObjectError + 5470, "CPolygonSpec", reason
End Sub
