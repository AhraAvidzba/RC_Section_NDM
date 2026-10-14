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

' Исходные вершины многоугольника до скругления: координаты, радиусы, имена и источники.
' Не строит касания, пути или области. После явного Seal исходные данные неизменяемы.
Private Type TVertex
    ID As String
    X As Double
    Y As Double
    Radius As Double
    Source As String
End Type
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type
Private mVertices() As TVertex ' Единственный массив полных записей в исходном порядке.
Private mCount As Long ' Число добавленных вершин; первую вершину повторно в конце не добавляют.
Private mCapacity As Long ' Внутренняя ёмкость массива, не ограниченная конкретным типом сечения.
Private mSealed As Boolean ' Завершение структурного ввода; сложную топологию проверяет геометрическое ядро.
Private mIDs As Object ' Поиск повторных ID с прежним различением регистра.

' Задаёт один источник вершины, радиуса и координат; геометрические операции здесь не выполняются.
Public Sub AddVertex(ByVal vertexID As String, ByVal x As Double, ByVal y As Double, _
        Optional ByVal filletRadius As Double = 0#, Optional ByVal sourceReference As String = "")
    If mSealed Then SpecError "Описание вершин уже завершено; создайте новое для других размеров.", sourceReference
    If Len(Trim$(vertexID)) = 0 Then SpecError "Не задано имя исходной вершины.", sourceReference
    If mIDs Is Nothing Then
        Set mIDs = CreateObject("Scripting.Dictionary"): mIDs.CompareMode = vbBinaryCompare
    End If
    If mIDs.Exists(vertexID) Then SpecError "Повторное имя исходной вершины '" & vertexID & "'.", sourceReference
    If Not Finite(x) Or Not Finite(y) Or Not Finite(filletRadius) Then _
        SpecError "Координаты и радиус вершины '" & vertexID & "' должны быть конечными числами.", sourceReference
    If filletRadius < 0# Then SpecError "Радиус вершины '" & vertexID & "' не может быть отрицательным.", sourceReference
    If mCount = mCapacity Then
        If mCapacity > 1073741823 Then SpecError "Слишком много исходных вершин.", sourceReference
        If mCapacity = 0 Then mCapacity = 4 Else mCapacity = mCapacity * 2
        ReDim Preserve mVertices(1 To mCapacity)
    End If
    With mVertices(mCount + 1)
        .ID = vertexID: .X = x: .Y = y: .Radius = filletRadius: .Source = sourceReference
    End With
    mIDs.Add vertexID, mCount + 1
    mCount = mCount + 1
End Sub

' Закрывает полный набор исходных вершин без сортировки или вычисления скруглений.
' Повторный Seal безопасен; неуспешный Seal не закрывает неполное описание.
Public Sub Seal()
    If mSealed Then Exit Sub
    If mCount < 3 Then SpecError "Для многоугольника нужны не менее трёх исходных вершин."
    ReDim Preserve mVertices(1 To mCount): mCapacity = mCount
    mSealed = True
End Sub

' Число исходных вершин, доступное также во время заполнения.
Public Property Get Count() As Long: Count = mCount: End Property
' Состояние описателя; операции фасада принимают только закрытый набор.
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property

' Смысловое имя исходной вершины по её порядковому номеру, 1..Count.
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

' Единственный внутренний мост к прежнему численному представлению; массивы не являются API рецепта.
Friend Sub CopyKernelData(ByRef vertices As Variant, ByRef radii As Variant, ByRef vertexIDs As Variant)
    RequireSealed
    Dim points() As Variant, assigned() As Double, names() As String, i As Long
    ReDim points(0 To mCount - 1): ReDim assigned(0 To mCount - 1): ReDim names(0 To mCount - 1)
    For i = 1 To mCount
        points(i - 1) = Array(mVertices(i).X, mVertices(i).Y)
        assigned(i - 1) = mVertices(i).Radius: names(i - 1) = mVertices(i).ID
    Next i
    vertices = points: radii = assigned: vertexIDs = names
End Sub

' Проверяет индекс исходного описания; чтение геометрических данных возможно после Seal.
Private Sub CheckIndex(ByVal index As Long)
    RequireSealed
    If index < 1 Or index > mCount Then SpecError "Номер исходной вершины вне диапазона 1..Count."
End Sub
Private Sub RequireSealed()
    If Not mSealed Then SpecError "Перед использованием завершите описание вершин методом Seal."
End Sub

' Экспонента IEEE754 проверяет NaN/Infinity без арифметики и без дополнительного геометрического допуска.
Private Function Finite(ByVal value As Double) As Boolean
    Dim number As TDoubleValue, words As TDoubleWords
    number.Value = value: LSet words = number
    Finite = ((words.High And &H7FF00000) <> &H7FF00000)
End Function
' Структурная ошибка сохраняет адрес именно добавляемого источника.
Private Sub SpecError(ByVal reason As String, Optional ByVal sourceReference As String = "")
    If Len(sourceReference) > 0 Then reason = reason & " Источник: " & sourceReference
    Err.Raise vbObjectError + 5470, "CPolygonSpec", reason
End Sub
