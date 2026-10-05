VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConcreteRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Готовая аналитическая область бетона и ее принадлежность результату
' ==========================================================================
' Контейнер хранит замкнутые ориентированные кольца из отрезков и дуг, площадь
' этих же колец и metadata владельца. Геометрические операции выполняет только
' CSectionGeometryQuery. Здесь нет НДС, нормативных коэффициентов, Excel или
' построения альтернативной области для экспорта. Публичное чтение не меняет
' снимок; геометрия заполняется общим движком до публикации результата.

Private mEdges() As TRegionEdge ' Граница в координатах модели, мм.
Private mCount As Long
Private mLoopCount As Long
Private mArea As Double ' Алгебраическая площадь outer минус opening, мм2.
Private mSource As String ' Достоверный контур либо восстановленная граница элементов.
Private mSectionRevision As Long ' Версия модели, к которой относится геометрия.
Private mOwnerID As String ' CombinationID конкретного результата, не ALL.
Private mStandard As String ' Нормативный владелец готовой области; не влияет на геометрию.
Private mAnchorID As String ' Критическая группа/стержень для проверки происхождения.
Private mLoopOpening() As Boolean ' Готовые роли колец, определенные общим topology-контрактом.
Private mBoundaryProbes As Variant ' [найден интервал, Qx, Qy, Px, Py]; диагностические точки, не аппроксимация площади.

' Принимает рассчитанный массив границ. Копирование отделяет готовый снимок
' от временных массивов topology/clipping, которые будут переиспользованы.
Friend Sub SetGeometry(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal loopCount As Long, ByVal area As Double, ByVal source As String, ByVal revision As Long)
    mCount = count: mLoopCount = loopCount: mArea = area
    mSource = source: mSectionRevision = revision
    Erase mEdges
    Erase mLoopOpening
    mBoundaryProbes = Empty
    If count = 0 Then Exit Sub
    ReDim mEdges(1 To count)
    Dim i As Long
    For i = 1 To count
        LSet mEdges(i) = edges(i)
    Next i
    If loopCount > 0 Then
        ReDim mLoopOpening(1 To loopCount)
        Dim query As CSectionGeometryQuery
        Set query = New CSectionGeometryQuery
        For i = 1 To loopCount: mLoopOpening(i) = query.LoopIsOpening(Me, i): Next i
    End If
End Sub

' Сохраняет готовые диагностические пары внутренних/внешних границ.
' Число точек задает потребитель; площадь и кольца по этим пробам не строятся.
Friend Sub SetBoundaryProbes(ByVal probes As Variant)
    mBoundaryProbes = probes
End Sub

' Возвращает независимый рабочий массив для следующей геометрической операции.
' Только общий движок изменяет этот массив, не затрагивая опубликованный снимок.
Friend Sub CopyEdges(ByRef edges() As TRegionEdge, ByRef count As Long)
    count = mCount
    Erase edges
    If count = 0 Then Exit Sub
    ReDim edges(1 To count)
    Dim i As Long
    For i = 1 To count
        LSet edges(i) = mEdges(i)
    Next i
End Sub

' Сохраняет принадлежность области инженерному результату без ее пересчета.
' Передается coordinator/calculator-ом; writer не назначает критическую группу.
Public Sub SetOwner(ByVal combinationID As String, ByVal standard As String, ByVal anchorID As String)
    mOwnerID = combinationID: mStandard = standard: mAnchorID = anchorID
End Sub

' Читает один аналитический сегмент для annotations/export. У дуги координаты
' центра и signed sweep сохраняются без аппроксимации промежуточными точками.
Public Sub GetSegment(ByVal index As Long, ByRef x1 As Double, ByRef y1 As Double, _
        ByRef x2 As Double, ByRef y2 As Double, ByRef centerX As Double, _
        ByRef centerY As Double, ByRef sweep As Double, ByRef loopID As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 5400, "CConcreteRegion", "Номер сегмента области вне допустимого диапазона."
    x1 = mEdges(index).X1: y1 = mEdges(index).Y1
    x2 = mEdges(index).X2: y2 = mEdges(index).Y2
    centerX = mEdges(index).CenterX: centerY = mEdges(index).CenterY
    sweep = mEdges(index).Sweep: loopID = mEdges(index).LoopID
End Sub

' Простые свойства возвращают неизменяемые значения готового снимка.
Public Property Get SegmentCount() As Long
    SegmentCount = mCount
End Property

' Возвращает копию сохраненных пар P/Q для пассивной сериализации результата.
' Empty означает, что такие пробы не требовались, например при центральном растяжении.
Public Property Get BoundaryProbes() As Variant
    BoundaryProbes = mBoundaryProbes
End Property
Public Property Get BoundaryProbeCount() As Long
    If IsArray(mBoundaryProbes) Then BoundaryProbeCount = UBound(mBoundaryProbes, 1)
End Property
Public Property Get LoopCount() As Long
    LoopCount = mLoopCount
End Property

' Возвращает роль сохраненного кольца через общий topology-контракт.
' Потребителям аннотаций не требуется собственный расчет ориентации границ.
Public Property Get LoopIsOpening(ByVal index As Long) As Boolean
    If index < 1 Or index > mLoopCount Then Err.Raise vbObjectError + 5400, "CConcreteRegion", "Номер кольца области вне допустимого диапазона."
    LoopIsOpening = mLoopOpening(index)
End Property
Public Property Get Area() As Double
    Area = mArea
End Property
Public Property Get Source() As String
    Source = mSource
End Property
Public Property Get SectionRevision() As Long
    SectionRevision = mSectionRevision
End Property
Public Property Get OwnerID() As String
    OwnerID = mOwnerID
End Property
Public Property Get Standard() As String
    Standard = mStandard
End Property
Public Property Get AnchorID() As String
    AnchorID = mAnchorID
End Property

' Возвращает независимый снимок той же области, сохраняя владельца и источник.
' Это не восстановление по площади: копируются все кольца и аналитические дуги.
Public Function Clone() As CConcreteRegion
    Dim result As CConcreteRegion
    Set result = New CConcreteRegion
    result.SetGeometry mEdges, mCount, mLoopCount, mArea, mSource, mSectionRevision
    result.SetOwner mOwnerID, mStandard, mAnchorID
    result.SetBoundaryProbes mBoundaryProbes
    Set Clone = result
End Function
