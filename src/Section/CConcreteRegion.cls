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
' Бетонная область с точными границами и площадью
' ==========================================================================
' Хранит наружные контуры и отверстия из отрезков и дуг и площадь области.
' CGeometryRegion использует её при построении бетонной сетки, а расчёт трещин —
' для сохранения областей взаимодействия бетона с арматурой.
' Границы и площадь подготавливает общий CSectionGeometryQuery.

Private mEdges() As TRegionEdge ' Граница в координатах модели, мм.
Private mCount As Long ' Число отрезков и дуг в границах бетонной области.
Private mLoopCount As Long ' Число замкнутых контуров, включая наружные границы и отверстия.
Private mArea As Double ' Площадь внутри наружных контуров за вычетом площадей отверстий, мм2.
Private mSource As String ' Достоверный контур либо восстановленная граница элементов.
Private mSectionRevision As Long ' Версия модели, к которой относится геометрия.
Private mOwnerID As String ' Номер сочетания нагрузок, к результату которого относится область; значение ALL здесь не используется.
Private mStandard As String ' Норматив расчёта, для которого подготовлена область; её геометрия от этого поля не меняется.
Private mAnchorID As String ' Группа арматуры или стержень, для которых получена эта область.
Private mLoopOpening() As Boolean ' Роли замкнутых контуров: наружная граница или граница отверстия.
Private mBoundaryProbes As Variant ' Пары точек на границах бетона вдоль контрольных прямых для проверки области взаимодействия; площадь рассчитывается по точным контурам.

' Копирует готовые границы области, чтобы дальнейшее использование рабочих массивов
' при пересечениях и обрезке не изменило сохранённый результат.
Friend Sub SetGeometry(ByRef edges() As TRegionEdge, ByVal count As Long, _
        ByVal loopCount As Long, ByVal area As Double, ByVal source As String, ByVal revision As Long, _
        ByRef loopOpening() As Boolean)
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
        For i = 1 To loopCount: mLoopOpening(i) = loopOpening(i): Next i
    End If
End Sub

' Сохраняет точки начала Q и конца P участка бетона на каждой контрольной прямой
' вместе с признаком, найден ли этот участок. Эти данные нужны для проверки
' границ области взаимодействия; её контуры и площадь вычисляются отдельно.
Friend Sub SetBoundaryProbes(ByVal probes As Variant)
    mBoundaryProbes = probes
End Sub

' Возвращает отдельную копию границ для следующей геометрической операции.
' Изменения рабочего массива не затрагивают сохранённую область.
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

' Запоминает, для какой инженерной проверки и группы арматуры подготовлена область,
' без нового расчёта. Эти данные передаёт расчёт проверки, а не класс записи на лист.
Public Sub SetOwner(ByVal combinationID As String, ByVal standard As String, ByVal anchorID As String)
    mOwnerID = combinationID: mStandard = standard: mAnchorID = anchorID
End Sub

' Возвращает точный отрезок или дугу границы для аннотаций и экспорта.
' Центр и угол дуги сохраняются без замены промежуточными точками.
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

' Возвращает точки начала Q и конца P бетонных участков на контрольных прямых
' для записи и проверки результата. Empty означает, что такие точки
' не требовались, например при центральном растяжении.
Public Property Get BoundaryProbes() As Variant
    BoundaryProbes = mBoundaryProbes
End Property
' Возвращает число контрольных прямых, для которых сохранена проверка границ бетонного участка.
Public Property Get BoundaryProbeCount() As Long
    If IsArray(mBoundaryProbes) Then BoundaryProbeCount = UBound(mBoundaryProbes, 1)
End Property
' Возвращает число замкнутых границ подготовленной бетонной области.
Public Property Get LoopCount() As Long
    LoopCount = mLoopCount
End Property

' Возвращает назначение контура: наружная граница или отверстие.
' При подготовке аннотаций заново определять ориентацию границы не нужно.
Public Property Get LoopIsOpening(ByVal index As Long) As Boolean
    If index < 1 Or index > mLoopCount Then Err.Raise vbObjectError + 5400, "CConcreteRegion", "Номер замкнутого контура области вне допустимого диапазона."
    LoopIsOpening = mLoopOpening(index)
End Property
' Возвращает площадь области по точным контурам, мм2; она рассчитывается отдельно от площади
' бетонной сетки.
Public Property Get Area() As Double
    Area = mArea
End Property
' Возвращает сведения о том, по каким исходным границам получена область.
Public Property Get Source() As String
    Source = mSource
End Property
' Ревизия исходной модели, к которой относится область.
Public Property Get SectionRevision() As Long
    SectionRevision = mSectionRevision
End Property
' Возвращает номер сочетания нагрузок, к результату которого относится область.
Public Property Get OwnerID() As String
    OwnerID = mOwnerID
End Property
' Возвращает норматив, по которому проверяется эта область; геометрический класс не рассчитывает
' нормативные коэффициенты.
Public Property Get Standard() As String
    Standard = mStandard
End Property
' Возвращает ID арматурной группы или стержня, для которых сохранена область.
Public Property Get AnchorID() As String
    AnchorID = mAnchorID
End Property

' Создаёт независимую копию области вместе с её назначением и источником.
' Копирует все замкнутые границы и дуги, сохраняя точную геометрию.
Public Function Clone() As CConcreteRegion
    Dim result As CConcreteRegion
    Set result = New CConcreteRegion
    result.SetGeometry mEdges, mCount, mLoopCount, mArea, mSource, mSectionRevision, mLoopOpening
    result.SetOwner mOwnerID, mStandard, mAnchorID
    result.SetBoundaryProbes mBoundaryProbes
    Set Clone = result
End Function
