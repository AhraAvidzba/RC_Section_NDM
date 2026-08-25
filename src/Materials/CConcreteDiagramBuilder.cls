VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConcreteDiagramBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Построитель диаграммы бетона по СП 63
' ==========================================================================
' Builder превращает пользовательские характеристики бетона в точки
' TwoLine/ThreeLine. Он работает во внутренних единицах: напряжения в МПа,
' деформации безразмерные, сжатие имеет отрицательный знак.
'
' В текущей реализации используются кратковременные деформационные параметры
' СП 63. Для CrackedNDS/a_crc это сохраняет принятую ранее схему проекта:
' коэффициент длительности раскрытия трещин учитывается в формуле a_crc, а
' диаграммы НДС строятся по кратковременным характеристикам.

Private Const EPS_B0_SHORT As Double = 0.002 ' εb0 для кратковременных нагрузок.
Private Const EPS_BT0_SHORT As Double = 0.0001 ' εbt0 для кратковременных нагрузок.
Private Const EPS_B1_RED_SHORT As Double = 0.0015 ' εb1,red для двухлинейной сжатой диаграммы.
Private Const EPS_BT1_RED_SHORT As Double = 0.00008 ' εbt1,red для двухлинейной растянутой диаграммы.
Private Const EPS_B2_SHORT As Double = 0.0035 ' εb2 для кратковременных нагрузок.
Private Const EPS_BT2_SHORT As Double = 0.00015 ' εbt2 для кратковременных нагрузок.

Private mCompressionLimit As Double ' Последняя построенная предельная деформация сжатого бетона.
Private mTensionLimit As Double ' Последняя построенная предельная деформация растянутого бетона.
Private mTensionEnabled As Boolean ' True, если в диаграмму включена растянутая ветвь.

' Строит CConcreteDiagramMaterial для заданной ГПС, типа диаграммы и режима
' растянутого бетона.
Public Function BuildMaterial(ByVal parameters As CConcreteMaterialParameters, _
        ByVal limitStateGroup As String, ByVal diagramType As String, _
        ByVal tensionBehavior As String) As CConcreteDiagramMaterial
    If parameters Is Nothing Then Err.Raise vbObjectError + 3220, "CConcreteDiagramBuilder", "Параметры бетона не переданы."

    diagramType = NormalizeDiagramType(diagramType)
    tensionBehavior = NormalizeTensionBehavior(tensionBehavior)

    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths parameters, limitStateGroup, rb, rbt

    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet

    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        AddThreeLineCompression points, rb, parameters.Eb
    Else
        AddTwoLineCompression points, rb
    End If

    points.AddPoint 0#, 0#

    mTensionEnabled = (StrComp(tensionBehavior, "UseDiagram", vbTextCompare) = 0)
    If mTensionEnabled Then
        If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
            AddThreeLineTension points, rbt, parameters.Ebt
        Else
            AddTwoLineTension points, rbt
        End If
    End If

    mCompressionLimit = -EPS_B2_SHORT
    If mTensionEnabled Then mTensionLimit = EPS_BT2_SHORT Else mTensionLimit = 0#

    Dim material As CConcreteDiagramMaterial
    Set material = New CConcreteDiagramMaterial
    material.InitializeFromPointSet points
    Set BuildMaterial = material
End Function

' Добавляет двухлинейную сжатую ветвь: линейный участок до εb1,red и плато до εb2.
Private Sub AddTwoLineCompression(ByVal points As CMaterialPointSet, ByVal rb As Double)
    points.AddPoint -EPS_B2_SHORT, -rb
    points.AddPoint -EPS_B1_RED_SHORT, -rb
End Sub

' Добавляет трехлинейную сжатую ветвь с точкой 0.6Rb.
Private Sub AddThreeLineCompression(ByVal points As CMaterialPointSet, ByVal rb As Double, ByVal eb As Double)
    Dim epsB1 As Double
    epsB1 = 0.6 * rb / eb
    points.AddPoint -EPS_B2_SHORT, -rb
    points.AddPoint -EPS_B0_SHORT, -rb
    points.AddPoint -epsB1, -0.6 * rb
End Sub

' Добавляет двухлинейную растянутую ветвь бетона.
Private Sub AddTwoLineTension(ByVal points As CMaterialPointSet, ByVal rbt As Double)
    points.AddPoint EPS_BT1_RED_SHORT, rbt
    points.AddPoint EPS_BT2_SHORT, rbt
End Sub

' Добавляет трехлинейную растянутую ветвь с точкой 0.6Rbt.
Private Sub AddThreeLineTension(ByVal points As CMaterialPointSet, ByVal rbt As Double, ByVal ebt As Double)
    Dim epsBt1 As Double
    epsBt1 = 0.6 * rbt / ebt
    points.AddPoint epsBt1, 0.6 * rbt
    points.AddPoint EPS_BT0_SHORT, rbt
    points.AddPoint EPS_BT2_SHORT, rbt
End Sub

' Возвращает сопротивления нужной группы предельных состояний.
Private Sub ConcreteStrengths(ByVal parameters As CConcreteMaterialParameters, ByVal limitStateGroup As String, _
        ByRef rb As Double, ByRef rbt As Double)
    If StrComp(Trim$(limitStateGroup), "II", vbTextCompare) = 0 Or StrComp(Trim$(limitStateGroup), "2", vbTextCompare) = 0 Then
        rb = parameters.RbSLS
        rbt = parameters.RbtSLS
    Else
        rb = parameters.RbULS
        rbt = parameters.RbtULS
    End If
End Sub

' Нормализует пользовательский тип диаграммы.
Private Function NormalizeDiagramType(ByVal diagramType As String) As String
    Select Case LCase$(Trim$(diagramType))
        Case "threeline", "three"
            NormalizeDiagramType = "ThreeLine"
        Case Else
            NormalizeDiagramType = "TwoLine"
    End Select
End Function

' Нормализует режим растянутого бетона.
Private Function NormalizeTensionBehavior(ByVal tensionBehavior As String) As String
    Select Case LCase$(Trim$(tensionBehavior))
        Case "usediagram", "use"
            NormalizeTensionBehavior = "UseDiagram"
        Case Else
            NormalizeTensionBehavior = "Ignore"
    End Select
End Function

Public Property Get CompressionLimit() As Double
    CompressionLimit = mCompressionLimit
End Property

Public Property Get TensionLimit() As Double
    TensionLimit = mTensionLimit
End Property

Public Property Get TensionEnabled() As Boolean
    TensionEnabled = mTensionEnabled
End Property
