VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSteelDiagramBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Построитель диаграммы арматуры по СП 63
' ==========================================================================
' Builder формирует TwoLine/ThreeLine диаграмму ненапрягаемой арматуры по
' выбранной группе предельных состояний. Сжатая и растянутая ветви строятся
' симметрично по структуре, но могут использовать разные пользовательские
' сопротивления и модули Rsc/Rs, Esc/Es.
'
' Внутри проекта сжатие имеет отрицательный знак, поэтому все точки сжатой
' ветви записываются с отрицательной деформацией и отрицательным напряжением.

Private Const EPS_S2_TWO_LINE As Double = 0.025 ' Предельная деформация TwoLine по СП 63 п. 6.2.14.
Private Const EPS_S2_THREE_LINE As Double = 0.015 ' Предельная деформация ThreeLine по СП 63 п. 6.2.15.

Private mCompressionLimit As Double ' Последняя построенная предельная деформация сжатой арматуры.
Private mTensionLimit As Double ' Последняя построенная предельная деформация растянутой арматуры.

' Строит CSteelDiagramMaterial для заданной ГПС и типа диаграммы.
Public Function BuildMaterial(ByVal parameters As CSteelMaterialParameters, _
        ByVal limitStateGroup As String, ByVal diagramType As String) As CSteelDiagramMaterial
    If parameters Is Nothing Then Err.Raise vbObjectError + 3230, "CSteelDiagramBuilder", "Параметры арматуры не переданы."

    diagramType = NormalizeDiagramType(diagramType)

    Dim rs As Double
    Dim rsc As Double
    Dim es As Double
    Dim esc As Double
    SteelStrengths parameters, limitStateGroup, rsc, rs, esc, es

    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        AddThreeLineCompression points, rsc, esc
        points.AddPoint 0#, 0#
        AddThreeLineTension points, rs, es
        mCompressionLimit = -EPS_S2_THREE_LINE
        mTensionLimit = EPS_S2_THREE_LINE
    Else
        AddTwoLineCompression points, rsc, esc
        points.AddPoint 0#, 0#
        AddTwoLineTension points, rs, es
        mCompressionLimit = -EPS_S2_TWO_LINE
        mTensionLimit = EPS_S2_TWO_LINE
    End If

    Dim material As CSteelDiagramMaterial
    Set material = New CSteelDiagramMaterial
    material.InitializeFromPointSet points
    Set BuildMaterial = material
End Function

' Двухлинейная сжатая ветвь: линейный участок до Rsc и плато до εsc2.
Private Sub AddTwoLineCompression(ByVal points As CMaterialPointSet, ByVal rsc As Double, ByVal esc As Double)
    Dim epsSc0 As Double
    epsSc0 = rsc / esc
    points.AddPoint -EPS_S2_TWO_LINE, -rsc
    points.AddPoint -epsSc0, -rsc
End Sub

' Двухлинейная растянутая ветвь: линейный участок до Rs и плато до εs2.
Private Sub AddTwoLineTension(ByVal points As CMaterialPointSet, ByVal rs As Double, ByVal es As Double)
    Dim epsS0 As Double
    epsS0 = rs / es
    points.AddPoint epsS0, rs
    points.AddPoint EPS_S2_TWO_LINE, rs
End Sub

' Трехлинейная сжатая ветвь для арматуры с условным пределом текучести.
Private Sub AddThreeLineCompression(ByVal points As CMaterialPointSet, ByVal rsc As Double, ByVal esc As Double)
    Dim epsSc0 As Double
    Dim epsSc1 As Double
    Dim epsScPl As Double
    epsSc0 = rsc / esc + 0.002
    epsSc1 = 0.9 * rsc / esc
    epsScPl = 2# * epsSc0 - epsSc1
    points.AddPoint -EPS_S2_THREE_LINE, -1.1 * rsc
    points.AddPoint -epsScPl, -1.1 * rsc
    points.AddPoint -epsSc0, -rsc
    points.AddPoint -epsSc1, -0.9 * rsc
End Sub

' Трехлинейная растянутая ветвь для арматуры с условным пределом текучести.
Private Sub AddThreeLineTension(ByVal points As CMaterialPointSet, ByVal rs As Double, ByVal es As Double)
    Dim epsS0 As Double
    Dim epsS1 As Double
    Dim epsSPl As Double
    epsS0 = rs / es + 0.002
    epsS1 = 0.9 * rs / es
    epsSPl = 2# * epsS0 - epsS1
    points.AddPoint epsS1, 0.9 * rs
    points.AddPoint epsS0, rs
    points.AddPoint epsSPl, 1.1 * rs
    points.AddPoint EPS_S2_THREE_LINE, 1.1 * rs
End Sub

' Возвращает сопротивления и модули нужной группы предельных состояний.
Private Sub SteelStrengths(ByVal parameters As CSteelMaterialParameters, ByVal limitStateGroup As String, _
        ByRef rsc As Double, ByRef rs As Double, ByRef esc As Double, ByRef es As Double)
    If StrComp(Trim$(limitStateGroup), "II", vbTextCompare) = 0 Or StrComp(Trim$(limitStateGroup), "2", vbTextCompare) = 0 Then
        rsc = parameters.RscSLS
        rs = parameters.RsSLS
    Else
        rsc = parameters.RscULS
        rs = parameters.RsULS
    End If
    esc = parameters.Esc
    es = parameters.Es
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

Public Property Get CompressionLimit() As Double
    CompressionLimit = mCompressionLimit
End Property

Public Property Get TensionLimit() As Double
    TensionLimit = mTensionLimit
End Property
