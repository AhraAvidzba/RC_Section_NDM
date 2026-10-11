VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarRuleOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Необязательные параметры сложного ряда. Фасад копирует их в definition,
' поэтому последующее изменение этого объекта не изменяет готовый рецепт.
Public StartSetback As Double
Public EndSetback As Double
Public EndMode As ERebarEndMode
Public RemainderPlacement As ERemainderPlacement
Public TangentOffset As Double
Public NormalOffset As Double
Public EverySecond As Boolean
Public DirectionX As Double
Public DirectionY As Double
Public SkipMissingProjection As Boolean
Public Publication As ERebarPublication
Public PositionOrder As ERebarPositionOrder
Public Coincidence As ERebarCoincidence
Public PlacementCheck As ERebarPlacementCheck
Public SourcePrefix As String
Public SourceSuffix As String
Public AnnotationGroup As String
Public Comment As String
Public SourceReference As String

Private Sub Class_Initialize()
    EndMode = IncludeBoth
    RemainderPlacement = RemainderCenter
    Publication = PublishByRule
    PositionOrder = AlongPath
    Coincidence = RejectCoincidence
    PlacementCheck = CheckCenter
End Sub

Friend Sub CopyTo(ByRef rule As TRebarRule)
    rule.StartSetback = StartSetback: rule.EndSetback = EndSetback
    rule.EndMode = EndMode: rule.RemainderPlacement = RemainderPlacement
    rule.TangentOffset = TangentOffset: rule.NormalOffset = NormalOffset
    rule.EverySecond = EverySecond
    rule.DirectionX = DirectionX: rule.DirectionY = DirectionY
    rule.SkipMissingProjection = SkipMissingProjection
    rule.Publication = Publication: rule.PlacementCheck = PlacementCheck
    rule.PositionOrder = PositionOrder: rule.Coincidence = Coincidence
    rule.SourcePrefix = SourcePrefix: rule.SourceSuffix = SourceSuffix
    rule.AnnotationGroup = AnnotationGroup: rule.Comment = Comment
    rule.SourceReference = SourceReference
End Sub
