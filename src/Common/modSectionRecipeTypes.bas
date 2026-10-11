Attribute VB_Name = "modSectionRecipeTypes"
Option Explicit

' Общие декларативные данные. Только внутренние единицы и скалярные снимки;
' никаких ссылок на Excel, решатель или специальные типы сечений.
Public Enum ESectionLoopRole
    LoopOuter = 1
    LoopOpening = 2
End Enum

Public Enum ERebarEndMode
    IncludeBoth = 1
    ExcludeBoth = 2
    IncludeStart = 3
    IncludeEnd = 4
    ClosedUniform = 5
End Enum

Public Enum ERebarDistribution
    ByCount = 1
    ByMaxStep = 2
    ByFixedStep = 3
    ExplicitPositions = 4
    ProjectFrom = 5
    DependentOffset = 6
End Enum

Public Enum ERemainderPlacement
    RemainderStart = 1
    RemainderEnd = 2
    RemainderCenter = 3
End Enum

Public Enum ERebarPublication
    PublishByRule = 1
    PublishPerSourcePosition = 2
End Enum

Public Enum ERebarPositionOrder
    AlongPath = 1
    FromPathMiddle = 2
End Enum

Public Enum ERebarCoincidence
    RejectCoincidence = 1
    MergeIdentical = 2
    KeepPublishedOwner = 3
End Enum

Public Enum ERebarPlacementCheck
    CheckCenter = 1
    CheckFullDisk = 2
End Enum

Public Type TSectionLoop
    PathID As String
    LoopID As String
    Role As ESectionLoopRole
    ParentID As String
    Comment As String
    ArcComment As String
    ContourIDs As Variant
End Type

Public Type TRebarRule
    RuleID As String
    Target As String
    Mode As ERebarDistribution
    Enabled As Boolean
    Diameter As Double
    SteelClass As String
    Count As Long
    Step As Double
    StartSetback As Double
    EndSetback As Double
    EndMode As ERebarEndMode
    RemainderPlacement As ERemainderPlacement
    SourceRuleID As String
    TangentOffset As Double
    NormalOffset As Double
    EverySecond As Boolean
    IncomingTangentAtJoint As Boolean
    DirectionX As Double
    DirectionY As Double
    SourceNormalSign As Double
    SuccessfulProjectionOrdinal As Boolean
    NewPublicationRoot As Boolean
    SkipMissingProjection As Boolean
    Publication As ERebarPublication
    PositionOrder As ERebarPositionOrder
    Coincidence As ERebarCoincidence
    PlacementCheck As ERebarPlacementCheck
    SourcePrefix As String
    SourceSuffix As String
    AnnotationGroup As String
    Comment As String
    SourceReference As String
    ExplicitStations As Variant
End Type

Public Type TSectionAnnotation
    AnnotationID As String
    Kind As String
    Target As String
    Name As String
    StartX As Double
    StartY As Double
    EndX As Double
    EndY As Double
    NormalX As Double
    NormalY As Double
    DirectionX As Double
    DirectionY As Double
    Value As Double
    DisplayIncrement As Double
    ReverseDirection As Boolean
    Comment As String
End Type

Public Type TSectionMeshOptions
    StepX As Double
    StepY As Double
    BoundarySubdivisions As Long
    MaterialID As Long
End Type

Public Type TRebarPosition
    X As Double
    Y As Double
    TangentX As Double
    TangentY As Double
    Station As Double
    RuleIndex As Long
    RootRuleIndex As Long
    RootOrdinal As Long
    RootPublicationRank As Long
    SourceOrdinal As Long
    SourcePositionID As String
    PathID As String
End Type

Public Sub SectionRecipeError(ByVal owner As String, ByVal reason As String, Optional ByVal sourceReference As String = "")
    If Len(sourceReference) > 0 Then reason = sourceReference & ": " & reason
    Err.Raise vbObjectError + 5530, owner, reason
End Sub

' Добавляет адрес исходных параметров к ошибке общего геометрического слоя;
' непредвиденные ошибки сохраняют свой исходный код и источник.
Public Sub SectionRecipeGeometryError(ByVal owner As String, ByVal sourceReference As String)
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    If source = "CSectionGeometryQuery" Or source = "CGeometryPath" Then SectionRecipeError owner, reason, sourceReference
    Err.Raise number, source, reason
End Sub
