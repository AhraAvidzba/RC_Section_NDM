VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGeometryDiagnostic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Машинная причина отказа геометрии; скалярный снимок без Excel и статуса НДС.
Private mCode As String, mEntity As String, mRelated As String, mReason As String, mSource As String
Private mDetail1 As Double, mDetail2 As Double

Friend Sub Initialize(ByVal code As String, ByVal entityID As String, ByVal relatedEntityID As String, _
        ByVal reason As String, Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#, _
        Optional ByVal sourceReference As String = "")
    If Len(mCode) > 0 Then Err.Raise vbObjectError + 5530, "CGeometryDiagnostic", "Диагностика уже опубликована."
    mCode = code: mEntity = entityID: mRelated = relatedEntityID: mReason = reason
    mDetail1 = detail1: mDetail2 = detail2: mSource = sourceReference
End Sub

Public Property Get Code() As String: Code = mCode: End Property
Public Property Get EntityID() As String: EntityID = mEntity: End Property
Public Property Get RelatedEntityID() As String: RelatedEntityID = mRelated: End Property
Public Property Get Reason() As String: Reason = mReason: End Property
Public Property Get Detail1() As Double: Detail1 = mDetail1: End Property
Public Property Get Detail2() As Double: Detail2 = mDetail2: End Property
Public Property Get SourceReference() As String: SourceReference = mSource: End Property
