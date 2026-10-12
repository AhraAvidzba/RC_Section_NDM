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

' Заполняет код, причину и связи диагностической записи с исходной геометрией или правилом.
Friend Sub Initialize(ByVal code As String, ByVal entityID As String, ByVal relatedEntityID As String, _
        ByVal reason As String, Optional ByVal detail1 As Double = 0#, Optional ByVal detail2 As Double = 0#, _
        Optional ByVal sourceReference As String = "")
    If Len(mCode) > 0 Then Err.Raise vbObjectError + 5530, "CGeometryDiagnostic", "Диагностика уже опубликована."
    mCode = code: mEntity = entityID: mRelated = relatedEntityID: mReason = reason
    mDetail1 = detail1: mDetail2 = detail2: mSource = sourceReference
End Sub

' Машинный код причины; назначается напрямую, не извлекается из русского текста.
Public Property Get Code() As String: Code = mCode: End Property
' Смысловой ID сущности отказа, если он известен этому сервису.
Public Property Get EntityID() As String: EntityID = mEntity: End Property
' ID связанного сегмента/позиции/стержня для локализации причины.
Public Property Get RelatedEntityID() As String: RelatedEntityID = mRelated: End Property
' Человекочитаемая причина; не selector/status.
Public Property Get Reason() As String: Reason = mReason: End Property
' Первое численное поле конкретного Code; единицы определяет Code.
Public Property Get Detail1() As Double: Detail1 = mDetail1: End Property
' Второе численное поле конкретного Code.
Public Property Get Detail2() As Double: Detail2 = mDetail2: End Property
' Реальный адрес/программный источник, добавленный SourceMap или правилом.
Public Property Get SourceReference() As String: SourceReference = mSource: End Property
