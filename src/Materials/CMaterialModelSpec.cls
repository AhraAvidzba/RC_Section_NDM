VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialModelSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' СПЕЦИФИКАЦИЯ МАТЕРИАЛЬНОЙ МОДЕЛИ ПРОФИЛЯ
' ==========================================================================
' CMaterialModelSpec хранит ровно тот набор пользовательских переключателей,
' по которому CMaterialModelProvider строит бетонную и арматурную диаграммы:
' ULS(I)/SLS(II), TwoLine/ThreeLine, учет растянутого бетона и тип диаграммы стали.
'
' Класс не строит диаграммы и не знает формул СП. Его задача - зафиксировать
' выбранную спецификацию в памяти результата, чтобы writer мог воспроизвести
' именно рассчитанное состояние, даже если другой профиль использует ту же
' роль материальной модели с другими настройками.

Private mValueSet As EMaterialValueSet ' ULS(I) или SLS(II).
Private mConcreteDiagram As EConcreteDiagramKind ' TwoLine или ThreeLine для бетона.
Private mConcreteTension As EConcreteTensionKind ' Ignore или UseDiagram для растянутого бетона.
Private mSteelDiagram As ESteelDiagramKind ' TwoLine или ThreeLine для арматуры.
Private mComplete As Boolean ' True, если все четыре поля успешно заданы.

' Заполняет спецификацию по текстам из rngCalculationProfiles. До успешного
' разбора всех полей она недоступна provider-у, включая повторный ввод.
Public Sub Initialize(ByVal valueSetText As String, ByVal concreteDiagramText As String, _
        ByVal concreteTensionText As String, ByVal steelDiagramText As String)
    mComplete = False
    mValueSet = MaterialValueSetFromText(valueSetText)
    mConcreteDiagram = ConcreteDiagramKindFromText(concreteDiagramText)
    mConcreteTension = ConcreteTensionKindFromText(concreteTensionText)
    mSteelDiagram = SteelDiagramKindFromText(steelDiagramText)
    mComplete = True
End Sub

' Сбрасывает спецификацию. Такой объект нельзя передавать provider-у.
Public Sub Clear()
    mComplete = False
End Sub

' Возвращает независимую копию, чтобы CSectionStateResult не зависел от
' дальнейших изменений исходного профиля.
Public Function Clone() As CMaterialModelSpec
    Dim copy As CMaterialModelSpec
    Set copy = New CMaterialModelSpec
    If mComplete Then
        copy.Initialize ValueSetText, ConcreteDiagramText, ConcreteTensionText, SteelDiagramText
    End If
    Set Clone = copy
End Function

' Короткий ключ удобно писать в диагностику и сравнивать в тестах.
Public Function SpecKey() As String
    If Not mComplete Then
        SpecKey = "<empty>"
    Else
        SpecKey = ValueSetText & "|" & ConcreteDiagramText & "|" & ConcreteTensionText & "|" & SteelDiagramText
    End If
End Function

Public Property Get IsComplete() As Boolean
    IsComplete = mComplete
End Property

Public Property Get ValueSet() As EMaterialValueSet
    ValueSet = mValueSet
End Property

Public Property Get ConcreteDiagram() As EConcreteDiagramKind
    ConcreteDiagram = mConcreteDiagram
End Property

Public Property Get ConcreteTension() As EConcreteTensionKind
    ConcreteTension = mConcreteTension
End Property

Public Property Get SteelDiagram() As ESteelDiagramKind
    SteelDiagram = mSteelDiagram
End Property

Public Property Get ValueSetText() As String
    ValueSetText = MaterialValueSetToText(mValueSet)
End Property

Public Property Get ConcreteDiagramText() As String
    ConcreteDiagramText = ConcreteDiagramKindToText(mConcreteDiagram)
End Property

Public Property Get ConcreteTensionText() As String
    ConcreteTensionText = ConcreteTensionKindToText(mConcreteTension)
End Property

Public Property Get SteelDiagramText() As String
    SteelDiagramText = SteelDiagramKindToText(mSteelDiagram)
End Property
