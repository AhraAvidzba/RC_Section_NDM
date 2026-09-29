VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Сводный результат расчета трещин одного LC
' ==========================================================================
' CCrackResult объединяет три самостоятельных смысловых результата:
' образование нормальной трещины, ширину раскрытия нормальной трещины и
' продольные трещины по сжатому бетону. Он не пишет Results и не запускает
' solver; это контейнер result-tree между orchestration-слоем и batch-снимком.

Private mFormation As CCrackFormationResult ' Результат поиска Ncrc/Mxy,crc.
Private mCurrentStateMeta As CResultMeta ' Meta обязательного CurrentCrackedState, если он уже решался.
Private mWidth As CCrackWidthResult ' Результат формульного расчета a_crc.
Private mLongitudinal As CLongitudinalCrackResult ' Результат проверки продольных трещин.

' Заполняет все доступные ветви расчета трещин. Любая ветвь может быть Nothing,
' если расчет еще не дошел до нее или она не нужна для выбранного профиля.
Public Sub Initialize(ByVal formation As CCrackFormationResult, _
        ByVal currentStateMeta As CResultMeta, ByVal width As CCrackWidthResult, _
        ByVal longitudinal As CLongitudinalCrackResult)
    Set mFormation = formation
    If currentStateMeta Is Nothing Then
        Set mCurrentStateMeta = Nothing
    Else
        Set mCurrentStateMeta = currentStateMeta.Clone
    End If
    Set mWidth = width
    Set mLongitudinal = longitudinal
End Sub

Public Property Get Formation() As CCrackFormationResult
    Set Formation = mFormation
End Property

Public Property Get Width() As CCrackWidthResult
    Set Width = mWidth
End Property

Public Property Get CurrentStateMeta() As CResultMeta
    If Not mCurrentStateMeta Is Nothing Then Set CurrentStateMeta = mCurrentStateMeta.Clone
End Property

Public Property Get Longitudinal() As CLongitudinalCrackResult
    Set Longitudinal = mLongitudinal
End Property
