VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLongitudinalCrackResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат проверки продольных трещин
' ==========================================================================
' CLongitudinalCrackResult хранит инженерский итог проверки сжатого бетона по
' максимальному напряжению текущего CrackedState. Проверка не решает НДС сама:
' при отсутствии обязательного CurrentCrackedState orchestration-слой должен
' передать meta rsBlockedByDependency, а внешний N/A сформирует policy.

Private mMeta As CResultMeta ' Типизированный статус проверки продольных трещин.
Private mMaxCompressionStress As Double ' Максимальное сжатие бетона в текущем CrackedState.
Private mRbMc2 As Double ' Допустимое Rb,mc2.
Private mUtilization As Double ' sigma_c,max / Rb,mc2.

' Создает нейтральный not-applicable результат до фактического расчета.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkLongitudinalCrack
End Sub

' Заполняет результат уже выполненной проверки без запуска solver-а.
Public Sub Initialize(ByVal meta As CResultMeta, ByVal maxCompressionStress As Double, _
        ByVal rbMc2 As Double, ByVal utilization As Double)
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkLongitudinalCrack, _
            "Meta результата продольных трещин не сформирована."
    Else
        Set mMeta = meta.Clone
    End If

    mMaxCompressionStress = maxCompressionStress
    mRbMc2 = rbMc2
    mUtilization = utilization
End Sub

Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Преобразует typed итог проверки в display, не храня вторую строку статуса.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

Public Property Get MaxCompressionStress() As Double
    MaxCompressionStress = mMaxCompressionStress
End Property

Public Property Get RbMc2() As Double
    RbMc2 = mRbMc2
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

' Возвращает запас завершенной проверки сжатого бетона; отсутствие результата
' или нулевая проверяемая величина не превращаются в бесконечный запас.
Public Property Get ReserveFactor() As Double
    If mUtilization <= 0# Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsFinished(Status) Then Exit Property
    ReserveFactor = 1# / mUtilization
End Property
