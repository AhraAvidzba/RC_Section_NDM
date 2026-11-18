VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPostResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемый вывод отдельной проверки T и продольной арматуры T+NM
' ==========================================================================
' Матрицы содержат внутренние единицы; преобразование для Excel выполняет
' адаптер вывода. Ни результат сочетания, ни сохранённое НДС не изменяются.

Private mSummary As Variant, mBars As Variant
Private mStateFailureReason As String ' Отдельная причина недоступности НДС для адресной диагностики его настройки.

' Принимает полностью подготовленные матрицы одного сочетания.
Friend Sub Initialize(ByRef summary As Variant, ByRef bars As Variant, Optional ByVal stateFailureReason As String = "")
    mSummary = summary: mBars = bars
    mStateFailureReason = stateFailureReason
End Sub

' Создаёт адресуемую ошибку отдельного T-этапа без изменения метаданных НДМ.
Friend Sub InitializeFailure(ByVal identifier As String, ByVal reason As String, Optional ByVal internalError As Boolean = False)
    Dim data(1 To 1, 1 To TORSION_POST_COLUMNS) As Variant
    Dim meta As New CResultMeta, policy As New CResultStatusPolicy
    If internalError Then
        meta.SetResult rsInternalError, rcInternalError, rkGeneric, reason
    Else
        meta.SetResult rsInvalidInput, rcInvalidInput, rkGeneric, reason
    End If
    data(1, 1) = identifier
    data(1, 15) = policy.ExternalStatus(meta): data(1, 21) = data(1, 15)
    data(1, 28) = data(1, 15): data(1, 36) = data(1, 15): data(1, 38) = data(1, 15)
    data(1, 39) = reason
    mSummary = data
End Sub

' Отдаёт копию строки сводки; вызывающий может преобразовать её единицы.
Public Property Get Summary() As Variant
    Summary = mSummary
End Property

' Отдаёт копию детализации; Empty означает отсутствие пригодного распределения.
Public Property Get Bars() As Variant
    Bars = mBars
End Property

' Возвращает причину отсутствия сохранённого НДС отдельно от ошибки проверки бетона.
Public Property Get StateFailureReason() As String
    StateFailureReason = mStateFailureReason
End Property
