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
' Хранит итог проверки сжатого бетона по максимальному напряжению
' текущего состояния CrackedState. При отсутствии пригодного состояния
' вызывающий расчёт передаёт rsBlockedByDependency, а CResultStatusPolicy
' формирует внешнее обозначение N/A.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mMeta As CResultMeta ' Типизированный статус проверки продольных трещин.
Private mMaxCompressionStress As Double ' Максимальное сжатие бетона в текущем CrackedState.
Private mRbMc2 As Double ' Допустимое Rb,mc2.
Private mUtilization As Double ' Отношение наибольшего сжимающего напряжения бетона к сопротивлению Rb,mc2.

' Создает нейтральный not-applicable результат до фактического расчета.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkLongitudinalCrack
End Sub

' Сохраняет итог уже выполненной проверки без нового решения равновесия.
Public Sub Initialize(ByVal meta As CResultMeta, ByVal maxCompressionStress As Double, _
        ByVal rbMc2 As Double, ByVal utilization As Double)
    AssertWritable
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkLongitudinalCrack, _
            "Не сформированы статус и причина результата проверки продольных трещин.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If

    mMaxCompressionStress = maxCompressionStress
    mRbMc2 = rbMc2
    mUtilization = utilization
End Sub

' Возвращает независимую копию статуса продольных трещин с их собственной причиной.
' Готовый результат через полученную копию изменить нельзя.
Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Возвращает обозначение статуса для пользователя по общим правилам,
' без отдельной изменяемой строки статуса.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

' Максимальное сжатие бетона в текущем CrackedState.
Public Property Get MaxCompressionStress() As Double
    MaxCompressionStress = mMaxCompressionStress
End Property

' Допустимое Rb,mc2.
Public Property Get RbMc2() As Double
    RbMc2 = mRbMc2
End Property

' Сохранённый коэффициент использования проверяемого предела; смысл предела задаёт результат
' этого класса.
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

' Завершает заполнение результата, запрещая изменение чисел, статуса и пояснения.
' Повторный вызов безопасен; сохранённые НДС этот объект не изменяет.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CLongitudinalCrackResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
