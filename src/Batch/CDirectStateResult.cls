VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDirectStateResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Инженерный результат прямого НДС одного сочетания
' ==========================================================================
' Хранит ссылку на канонический CSectionStateResult и итоговую meta проверки
' модели прочности. Плоскость, усилия, невязки и экстремумы принадлежат State
' и не дублируются здесь. Класс не запускает solver; display формирует policy.

Private mMeta As CResultMeta ' Инженерный итог проверки прямого НДС.
Private mState As CSectionStateResult ' Единственный численный снимок named state.
Private mExtensionEnabled As Boolean ' Настройка глобального Extension при запуске.

' Создает нейтральный result без фиктивного физического состояния.
Private Sub Class_Initialize()
    Clear
End Sub

' Сбрасывает состояние и meta для безопасного повторного использования.
Public Sub Clear()
    Set mState = Nothing
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkDirectState
    mExtensionEnabled = False
End Sub

' Устанавливает итоговую meta; пустая meta является нарушением контракта.
Public Sub SetMeta(ByVal meta As CResultMeta)
    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkDirectState, _
            "Не сформирован результат проверки прямого НДС."
    Else
        Set mMeta = meta.Clone
    End If
End Sub

' Привязывает result к сохраненному State без копии численных полей.
' Ошибка получения State сохраняется в его meta, а не превращается в успех.
Public Sub StoreStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabledValue As Boolean)
    Set mState = stateResult
    mExtensionEnabled = extensionEnabledValue
    If stateResult Is Nothing Then
        SetMeta Nothing
    Else
        SetMeta stateResult.ResultMeta
    End If
End Sub

' Возвращает сохраненное State. При незапрошенном расчете возвращает Nothing.
Public Property Get StateResult() As CSectionStateResult
    Set StateResult = mState
End Property

' Meta возвращается копией: потребитель не может изменить итог проверки.
Public Property Get Meta() As CResultMeta
    Set Meta = mMeta.Clone
End Property

' Display и признаки доступности вычисляются из канонических объектов.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

Public Property Get StateAvailable() As Boolean
    StateAvailable = Not mState Is Nothing
End Property

Public Property Get StateConverged() As Boolean
    If Not mState Is Nothing Then StateConverged = mState.Converged
End Property

Public Property Get ExtensionUsed() As Boolean
    If Not mState Is Nothing Then ExtensionUsed = mState.ExtensionUsed
End Property

Public Property Get ExtensionEnabled() As Boolean
    ExtensionEnabled = mExtensionEnabled
End Property
