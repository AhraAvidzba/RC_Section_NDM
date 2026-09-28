VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Запрос именованного НДС
' ==========================================================================
' CStateRequest отделяет физический запрос состояния от технических настроек
' решения. В ключ переиспользования входят только величины, которые определяют
' сам результат: тип named-state, целевые N/Mx/My, роль и спецификация
' материалов, а также фактический режим диаграммы. Warm-start, стартовая
' плоскость и retry-политика в ключ не входят: они меняют шанс сходимости, но
' не должны создавать другое физическое состояние.

Private mStateType As ESectionStateType ' Какое named-state нужно получить.
Private mMaterialRole As ECalculationPurpose ' Роль материала: Strength, CrackInitiation, CrackedState.
Private mMaterialSpec As CMaterialModelSpec ' Фактическая спецификация диаграмм для этого state.
Private mTargetN As Double ' Целевая продольная сила, Н.
Private mTargetMx As Double ' Целевой момент Mx, Н*мм.
Private mTargetMy As Double ' Целевой момент My, Н*мм.
Private mUseDirectStateExtension As Boolean ' True, если физическая диаграмма state включает technical extension.
Private mAllowRetryStarts As Boolean ' Solve-option: можно ли пробовать дополнительные стартовые плоскости.

' Заполняет запрос всеми физическими параметрами state и отдельным флагом
' solve-options. MaterialSpec копируется, чтобы ключ не зависел от дальнейших
' изменений профиля.
Public Sub Initialize(ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal useDirectStateExtension As Boolean, _
        Optional ByVal allowRetryStarts As Boolean = False)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3980, "CStateRequest", "MaterialModelSpec не передан."
    mStateType = stateType
    mMaterialRole = materialRole
    Set mMaterialSpec = materialSpec.Clone
    mTargetN = targetN
    mTargetMx = targetMx
    mTargetMy = targetMy
    mUseDirectStateExtension = useDirectStateExtension
    mAllowRetryStarts = allowRetryStarts
End Sub

' Возвращает ключ эквивалентности физического state. Диагностика, комментарии,
' warm-start и retry-параметры сюда сознательно не входят.
Public Function EquivalenceKey() As String
    EquivalenceKey = CStr(CLng(mStateType)) & "|" & _
        CStr(CLng(mMaterialRole)) & "|" & _
        mMaterialSpec.SpecKey & "|" & _
        BoolKey(mUseDirectStateExtension) & "|" & _
        NumberKey(mTargetN) & "|" & _
        NumberKey(mTargetMx) & "|" & _
        NumberKey(mTargetMy)
End Function

' Форматирует число без региональных запятых, чтобы один и тот же запрос
' давал одинаковый ключ при разных настройках Windows/Excel.
Private Function NumberKey(ByVal value As Double) As String
    NumberKey = Replace$(Format$(value, "0.###############"), ",", ".")
End Function

' Коротко кодирует логический флаг для ключа repository.
Private Function BoolKey(ByVal value As Boolean) As String
    If value Then BoolKey = "1" Else BoolKey = "0"
End Function

Public Property Get StateType() As ESectionStateType
    StateType = mStateType
End Property

Public Property Get MaterialRole() As ECalculationPurpose
    MaterialRole = mMaterialRole
End Property

Public Property Get MaterialSpec() As CMaterialModelSpec
    Set MaterialSpec = mMaterialSpec
End Property

Public Property Get TargetN() As Double
    TargetN = mTargetN
End Property

Public Property Get TargetMx() As Double
    TargetMx = mTargetMx
End Property

Public Property Get TargetMy() As Double
    TargetMy = mTargetMy
End Property

Public Property Get UseDirectStateExtension() As Boolean
    UseDirectStateExtension = mUseDirectStateExtension
End Property

Public Property Get AllowRetryStarts() As Boolean
    AllowRetryStarts = mAllowRetryStarts
End Property
