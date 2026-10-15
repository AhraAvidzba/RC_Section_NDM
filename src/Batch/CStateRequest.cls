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
' Запрос напряжённо-деформированного состояния сечения
' ==========================================================================
' Хранит тип состояния, целевые N/Mx/My, назначение модели материалов
' и фактический режим диаграмм. Эти параметры образуют ключ повторного использования.
' Начальное приближение и настройки повторных попыток влияют на сходимость,
' но не изменяют физическую постановку запроса и потому не входят в этот ключ.

Private mStateType As ESectionStateType ' Тип расчётного состояния, которое требуется получить.
Private mMaterialRole As ECalculationPurpose ' Назначение материалов: прочность, образование трещины или НДС с трещинами.
Private mMaterialSpec As CMaterialModelSpec ' Фактически выбранные параметры диаграмм для данного НДС.
Private mTargetN As Double ' Целевая продольная сила, Н.
Private mTargetMx As Double ' Целевой момент Mx, Н*мм.
Private mTargetMy As Double ' Целевой момент My, Н*мм.
Private mUseDiagramExtension As Boolean ' Разрешено ли вспомогательное продолжение диаграмм для поиска равновесия.
Private mAllowRetryStarts As Boolean ' Разрешены ли дополнительные начальные приближения при несходимости.

' Сохраняет физические параметры запроса НДС и отдельное разрешение повторных стартов.
' Параметры материалов копируются, чтобы последующее изменение профиля
' не поменяло уже сформированный запрос.
Public Sub Initialize(ByVal stateType As ESectionStateType, _
        ByVal materialRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal useDiagramExtension As Boolean, _
        Optional ByVal allowRetryStarts As Boolean = False)
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3980, "CStateRequest", "MaterialModelSpec не передан."
    mStateType = stateType
    mMaterialRole = materialRole
    Set mMaterialSpec = materialSpec.Clone
    mTargetN = targetN
    mTargetMx = targetMx
    mTargetMy = targetMy
    mUseDiagramExtension = useDiagramExtension
    mAllowRetryStarts = allowRetryStarts
End Sub

' Строит ключ для поиска того же физического запроса в сохранённых результатах.
' Начальное приближение, диагностика, комментарии и параметры повторных попыток
' в этот ключ не входят.
Public Function EquivalenceKey() As String
    EquivalenceKey = CStr(CLng(mStateType)) & "|" & _
        CStr(CLng(mMaterialRole)) & "|" & _
        mMaterialSpec.SpecKey & "|" & _
        BoolKey(mUseDiagramExtension) & "|" & _
        NumberKey(mTargetN) & "|" & _
        NumberKey(mTargetMx) & "|" & _
        NumberKey(mTargetMy)
End Function

' Форматирует число без региональных запятых, чтобы один и тот же запрос
' давал одинаковый ключ при разных настройках Windows/Excel.
Private Function NumberKey(ByVal value As Double) As String
    NumberKey = Replace$(Format$(value, "0.###############"), ",", ".")
End Function

' Записывает логическое значение короткой строкой для ключа запроса в хранилище.
Private Function BoolKey(ByVal value As Boolean) As String
    If value Then BoolKey = "1" Else BoolKey = "0"
End Function

' Возвращает тип расчётного состояния, которое требуется получить.
Public Property Get StateType() As ESectionStateType
    StateType = mStateType
End Property

' Возвращает назначение материалов: прочность, образование трещины или НДС с трещинами.
Public Property Get MaterialRole() As ECalculationPurpose
    MaterialRole = mMaterialRole
End Property

' Возвращает отдельную копию параметров материалов. Её изменение не влияет
' на ключ запроса, уже переданного хранилищу состояний.
Public Property Get MaterialSpec() As CMaterialModelSpec
    If Not mMaterialSpec Is Nothing Then Set MaterialSpec = mMaterialSpec.Clone
End Property

' Целевая продольная сила, Н.
Public Property Get TargetN() As Double
    TargetN = mTargetN
End Property

' Целевой момент Mx, Н*мм.
Public Property Get TargetMx() As Double
    TargetMx = mTargetMx
End Property

' Целевой момент My, Н*мм.
Public Property Get TargetMy() As Double
    TargetMy = mTargetMy
End Property

' Показывает, разрешено ли вспомогательное продолжение диаграмм для поиска равновесия.
Public Property Get UseDiagramExtension() As Boolean
    UseDiagramExtension = mUseDiagramExtension
End Property

' Показывает, разрешены ли дополнительные начальные приближения при несходимости.
Public Property Get AllowRetryStarts() As Boolean
    AllowRetryStarts = mAllowRetryStarts
End Property
