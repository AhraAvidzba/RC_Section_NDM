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

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mMeta As CResultMeta ' Инженерный итог проверки прямого НДС.
Private mState As CSectionStateResult ' Единственный численный снимок named state.
Private mExtensionEnabled As Boolean ' Настройка глобального Extension при запуске.
Private mConcreteCompressionLimit As Double ' Физический предел сжатия использованной модели.
Private mConcreteTensionLimit As Double ' Физический предел активной растянутой ветви.
Private mConcreteTensionEnabled As Boolean ' Ignore не участвует в деформационном запасе.
Private mSteelCompressionLimit As Double ' Предел арматуры для отрицательного знака.
Private mSteelTensionLimit As Double ' Предел арматуры для положительного знака.

' Создает нейтральный result без фиктивного физического состояния.
Private Sub Class_Initialize()
    Clear
End Sub

' Сбрасывает состояние и meta для безопасного повторного использования.
Public Sub Clear()
    AssertWritable
    Set mState = Nothing
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkDirectState
    mExtensionEnabled = False
    mConcreteCompressionLimit = 0#
    mConcreteTensionLimit = 0#
    mConcreteTensionEnabled = False
    mSteelCompressionLimit = 0#
    mSteelTensionLimit = 0#
End Sub

' Устанавливает итоговую meta; пустая meta является нарушением контракта.
Public Sub SetMeta(ByVal meta As CResultMeta)
    AssertWritable
    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkDirectState, _
            "Не сформирован результат проверки прямого НДС.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If
End Sub

' Привязывает result к сохраненному State без копии численных полей.
' Ошибка получения State сохраняется в его meta, а не превращается в успех.
Public Sub StoreStateResult(ByVal stateResult As CSectionStateResult, ByVal extensionEnabledValue As Boolean, _
        Optional ByVal physicalConcrete As CMaterialDiagram = Nothing, _
        Optional ByVal physicalSteel As CMaterialDiagram = Nothing)
    AssertWritable
    Clear
    Set mState = stateResult
    mExtensionEnabled = extensionEnabledValue
    If Not physicalConcrete Is Nothing Then
        mConcreteCompressionLimit = physicalConcrete.UltimateCompressionStrain
        mConcreteTensionLimit = physicalConcrete.UltimateTensionStrain
        mConcreteTensionEnabled = physicalConcrete.HasPhysicalTensionLimit
    End If
    If Not physicalSteel Is Nothing Then
        mSteelCompressionLimit = physicalSteel.UltimateCompressionStrain
        mSteelTensionLimit = physicalSteel.UltimateTensionStrain
    End If
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

' Подтверждает равновесие только при наличии сохраненного State.
' False при Nothing не означает численную попытку или физический отказ.
Public Property Get StateConverged() As Boolean
    If Not mState Is Nothing Then StateConverged = mState.Converged
End Property

' Возвращает использование extension именно конечным State, не промежуточной
' итерацией. При отсутствии State признак не выставляется.
Public Property Get ExtensionUsed() As Boolean
    If Not mState Is Nothing Then ExtensionUsed = mState.ExtensionUsed
End Property

Public Property Get ExtensionEnabled() As Boolean
    ExtensionEnabled = mExtensionEnabled
End Property

' Вычисляет запас по реально присутствующим знакам деформаций и исходным
' физическим пределам. Численное продолжение не увеличивает этот запас.
' При отсутствии проверенного состояния возвращает 0 как признак недоступности.
Public Property Get StrainReserve() As Double
    If mState Is Nothing Then Exit Property
    If Not mState.Converged Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsFinished(Status) Then Exit Property
    Dim reserve As Double
    reserve = 1E+100
    AddStrainReserve reserve, mState.HasConcreteCompression, mConcreteCompressionLimit, mState.MinConcreteStrain
    If mConcreteTensionEnabled Then _
        AddStrainReserve reserve, mState.HasConcreteTension, mConcreteTensionLimit, mState.MaxConcreteStrain
    AddStrainReserve reserve, mState.HasSteelCompression, mSteelCompressionLimit, mState.MinSteelStrain
    AddStrainReserve reserve, mState.HasSteelTension, mSteelTensionLimit, mState.MaxSteelStrain
    If reserve < 1E+90 Then StrainReserve = reserve
End Property

' Добавляет применимый знак материала в минимум; нулевая деформация не дает
' искусственного бесконечного запаса и отсутствующая зона не участвует.
Private Sub AddStrainReserve(ByRef reserve As Double, ByVal hasValue As Boolean, _
        ByVal limitValue As Double, ByVal actualValue As Double)
    If Not hasValue Then Exit Sub
    If Abs(limitValue) <= 0.000000000000001 Then Exit Sub
    If Abs(actualValue) <= 0.000000000000001 Then Exit Sub
    Dim candidate As Double
    candidate = Abs(limitValue) / Abs(actualValue)
    If candidate < reserve Then reserve = candidate
End Sub

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Закрывает также вложенное НДС без копирования сетки. Повторный вызов
' безопасен; сохраненная ссылка не позволяет перезаполнить direct result.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mState Is Nothing Then mState.Freeze
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CDirectStateResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
