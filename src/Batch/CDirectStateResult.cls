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
' Результат НДС при заданном сочетании нагрузок
' ==========================================================================
' Хранит ссылку на CSectionStateResult и итог проверки по модели прочности.
' Поле деформаций, усилия, невязки и экстремумы берутся из сохранённого состояния.
' Внешний статус для пользователя формирует CResultStatusPolicy.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mMeta As CResultMeta ' Инженерный итог проверки прямого НДС.
Private mState As CSectionStateResult ' Сохранённое напряжённо-деформированное состояние сечения.
Private mExtensionEnabled As Boolean ' Разрешение численного продолжения диаграмм на момент запуска расчёта.
Private mConcreteCompressionLimit As Double ' Физический предел сжатия использованной модели.
Private mConcreteTensionLimit As Double ' Физический предел активной растянутой ветви.
Private mConcreteTensionEnabled As Boolean ' При отключённой растянутой ветви Ignore растяжение бетона не ограничивает запас по деформациям.
Private mSteelCompressionLimit As Double ' Предел арматуры для отрицательного знака.
Private mSteelTensionLimit As Double ' Предел арматуры для положительного знака.

' Создаёт пустой результат проверки, не подставляя несуществующее расчётное состояние.
Private Sub Class_Initialize()
    Clear
End Sub

' Очищает сохранённое НДС, статус и пояснение перед новым заполнением результата.
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

' Сохраняет итоговый статус и пояснение; отсутствие переданного объекта статуса вызывает ошибку.
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

' Связывает проверку с готовым сохранённым НДС без копирования чисел.
' Если получить НДС не удалось, его ошибка сохраняется в статусе состояния.
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

' Возвращает сохранённое НДС. Если расчёт не запрошен, возвращает Nothing.
Public Property Get StateResult() As CSectionStateResult
    Set StateResult = mState
End Property

' Возвращает копию статуса и пояснения, чтобы вызывающий код не мог изменить готовый результат.
Public Property Get Meta() As CResultMeta
    Set Meta = mMeta.Clone
End Property

' Возвращает обозначение статуса для пользователя по сохранённому результату проверки.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

' Показывает, сохранён ли объект напряжённо-деформированного состояния.
Public Property Get StateAvailable() As Boolean
    StateAvailable = Not mState Is Nothing
End Property

' Подтверждает равновесие только по сохранённому НДС.
' Если НДС отсутствует, False означает отсутствие решения, а не доказанное разрушение.
Public Property Get StateConverged() As Boolean
    If Not mState Is Nothing Then StateConverged = mState.Converged
End Property

' Показывает, использовано ли численное продолжение диаграмм в конечном НДС.
' При отсутствии НДС возвращает False; промежуточные итерации не учитываются.
Public Property Get ExtensionUsed() As Boolean
    If Not mState Is Nothing Then ExtensionUsed = mState.ExtensionUsed
End Property

' Разрешение численного продолжения диаграмм на момент запуска расчёта.
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
    If Not policy.IsMetaFinished(mMeta) Then Exit Property
    Dim reserve As Double
    reserve = 1E+100
    AddStrainReserve reserve, mState.HasConcreteCompression, mConcreteCompressionLimit, mState.MinConcreteStrain
    If mConcreteTensionEnabled Then _
        AddStrainReserve reserve, mState.HasConcreteTension, mConcreteTensionLimit, mState.MaxConcreteStrain
    AddStrainReserve reserve, mState.HasSteelCompression, mSteelCompressionLimit, mState.MinSteelStrain
    AddStrainReserve reserve, mState.HasSteelTension, mSteelTensionLimit, mState.MaxSteelStrain
    If reserve < 1E+90 Then StrainReserve = reserve
End Property

' Учитывает запас для заданного знака деформаций материала при выборе минимального запаса.
' Отсутствующая зона и нулевая деформация не создают дополнительный запас.
Private Sub AddStrainReserve(ByRef reserve As Double, ByVal hasValue As Boolean, _
        ByVal limitValue As Double, ByVal actualValue As Double)
    If Not hasValue Then Exit Sub
    If Abs(limitValue) <= 0.000000000000001 Then Exit Sub
    If Abs(actualValue) <= 0.000000000000001 Then Exit Sub
    Dim candidate As Double
    candidate = Abs(limitValue) / Abs(actualValue)
    If candidate < reserve Then reserve = candidate
End Sub

' Завершает заполнение проверки и связанного НДС, запрещая дальнейшие изменения.
' Сетка не копируется; повторный вызов безопасен.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mState Is Nothing Then mState.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CDirectStateResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
