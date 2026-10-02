VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSteelMaterialParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Параметры арматуры для построения диаграмм
' ==========================================================================
' Класс хранит только исходные характеристики ненапрягаемой арматуры.
' Он не выбирает расчетный режим, не строит диаграмму и не содержит настроек
' трещин. Готовую CMaterialDiagram по этим параметрам строит
' CMaterialModelProvider, а solver работает уже с универсальной диаграммой.
'
' Rsc,ser и Esc оставлены как программные пользовательские параметры для
' симметрии сжатой и растянутой ветви и для исследовательских сценариев.
' Отдельные нормативные обозначения Rsc,ser и Esc для обычной арматуры в
' СП 63 не вводятся, поэтому типовые defaults: Rsc,ser = Rs,ser, Esc = Es.

Private mRsULS As Double ' Расчетное сопротивление растяжению для I ГПС, МПа.
Private mRscULS As Double ' Расчетное сопротивление сжатию для I ГПС, МПа.
Private mRsSLS As Double ' Сопротивление растяжению для II ГПС, МПа.
Private mRscSLS As Double ' Программное сопротивление сжатию для II ГПС, default RsSLS, МПа.
Private mEs As Double ' Модуль арматуры при растяжении, МПа.
Private mEsc As Double ' Программный модуль при сжатии, default Es, МПа.
Private mTwoLineEsc2 As Double ' Предельная сжатая деформация eps_sc2 для двухлинейной диаграммы.
Private mTwoLineEs2 As Double ' Предельная растянутая деформация eps_s2 для двухлинейной диаграммы.
Private mThreeLineEsc2 As Double ' Предельная сжатая деформация eps_sc2 для трехлинейной диаграммы.
Private mThreeLineEs2 As Double ' Предельная растянутая деформация eps_s2 для трехлинейной диаграммы.

' Читает параметры из Config и переводит прочности/модули в МПа.
' Коэффициенты раскрытия трещин здесь не читаются: phi2 задается
' отдельной настройкой SLS.Crack.Phi2 и не является свойством материала.
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3200, "CSteelMaterialParameters", "Настройки арматуры не переданы."

    mRsULS = settings.GetRequiredDouble("Steel.Rs.ULS")
    mRscULS = settings.GetRequiredDouble("Steel.Rsc.ULS")
    mRsSLS = settings.GetRequiredDouble("Steel.Rs.SLS")
    mRscSLS = settings.GetRequiredDouble("Steel.Rsc.SLS")
    mEs = settings.GetRequiredDouble("Steel.Es")
    mEsc = settings.GetRequiredDouble("Steel.Esc")
    mTwoLineEsc2 = settings.GetRequiredDouble("Steel.TwoLine.Esc2")
    mTwoLineEs2 = settings.GetRequiredDouble("Steel.TwoLine.Es2")
    mThreeLineEsc2 = settings.GetRequiredDouble("Steel.ThreeLine.Esc2")
    mThreeLineEs2 = settings.GetRequiredDouble("Steel.ThreeLine.Es2")

    If Not units Is Nothing Then
        mRsULS = units.InputStressToInternal(mRsULS)
        mRscULS = units.InputStressToInternal(mRscULS)
        mRsSLS = units.InputStressToInternal(mRsSLS)
        mRscSLS = units.InputStressToInternal(mRscSLS)
        mEs = units.InputStressToInternal(mEs)
        mEsc = units.InputStressToInternal(mEsc)
    End If

    Validate settings
End Sub

' Проверяет, что характеристики можно использовать для построения диаграмм.
Private Sub Validate(Optional ByVal settings As CSystemSettingsReader = Nothing)
    RequirePositive settings, "Steel.Rs.ULS", mRsULS, 3201
    RequirePositive settings, "Steel.Rsc.ULS", mRscULS, 3202
    RequirePositive settings, "Steel.Rs.SLS", mRsSLS, 3203
    RequirePositive settings, "Steel.Rsc.SLS", mRscSLS, 3204
    RequirePositive settings, "Steel.Es", mEs, 3205
    RequirePositive settings, "Steel.Esc", mEsc, 3206
    RequirePositive settings, "Steel.TwoLine.Es2", mTwoLineEs2, 3207
    RequirePositive settings, "Steel.TwoLine.Esc2", mTwoLineEsc2, 3208
    RequirePositive settings, "Steel.ThreeLine.Es2", mThreeLineEs2, 3209
    RequirePositive settings, "Steel.ThreeLine.Esc2", mThreeLineEsc2, 3210
End Sub

' Проверяет физически положительный параметр арматуры и сохраняет прежний
' машинный код. Адрес нужной сжатой/растянутой ячейки добавляется Excel-reader-ом.
Private Sub RequirePositive(ByVal settings As CSystemSettingsReader, ByVal key As String, _
        ByVal value As Double, ByVal errorCode As Long)
    If value > 0# Then Exit Sub
    Dim message As String
    If settings Is Nothing Then
        message = key & " должен быть положительным."
    Else
        message = settings.InputErrorMessage(key, "Характеристика арматуры должна быть больше нуля.", _
            "Введите положительное значение в указанных на листе единицах измерения.")
    End If
    Err.Raise vbObjectError + errorCode, "CSteelMaterialParameters", message
End Sub

Public Property Get RsULS() As Double
    RsULS = mRsULS
End Property

Public Property Get RscULS() As Double
    RscULS = mRscULS
End Property

Public Property Get RsSLS() As Double
    RsSLS = mRsSLS
End Property

Public Property Get RscSLS() As Double
    RscSLS = mRscSLS
End Property

Public Property Get Es() As Double
    Es = mEs
End Property

Public Property Get Esc() As Double
    Esc = mEsc
End Property

Public Property Get TwoLineEs2() As Double
    TwoLineEs2 = mTwoLineEs2
End Property

Public Property Get TwoLineEsc2() As Double
    TwoLineEsc2 = mTwoLineEsc2
End Property

Public Property Get ThreeLineEs2() As Double
    ThreeLineEs2 = mThreeLineEs2
End Property

Public Property Get ThreeLineEsc2() As Double
    ThreeLineEsc2 = mThreeLineEsc2
End Property

' ==========================================================================
' Создает независимый небольшой снимок параметров арматуры. Provider хранит
' свою копию и не зависит от последующей настройки исходного объекта.
Public Function Clone() As CSteelMaterialParameters
    Dim result As CSteelMaterialParameters
    Set result = New CSteelMaterialParameters
    result.Initialize mRscULS, mRsULS, mRscSLS, mRsSLS, mEsc, mEs, _
        mTwoLineEsc2, mTwoLineEs2, mThreeLineEsc2, mThreeLineEs2
    Set Clone = result
End Function

' Заполняет нормализованные параметры внутренних snapshots и тестов.
' Пользовательский ввод остается обязанностью LoadFromSettings.
Public Sub Initialize(ByVal rscULS As Double, ByVal rsULS As Double, _
        ByVal rscSLS As Double, ByVal rsSLS As Double, ByVal esc As Double, _
        ByVal es As Double, Optional ByVal twoLineEsc2 As Double = 0.025, _
        Optional ByVal twoLineEs2 As Double = 0.025, _
        Optional ByVal threeLineEsc2 As Double = 0.015, _
        Optional ByVal threeLineEs2 As Double = 0.015)
    mRscULS = rscULS
    mRsULS = rsULS
    mRscSLS = rscSLS
    mRsSLS = rsSLS
    mEsc = esc
    mEs = es
    mTwoLineEsc2 = twoLineEsc2
    mTwoLineEs2 = twoLineEs2
    mThreeLineEsc2 = threeLineEsc2
    mThreeLineEs2 = threeLineEs2
    Validate
End Sub
