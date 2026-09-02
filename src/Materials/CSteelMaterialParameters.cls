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
' Коэффициенты раскрытия трещин здесь не читаются: phi2 теперь задается
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

    Validate
End Sub

' Проверяет, что характеристики можно использовать для построения диаграмм.
Private Sub Validate()
    If mRsULS <= 0# Then Err.Raise vbObjectError + 3201, "CSteelMaterialParameters", "Steel.Rs.ULS должен быть положительным."
    If mRscULS <= 0# Then Err.Raise vbObjectError + 3202, "CSteelMaterialParameters", "Steel.Rsc.ULS должен быть положительным."
    If mRsSLS <= 0# Then Err.Raise vbObjectError + 3203, "CSteelMaterialParameters", "Steel.Rs.SLS должен быть положительным."
    If mRscSLS <= 0# Then Err.Raise vbObjectError + 3204, "CSteelMaterialParameters", "Steel.Rsc.SLS должен быть положительным."
    If mEs <= 0# Then Err.Raise vbObjectError + 3205, "CSteelMaterialParameters", "Steel.Es должен быть положительным."
    If mEsc <= 0# Then Err.Raise vbObjectError + 3206, "CSteelMaterialParameters", "Steel.Esc должен быть положительным."
    If mTwoLineEs2 <= 0# Then Err.Raise vbObjectError + 3207, "CSteelMaterialParameters", "Steel.TwoLine.Es2 должен быть положительным."
    If mTwoLineEsc2 <= 0# Then Err.Raise vbObjectError + 3208, "CSteelMaterialParameters", "Steel.TwoLine.Esc2 должен быть положительным."
    If mThreeLineEs2 <= 0# Then Err.Raise vbObjectError + 3209, "CSteelMaterialParameters", "Steel.ThreeLine.Es2 должен быть положительным."
    If mThreeLineEsc2 <= 0# Then Err.Raise vbObjectError + 3210, "CSteelMaterialParameters", "Steel.ThreeLine.Esc2 должен быть положительным."
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
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Тесты передают параметры напрямую, чтобы проверять построение диаграмм без
' временного листа Config. В пользовательском сценарии значения должны идти
' через LoadFromSettings и единый CMaterialModelProvider.

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
