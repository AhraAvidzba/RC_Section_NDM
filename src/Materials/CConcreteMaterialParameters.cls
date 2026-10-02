VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConcreteMaterialParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Параметры бетона для построения диаграмм
' ==========================================================================
' Класс хранит пользовательские характеристики бетона, которые являются
' исходными данными для программного построения TwoLine/ThreeLine диаграмм
' по СП 63. Он не строит диаграммы сам и не участвует в solver-е.
'
' Пользователь задает сопротивления и модули напрямую в Config. Класс не
' подбирает их по классу бетона и не заменяет введенные характеристики.

Private mRbULS As Double ' Расчетное сопротивление бетона сжатию для I ГПС, МПа.
Private mRbtULS As Double ' Расчетное сопротивление бетона растяжению для I ГПС, МПа.
Private mRbSLS As Double ' Сопротивление бетона сжатию для II ГПС, МПа.
Private mRbtSLS As Double ' Сопротивление бетона растяжению Rbt,ser для II ГПС, МПа.
Private mRbMc2 As Double ' Допустимое сжатие бетона для проверки продольных трещин по СП 35, МПа.
Private mEb As Double ' Начальный модуль бетона при сжатии, МПа.
Private mEbt As Double ' Начальный модуль бетона при растяжении, МПа; default по СП равен Eb.
Private mEb1Red As Double ' Пользовательская деформация eps_b1,red для двухлинейной сжатой ветви.
Private mEbt1Red As Double ' Пользовательская деформация eps_bt1,red для двухлинейной растянутой ветви.
Private mEb0 As Double ' Пользовательская деформация eps_b0 для трехлинейной сжатой ветви.
Private mEbt0 As Double ' Пользовательская деформация eps_bt0 для трехлинейной растянутой ветви.
Private mEb2 As Double ' Предельная деформация eps_b2 для TwoLine/ThreeLine при сжатии.
Private mEbt2 As Double ' Предельная деформация eps_bt2 для TwoLine/ThreeLine при растяжении.

' Читает параметры из Config и сразу переводит напряжения в МПа.
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3190, "CConcreteMaterialParameters", "Настройки бетона не переданы."

    mRbULS = settings.GetRequiredDouble("Concrete.Rb.ULS")
    mRbtULS = settings.GetRequiredDouble("Concrete.Rbt.ULS")
    mRbSLS = settings.GetRequiredDouble("Concrete.Rb.SLS")
    mRbtSLS = settings.GetRequiredDouble("Concrete.Rbt.SLS")
    mRbMc2 = settings.GetRequiredDouble("Concrete.Rb.mc2")
    mEb = settings.GetRequiredDouble("Concrete.Eb")
    mEbt = settings.GetRequiredDouble("Concrete.Ebt")
    mEb1Red = settings.GetRequiredDouble("Concrete.Eb1Red")
    mEbt1Red = settings.GetRequiredDouble("Concrete.Ebt1Red")
    mEb0 = settings.GetRequiredDouble("Concrete.Eb0")
    mEbt0 = settings.GetRequiredDouble("Concrete.Ebt0")
    mEb2 = settings.GetRequiredDouble("Concrete.Eb2")
    mEbt2 = settings.GetRequiredDouble("Concrete.Ebt2")

    If Not units Is Nothing Then
        mRbULS = units.InputStressToInternal(mRbULS)
        mRbtULS = units.InputStressToInternal(mRbtULS)
        mRbSLS = units.InputStressToInternal(mRbSLS)
        mRbtSLS = units.InputStressToInternal(mRbtSLS)
        mRbMc2 = units.InputStressToInternal(mRbMc2)
        mEb = units.InputStressToInternal(mEb)
        mEbt = units.InputStressToInternal(mEbt)
    End If

    Validate
End Sub

' Проверяет, что характеристики можно использовать для диаграмм.
Private Sub Validate()
    If mRbULS <= 0# Then Err.Raise vbObjectError + 3191, "CConcreteMaterialParameters", "Concrete.Rb.ULS должен быть положительным."
    If mRbtULS <= 0# Then Err.Raise vbObjectError + 3192, "CConcreteMaterialParameters", "Concrete.Rbt.ULS должен быть положительным."
    If mRbSLS <= 0# Then Err.Raise vbObjectError + 3193, "CConcreteMaterialParameters", "Concrete.Rb.SLS должен быть положительным."
    If mRbtSLS <= 0# Then Err.Raise vbObjectError + 3194, "CConcreteMaterialParameters", "Concrete.Rbt.SLS должен быть положительным."
    If mRbMc2 <= 0# Then Err.Raise vbObjectError + 3203, "CConcreteMaterialParameters", "Concrete.Rb.mc2 должен быть положительным."
    If mEb <= 0# Then Err.Raise vbObjectError + 3195, "CConcreteMaterialParameters", "Concrete.Eb должен быть положительным."
    If mEbt <= 0# Then Err.Raise vbObjectError + 3196, "CConcreteMaterialParameters", "Concrete.Ebt должен быть положительным."
    If mEb1Red <= 0# Then Err.Raise vbObjectError + 3197, "CConcreteMaterialParameters", "Concrete.Eb1Red должен быть положительным."
    If mEbt1Red <= 0# Then Err.Raise vbObjectError + 3198, "CConcreteMaterialParameters", "Concrete.Ebt1Red должен быть положительным."
    If mEb0 <= 0# Then Err.Raise vbObjectError + 3199, "CConcreteMaterialParameters", "Concrete.Eb0 должен быть положительным."
    If mEbt0 <= 0# Then Err.Raise vbObjectError + 3200, "CConcreteMaterialParameters", "Concrete.Ebt0 должен быть положительным."
    If mEb2 <= 0# Then Err.Raise vbObjectError + 3201, "CConcreteMaterialParameters", "Concrete.Eb2 должен быть положительным."
    If mEbt2 <= 0# Then Err.Raise vbObjectError + 3202, "CConcreteMaterialParameters", "Concrete.Ebt2 должен быть положительным."
End Sub

Public Property Get RbULS() As Double
    RbULS = mRbULS
End Property

Public Property Get RbtULS() As Double
    RbtULS = mRbtULS
End Property

Public Property Get RbSLS() As Double
    RbSLS = mRbSLS
End Property

Public Property Get RbtSLS() As Double
    RbtSLS = mRbtSLS
End Property

Public Property Get RbMc2() As Double
    RbMc2 = mRbMc2
End Property

Public Property Get Eb() As Double
    Eb = mEb
End Property

Public Property Get Ebt() As Double
    Ebt = mEbt
End Property

Public Property Get Eb1Red() As Double
    Eb1Red = mEb1Red
End Property

Public Property Get Ebt1Red() As Double
    Ebt1Red = mEbt1Red
End Property

Public Property Get Eb0() As Double
    Eb0 = mEb0
End Property

Public Property Get Ebt0() As Double
    Ebt0 = mEbt0
End Property

Public Property Get Eb2() As Double
    Eb2 = mEb2
End Property

Public Property Get Ebt2() As Double
    Ebt2 = mEbt2
End Property

' ==========================================================================
' Создает независимый небольшой снимок параметров материала. Provider владеет
' своей копией; изменение исходного объекта не меняет уже построенные диаграммы.
Public Function Clone() As CConcreteMaterialParameters
    Dim result As CConcreteMaterialParameters
    Set result = New CConcreteMaterialParameters
    result.Initialize mRbULS, mRbtULS, mRbSLS, mRbtSLS, mEb, mEbt, _
        mEb1Red, mEbt1Red, mEb0, mEbt0, mEb2, mEbt2, mRbMc2
    Set Clone = result
End Function

' Заполняет параметры нормализованными числами для внутренних snapshots и
' тестовых сценариев. Пользовательский ввод читается через LoadFromSettings.
Public Sub Initialize(ByVal rbULS As Double, ByVal rbtULS As Double, _
        ByVal rbSLS As Double, ByVal rbtSLS As Double, ByVal eb As Double, _
        Optional ByVal ebt As Double = 0#, _
        Optional ByVal eb1Red As Double = 0.0015, Optional ByVal ebt1Red As Double = 0.00008, _
        Optional ByVal eb0 As Double = 0.002, Optional ByVal ebt0 As Double = 0.0001, _
        Optional ByVal eb2 As Double = 0.0035, Optional ByVal ebt2 As Double = 0.00015, _
        Optional ByVal rbMc2 As Double = 0#)
    mRbULS = rbULS
    mRbtULS = rbtULS
    mRbSLS = rbSLS
    mRbtSLS = rbtSLS
    mRbMc2 = rbMc2
    mEb = eb
    mEbt = ebt
    mEb1Red = eb1Red
    mEbt1Red = ebt1Red
    mEb0 = eb0
    mEbt0 = ebt0
    mEb2 = eb2
    mEbt2 = ebt2
    Validate
End Sub
