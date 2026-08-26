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
' Настройка класса бетона здесь намеренно отсутствует: пользователь задает Rb/Rbt,
' Rb,ser/Rbt,ser, Eb/Ebt напрямую. Если когда-нибудь понадобится подбор по
' классу бетона, это должен быть отдельный слой автозаполнения Config, а не
' скрытая логика внутри расчета.

Private mRbULS As Double ' Расчетное сопротивление бетона сжатию для I ГПС, МПа.
Private mRbtULS As Double ' Расчетное сопротивление бетона растяжению для I ГПС, МПа.
Private mRbSLS As Double ' Сопротивление бетона сжатию для II ГПС, МПа.
Private mRbtSLS As Double ' Сопротивление бетона растяжению Rbt,ser для II ГПС, МПа.
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

    mRbULS = settings.GetDouble("Concrete.Rb.ULS", 15.5)
    mRbtULS = settings.GetDouble("Concrete.Rbt.ULS", 1.1)
    mRbSLS = settings.GetDouble("Concrete.Rb.SLS", 22#)
    mRbtSLS = settings.GetDouble("Concrete.Rbt.SLS", 1.8)
    mEb = settings.GetDouble("Concrete.Eb", 32500#)
    mEbt = settings.GetDouble("Concrete.Ebt", mEb)
    mEb1Red = settings.GetDouble("Concrete.Eb1Red", 0.0015)
    mEbt1Red = settings.GetDouble("Concrete.Ebt1Red", 0.00008)
    mEb0 = settings.GetDouble("Concrete.Eb0", 0.002)
    mEbt0 = settings.GetDouble("Concrete.Ebt0", 0.0001)
    mEb2 = settings.GetDouble("Concrete.Eb2", 0.0035)
    mEbt2 = settings.GetDouble("Concrete.Ebt2", 0.00015)
    If mEbt <= 0# Then mEbt = mEb

    If Not units Is Nothing Then
        mRbULS = units.InputStressToInternal(mRbULS)
        mRbtULS = units.InputStressToInternal(mRbtULS)
        mRbSLS = units.InputStressToInternal(mRbSLS)
        mRbtSLS = units.InputStressToInternal(mRbtSLS)
        mEb = units.InputStressToInternal(mEb)
        mEbt = units.InputStressToInternal(mEbt)
    End If

    Validate
End Sub

' Заполняет параметры напрямую без чтения Config.
' Метод нужен тестам и служебным сценариям, где важно проверить построитель
' диаграмм из известных чисел, не создавая временный Excel-диапазон.
Public Sub Initialize(ByVal rbULS As Double, ByVal rbtULS As Double, _
        ByVal rbSLS As Double, ByVal rbtSLS As Double, ByVal eb As Double, _
        Optional ByVal ebt As Double = 0#, _
        Optional ByVal eb1Red As Double = 0.0015, Optional ByVal ebt1Red As Double = 0.00008, _
        Optional ByVal eb0 As Double = 0.002, Optional ByVal ebt0 As Double = 0.0001, _
        Optional ByVal eb2 As Double = 0.0035, Optional ByVal ebt2 As Double = 0.00015)
    mRbULS = rbULS
    mRbtULS = rbtULS
    mRbSLS = rbSLS
    mRbtSLS = rbtSLS
    mEb = eb
    mEbt = ebt
    mEb1Red = eb1Red
    mEbt1Red = ebt1Red
    mEb0 = eb0
    mEbt0 = ebt0
    mEb2 = eb2
    mEbt2 = ebt2
    If mEbt <= 0# Then mEbt = mEb
    Validate
End Sub

' Проверяет, что характеристики можно использовать для диаграмм.
Private Sub Validate()
    If mRbULS <= 0# Then Err.Raise vbObjectError + 3191, "CConcreteMaterialParameters", "Concrete.Rb.ULS должен быть положительным."
    If mRbtULS <= 0# Then Err.Raise vbObjectError + 3192, "CConcreteMaterialParameters", "Concrete.Rbt.ULS должен быть положительным."
    If mRbSLS <= 0# Then Err.Raise vbObjectError + 3193, "CConcreteMaterialParameters", "Concrete.Rb.SLS должен быть положительным."
    If mRbtSLS <= 0# Then Err.Raise vbObjectError + 3194, "CConcreteMaterialParameters", "Concrete.Rbt.SLS должен быть положительным."
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
