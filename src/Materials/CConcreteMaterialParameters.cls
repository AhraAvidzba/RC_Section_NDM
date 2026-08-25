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

' Читает параметры из Config и сразу переводит напряжения в МПа.
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3190, "CConcreteMaterialParameters", "Настройки бетона не переданы."

    mRbULS = settings.GetDouble("Concrete.Rb.ULS", 15.5)
    mRbtULS = settings.GetDouble("Concrete.Rbt.ULS", 1.1)
    mRbSLS = settings.GetDouble("Concrete.Rb.SLS", 22#)
    mRbtSLS = settings.GetDouble("Concrete.Rbt.SLS", 1.8)
    mEb = settings.GetDouble("Concrete.Eb", 32500#)
    mEbt = settings.GetDouble("Concrete.Ebt", mEb)
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
        Optional ByVal ebt As Double = 0#)
    mRbULS = rbULS
    mRbtULS = rbtULS
    mRbSLS = rbSLS
    mRbtSLS = rbtSLS
    mEb = eb
    mEbt = ebt
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
