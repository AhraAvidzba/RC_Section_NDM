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
' Класс хранит исходные характеристики ненапрягаемой арматуры. Расчетное
' ядро не читает их напрямую: builders строят по ним диаграммы, а solver
' получает уже готовый CSteelDiagramMaterial.
'
' Rsc,ser и Esc оставлены как программные пользовательские параметры для
' симметрии сжатой и растянутой ветви и для исследований. В СП 63 отдельные
' обозначения Rsc,ser и Esc для обычной арматуры не вводятся, поэтому
' нормативные defaults: Rsc,ser = Rs,ser, Esc = Es.

Private mRsULS As Double ' Расчетное сопротивление растяжению для I ГПС, МПа.
Private mRscULS As Double ' Расчетное сопротивление сжатию для I ГПС, МПа.
Private mRsSLS As Double ' Сопротивление растяжению для II ГПС, МПа.
Private mRscSLS As Double ' Программное сопротивление сжатию для II ГПС, default RsSLS, МПа.
Private mEs As Double ' Модуль арматуры при растяжении, МПа.
Private mEsc As Double ' Программный модуль при сжатии, default Es, МПа.
Private mRebarProfile As String ' Ribbed/Smooth для коэффициента phi2 в расчете трещин.

' Читает параметры из Config и сразу переводит напряжения/модули в МПа.
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3200, "CSteelMaterialParameters", "Настройки арматуры не переданы."

    mRsULS = settings.GetDouble("Steel.Rs.ULS", 350#)
    mRscULS = settings.GetDouble("Steel.Rsc.ULS", mRsULS)
    mRsSLS = settings.GetDouble("Steel.Rs.SLS", mRsULS)
    mRscSLS = settings.GetDouble("Steel.Rsc.SLS", mRsSLS)
    mEs = settings.GetDouble("Steel.Es", 200000#)
    mEsc = settings.GetDouble("Steel.Esc", mEs)
    If mRscSLS <= 0# Then mRscSLS = mRsSLS
    If mEsc <= 0# Then mEsc = mEs
    mRebarProfile = settings.GetString("Steel.RebarProfile", "Ribbed")

    If Not units Is Nothing Then
        mRsULS = units.InputStressToInternal(mRsULS)
        mRscULS = units.InputStressToInternal(mRscULS)
        mRsSLS = units.InputStressToInternal(mRsSLS)
        mRscSLS = units.InputStressToInternal(mRscSLS)
        mEs = units.InputStressToInternal(mEs)
        mEsc = units.InputStressToInternal(mEsc)
    End If

    NormalizeProfile
    Validate
End Sub

' Заполняет параметры напрямую без чтения Config.
' Используется в тестах построителей диаграмм, чтобы сверять нормативные точки
' на фиксированном наборе характеристик без зависимости от состояния книги.
Public Sub Initialize(ByVal rscULS As Double, ByVal rsULS As Double, _
        ByVal rscSLS As Double, ByVal rsSLS As Double, ByVal esc As Double, _
        ByVal es As Double, Optional ByVal rebarProfile As String = "Ribbed")
    mRscULS = rscULS
    mRsULS = rsULS
    mRscSLS = rscSLS
    mRsSLS = rsSLS
    mEsc = esc
    mEs = es
    If mRscSLS <= 0# Then mRscSLS = mRsSLS
    If mEsc <= 0# Then mEsc = mEs
    mRebarProfile = rebarProfile
    NormalizeProfile
    Validate
End Sub

' Нормализует пользовательскую подпись профиля арматуры.
Private Sub NormalizeProfile()
    Select Case LCase$(Trim$(mRebarProfile))
        Case "smooth", "гладкая"
            mRebarProfile = "Smooth"
        Case Else
            mRebarProfile = "Ribbed"
    End Select
End Sub

' Проверяет, что характеристики можно использовать для построения диаграмм.
Private Sub Validate()
    If mRsULS <= 0# Then Err.Raise vbObjectError + 3201, "CSteelMaterialParameters", "Steel.Rs.ULS должен быть положительным."
    If mRscULS <= 0# Then Err.Raise vbObjectError + 3202, "CSteelMaterialParameters", "Steel.Rsc.ULS должен быть положительным."
    If mRsSLS <= 0# Then Err.Raise vbObjectError + 3203, "CSteelMaterialParameters", "Steel.Rs.SLS должен быть положительным."
    If mRscSLS <= 0# Then Err.Raise vbObjectError + 3204, "CSteelMaterialParameters", "Steel.Rsc.SLS должен быть положительным."
    If mEs <= 0# Then Err.Raise vbObjectError + 3205, "CSteelMaterialParameters", "Steel.Es должен быть положительным."
    If mEsc <= 0# Then Err.Raise vbObjectError + 3206, "CSteelMaterialParameters", "Steel.Esc должен быть положительным."
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

Public Property Get RebarProfile() As String
    RebarProfile = mRebarProfile
End Property
