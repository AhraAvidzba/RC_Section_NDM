VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUnitSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Централизованный пересчет единиц и пользовательских знаков
' ==========================================================================
' CUnitSystem - единственное место, где внешние INPUT/OUTPUT единицы и
' пользовательская система знаков переводятся во внутреннюю систему ядра и
' обратно. Solver-ы работают только во внутренних единицах и не должны знать,
' ввел пользователь силы в Н, кН или tf.
'
' Важно: OUTPUT-знак применяется при записи результатов, а не повторно при
' построении схемы из уже сохраненного Results.

Private Const TF_IN_N As Double = 9806.65
Private Const KGF_CM2_IN_MPA As Double = 0.0980665
Private Const TF_M2_IN_MPA As Double = 0.00980665

Private mInputLength As String
Private mOutputLength As String
Private mInputArea As String
Private mOutputArea As String
Private mInputForce As String
Private mOutputForce As String
Private mInputMoment As String
Private mOutputMoment As String
Private mInputStress As String
Private mOutputStress As String
Private mInputCurvature As String
Private mOutputCurvature As String
Private mSignN As String
Private mSignMx As String
Private mSignMy As String

' Устанавливает внутренние defaults, используемые до чтения Config или в автономных тестах.
Public Sub InitializeDefaults()
    mInputLength = "mm": mOutputLength = "mm"
    mInputArea = "mm2": mOutputArea = "mm2"
    mInputForce = "N": mOutputForce = "N"
    mInputMoment = "N*mm": mOutputMoment = "N*mm"
    mInputStress = "MPa": mOutputStress = "MPa"
    mInputCurvature = "1/mm": mOutputCurvature = "1/mm"
    mSignN = "Tension"
    mSignMx = "+Y tension"
    mSignMy = "+X tension"
End Sub

' Читает настройки единиц и знаков из общего reader-а Config с валидацией каждого значения.
Public Sub LoadFromSettings(ByVal settings As CSystemSettingsReader)
    InitializeDefaults
    If settings Is Nothing Then Exit Sub

    mInputLength = ValidLengthUnit(settings.GetString("Units.Length.Input", mInputLength))
    mOutputLength = ValidLengthUnit(settings.GetString("Units.Length.Output", mOutputLength))
    mInputArea = ValidAreaUnit(settings.GetString("Units.Area.Input", mInputArea))
    mOutputArea = ValidAreaUnit(settings.GetString("Units.Area.Output", mOutputArea))
    mInputForce = ValidForceUnit(settings.GetString("Units.Force.Input", mInputForce))
    mOutputForce = ValidForceUnit(settings.GetString("Units.Force.Output", mOutputForce))
    mInputMoment = ValidMomentUnit(settings.GetString("Units.Moment.Input", mInputMoment))
    mOutputMoment = ValidMomentUnit(settings.GetString("Units.Moment.Output", mOutputMoment))
    mInputStress = ValidStressUnit(settings.GetString("Units.Stress.Input", mInputStress))
    mOutputStress = ValidStressUnit(settings.GetString("Units.Stress.Output", mOutputStress))
    mInputCurvature = ValidCurvatureUnit(settings.GetString("Units.Curvature.Input", mInputCurvature))
    mOutputCurvature = ValidCurvatureUnit(settings.GetString("Units.Curvature.Output", mOutputCurvature))

    mSignN = ValidNSign(settings.GetString("Sign.N.User", mSignN))
    mSignMx = ValidMxSign(settings.GetString("Sign.Mx.User", mSignMx))
    mSignMy = ValidMySign(settings.GetString("Sign.My.User", mSignMy))
End Sub

' Переводит длину из пользовательских INPUT-единиц во внутренние мм.
Public Function InputLengthToInternal(ByVal value As Double) As Double
    InputLengthToInternal = value * LengthFactorToMm(mInputLength)
End Function

' Переводит внутреннюю длину в пользовательские OUTPUT-единицы.
Public Function InternalLengthToOutput(ByVal value As Double) As Double
    InternalLengthToOutput = value / LengthFactorToMm(mOutputLength)
End Function

' Переводит длину из OUTPUT-единиц обратно во внутренние мм.
Public Function OutputLengthToInternal(ByVal value As Double) As Double
    OutputLengthToInternal = value * LengthFactorToMm(mOutputLength)
End Function

' Переводит площадь из пользовательских INPUT-единиц во внутренние мм2.
Public Function InputAreaToInternal(ByVal value As Double) As Double
    InputAreaToInternal = value * AreaFactorToMm2(mInputArea)
End Function

' Переводит внутреннюю площадь в пользовательские OUTPUT-единицы.
Public Function InternalAreaToOutput(ByVal value As Double) As Double
    InternalAreaToOutput = value / AreaFactorToMm2(mOutputArea)
End Function

' Переводит площадь из OUTPUT-единиц обратно во внутренние мм2.
Public Function OutputAreaToInternal(ByVal value As Double) As Double
    OutputAreaToInternal = value * AreaFactorToMm2(mOutputArea)
End Function

' Переводит величину размерности L4 из INPUT-единиц во внутренние мм4.
Public Function InputFourthPowerLengthToInternal(ByVal value As Double) As Double
    InputFourthPowerLengthToInternal = value * LengthFactorToMm(mInputLength) ^ 4
End Function

' Переводит внутреннюю величину L4 в OUTPUT-единицы.
Public Function InternalFourthPowerLengthToOutput(ByVal value As Double) As Double
    InternalFourthPowerLengthToOutput = value / (LengthFactorToMm(mOutputLength) ^ 4)
End Function

' Переводит величину L4 из OUTPUT-единиц обратно во внутренние мм4.
Public Function OutputFourthPowerLengthToInternal(ByVal value As Double) As Double
    OutputFourthPowerLengthToInternal = value * LengthFactorToMm(mOutputLength) ^ 4
End Function

' Переводит продольную силу из INPUT-единиц во внутренние Н с учетом пользовательского знака N.
Public Function InputForceToInternal(ByVal value As Double) As Double
    InputForceToInternal = value * ForceFactorToN(mInputForce) * NSignFactor()
End Function

' Переводит внутреннюю силу в OUTPUT-единицы с обратным пользовательским знаком N.
Public Function InternalForceToOutput(ByVal value As Double) As Double
    InternalForceToOutput = value / ForceFactorToN(mOutputForce) * NSignFactor()
End Function

' Переводит жесткость размерности сила*длина^2 из внутренних Н*мм2
' в выбранные пользователем OUTPUT-единицы силы и длины. Знак не применяется:
' D является положительной жесткостью, а не направленной силой.
Public Function InternalForceLengthSquaredToOutput(ByVal value As Double) As Double
    InternalForceLengthSquaredToOutput = value / (ForceFactorToN(mOutputForce) * LengthFactorToMm(mOutputLength) ^ 2)
End Function

' Переводит Mx из INPUT-единиц во внутренние Н*мм с учетом пользовательского знака Mx.
Public Function InputMomentMxToInternal(ByVal value As Double) As Double
    InputMomentMxToInternal = value * MomentFactorToNmm(mInputMoment) * MxSignFactor()
End Function

' Переводит My из INPUT-единиц во внутренние Н*мм с учетом пользовательского знака My.
Public Function InputMomentMyToInternal(ByVal value As Double) As Double
    InputMomentMyToInternal = value * MomentFactorToNmm(mInputMoment) * MySignFactor()
End Function

' Переводит пользовательскую настройку вида "момент на единицу длины" во
' внутренние Н*мм/мм. Знак момента здесь не применяется: это положительный
' инженерный порог, а не расчетное усилие.
Public Function InputMomentPerLengthToInternal(ByVal value As Double) As Double
    InputMomentPerLengthToInternal = Abs(value) * MomentFactorToNmm(mInputMoment) / LengthFactorToMm(mInputLength)
End Function

' Переводит внутренний Mx в OUTPUT-единицы с учетом пользовательского знака Mx.
Public Function InternalMomentMxToOutput(ByVal value As Double) As Double
    InternalMomentMxToOutput = value / MomentFactorToNmm(mOutputMoment) * MxSignFactor()
End Function

' Переводит внутренний My в OUTPUT-единицы с учетом пользовательского знака My.
Public Function InternalMomentMyToOutput(ByVal value As Double) As Double
    InternalMomentMyToOutput = value / MomentFactorToNmm(mOutputMoment) * MySignFactor()
End Function

' Переводит модуль момента без применения знаковых соглашений.
Public Function InternalMomentMagnitudeToOutput(ByVal value As Double) As Double
    InternalMomentMagnitudeToOutput = value / MomentFactorToNmm(mOutputMoment)
End Function

' Переводит напряжение из INPUT-единиц во внутренние МПа.
Public Function InputStressToInternal(ByVal value As Double) As Double
    InputStressToInternal = value * StressFactorToMPa(mInputStress)
End Function

' Переводит внутреннее напряжение в OUTPUT-единицы.
Public Function InternalStressToOutput(ByVal value As Double) As Double
    InternalStressToOutput = value / StressFactorToMPa(mOutputStress)
End Function

' Переводит напряжение из OUTPUT-единиц обратно во внутренние МПа.
Public Function OutputStressToInternal(ByVal value As Double) As Double
    OutputStressToInternal = value * StressFactorToMPa(mOutputStress)
End Function

' Переводит кривизну из INPUT-единиц во внутренние 1/мм.
Public Function InputCurvatureToInternal(ByVal value As Double) As Double
    InputCurvatureToInternal = value * CurvatureFactorToInvMm(mInputCurvature)
End Function

' Переводит внутреннюю кривизну в OUTPUT-единицы.
Public Function InternalCurvatureToOutput(ByVal value As Double) As Double
    InternalCurvatureToOutput = value / CurvatureFactorToInvMm(mOutputCurvature)
End Function

' Переводит кривизну из OUTPUT-единиц обратно во внутренние 1/мм.
Public Function OutputCurvatureToInternal(ByVal value As Double) As Double
    OutputCurvatureToInternal = value * CurvatureFactorToInvMm(mOutputCurvature)
End Function

Public Property Get OutputLengthUnit() As String
    OutputLengthUnit = mOutputLength
End Property

Public Property Get OutputAreaUnit() As String
    OutputAreaUnit = mOutputArea
End Property

Public Property Get OutputMomentUnit() As String
    OutputMomentUnit = mOutputMoment
End Property

' Возвращает подпись INPUT-единицы для настроек размерности момент/длина.
Public Property Get InputMomentPerLengthUnit() As String
    InputMomentPerLengthUnit = mInputMoment & "/" & mInputLength
End Property

Public Property Get OutputForceUnit() As String
    OutputForceUnit = mOutputForce
End Property

' Возвращает подпись OUTPUT-единицы для жесткости D: сила*длина2.
Public Property Get OutputForceLengthSquaredUnit() As String
    OutputForceLengthSquaredUnit = mOutputForce & "*" & mOutputLength & "2"
End Property

Public Property Get OutputStressUnit() As String
    OutputStressUnit = mOutputStress
End Property

Public Property Get OutputCurvatureUnit() As String
    OutputCurvatureUnit = mOutputCurvature
End Property

Public Property Get OutputSignN() As String
    OutputSignN = mSignN
End Property

Public Property Get OutputSignMx() As String
    OutputSignMx = mSignMx
End Property

Public Property Get OutputSignMy() As String
    OutputSignMy = mSignMy
End Property

' Возвращает множитель перевода заданной единицы длины в мм.
Private Function LengthFactorToMm(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "mm": LengthFactorToMm = 1#
        Case "cm": LengthFactorToMm = 10#
        Case "m": LengthFactorToMm = 1000#
        Case Else: Err.Raise vbObjectError + 4520, "CUnitSystem", "Недопустимая единица длины: " & unitName
    End Select
End Function

' Возвращает множитель перевода заданной единицы площади в мм2.
Private Function AreaFactorToMm2(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "mm2": AreaFactorToMm2 = 1#
        Case "cm2": AreaFactorToMm2 = 100#
        Case "m2": AreaFactorToMm2 = 1000000#
        Case Else: Err.Raise vbObjectError + 4521, "CUnitSystem", "Недопустимая единица площади: " & unitName
    End Select
End Function

' Возвращает множитель перевода заданной единицы силы в Н.
Private Function ForceFactorToN(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "n": ForceFactorToN = 1#
        Case "kn": ForceFactorToN = 1000#
        Case "tf": ForceFactorToN = TF_IN_N
        Case Else: Err.Raise vbObjectError + 4522, "CUnitSystem", "Недопустимая единица силы: " & unitName
    End Select
End Function

' Возвращает множитель перевода заданной единицы момента в Н*мм.
Private Function MomentFactorToNmm(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "n*mm": MomentFactorToNmm = 1#
        Case "kn*m": MomentFactorToNmm = 1000000#
        Case "tf*m": MomentFactorToNmm = TF_IN_N * 1000#
        Case Else: Err.Raise vbObjectError + 4523, "CUnitSystem", "Недопустимая единица момента: " & unitName
    End Select
End Function

' Возвращает множитель перевода заданной единицы напряжения в МПа.
Private Function StressFactorToMPa(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "pa": StressFactorToMPa = 0.000001
        Case "kpa": StressFactorToMPa = 0.001
        Case "mpa": StressFactorToMPa = 1#
        Case "kgf/cm2": StressFactorToMPa = KGF_CM2_IN_MPA
        Case "tf/m2": StressFactorToMPa = TF_M2_IN_MPA
        Case Else: Err.Raise vbObjectError + 4524, "CUnitSystem", "Недопустимая единица напряжения: " & unitName
    End Select
End Function

' Возвращает множитель перевода заданной единицы кривизны в 1/мм.
Private Function CurvatureFactorToInvMm(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "1/mm": CurvatureFactorToInvMm = 1#
        Case "1/m": CurvatureFactorToInvMm = 0.001
        Case Else: Err.Raise vbObjectError + 4525, "CUnitSystem", "Недопустимая единица кривизны: " & unitName
    End Select
End Function

' Возвращает знаковый множитель для N: внутри ядра сжатие отрицательно.
Private Function NSignFactor() As Double
    If StrComp(mSignN, "Compression", vbTextCompare) = 0 Then NSignFactor = -1# Else NSignFactor = 1#
End Function

' Возвращает знаковый множитель для Mx согласно выбранному пользователем направлению растяжения.
Private Function MxSignFactor() As Double
    If StrComp(mSignMx, "-Y tension", vbTextCompare) = 0 Then MxSignFactor = -1# Else MxSignFactor = 1#
End Function

' Возвращает знаковый множитель для My согласно выбранному пользователем направлению растяжения.
Private Function MySignFactor() As Double
    If StrComp(mSignMy, "-X tension", vbTextCompare) = 0 Then MySignFactor = -1# Else MySignFactor = 1#
End Function

' Валидирует строку единицы длины из Config и возвращает ее без изменения регистра пользователя.
Private Function ValidLengthUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "mm", "cm", "m": ValidLengthUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4530, "CUnitSystem", "Units.Length должен быть mm, cm или m."
    End Select
End Function

' Валидирует строку единицы площади из Config.
Private Function ValidAreaUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "mm2", "cm2", "m2": ValidAreaUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4531, "CUnitSystem", "Units.Area должен быть mm2, cm2 или m2."
    End Select
End Function

' Валидирует строку единицы силы из Config.
Private Function ValidForceUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "n", "kn", "tf": ValidForceUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4532, "CUnitSystem", "Units.Force должен быть N, kN или tf."
    End Select
End Function

' Валидирует строку единицы момента из Config.
Private Function ValidMomentUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "n*mm", "kn*m", "tf*m": ValidMomentUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4533, "CUnitSystem", "Units.Moment должен быть N*mm, kN*m или tf*m."
    End Select
End Function

' Валидирует строку единицы напряжения из Config.
Private Function ValidStressUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "pa", "kpa", "mpa", "kgf/cm2", "tf/m2": ValidStressUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4534, "CUnitSystem", "Units.Stress содержит неподдерживаемое значение."
    End Select
End Function

' Валидирует строку единицы кривизны из Config.
Private Function ValidCurvatureUnit(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "1/mm", "1/m": ValidCurvatureUnit = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4535, "CUnitSystem", "Units.Curvature должен быть 1/mm или 1/m."
    End Select
End Function

' Валидирует пользовательское соглашение знака продольной силы.
Private Function ValidNSign(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "tension", "compression": ValidNSign = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4536, "CUnitSystem", "Sign.N.User должен быть Tension или Compression."
    End Select
End Function

' Валидирует пользовательское соглашение знака Mx.
Private Function ValidMxSign(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "+y tension", "-y tension": ValidMxSign = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4537, "CUnitSystem", "Sign.Mx.User должен быть +Y tension или -Y tension."
    End Select
End Function

' Валидирует пользовательское соглашение знака My.
Private Function ValidMySign(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "+x tension", "-x tension": ValidMySign = Trim$(value)
        Case Else: Err.Raise vbObjectError + 4538, "CUnitSystem", "Sign.My.User должен быть +X tension или -X tension."
    End Select
End Function

