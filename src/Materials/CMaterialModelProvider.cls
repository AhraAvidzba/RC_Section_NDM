VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialModelProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ЕДИНАЯ ТОЧКА ПОЛУЧЕНИЯ РАСЧЕТНЫХ ДИАГРАММ МАТЕРИАЛОВ
' ==========================================================================
' Provider получает исходные параметры бетона и арматуры, читает режимы
' расчетов с листа Config и строит готовые CMaterialDiagram для физических
' целей Strength, Mcrc и CrackedNDS. В этом классе остается выбор I/II ГПС,
' сопротивлений R/Rser, типа TwoLine/ThreeLine и режима растянутого бетона.
'
' Прямой расчет НДС использует отдельную цель StateSolution. Для нее provider
' берет одну из физических баз, Strength или CrackedNDS, и при включенной
' настройке Solver.DirectState.DiagramExtension возвращает временную диаграмму
' с малой технической ветвью за физическим пределом. CSectionSolver об этом не
' знает: он по-прежнему получает готовые диаграммы и вызывает общий контракт
' GetStress/GetTangentModulus/GetSecantModulus.

Private Const SLOT_STRENGTH As Long = 1 ' Массивный слот физической диаграммы I ГПС.
Private Const SLOT_MCRC As Long = 2 ' Массивный слот физической диаграммы образования трещины.
Private Const SLOT_CRACKED_NDS As Long = 3 ' Массивный слот физической диаграммы раскрытой трещины.
Private Const STATE_EXTENSION_E_FACTOR As Double = 0.01 ' Жесткость extension: 1% от начального модуля E.
Private Const STATE_EXTENSION_STRAIN_LIMIT As Double = 10# ' Техническая дальняя точка extension; не является eps_ult.
Private Const EXTENSION_TOLERANCE As Double = 0.000000000001 ' Допуск, чтобы не добавлять почти совпадающую extension-точку.

Private mConcreteParameters As CConcreteMaterialParameters ' Исходные параметры бетона из Config.
Private mSteelParameters As CSteelMaterialParameters ' Исходные параметры арматуры из Config.
Private mConcreteDiagram(1 To 3) As CMaterialDiagram ' Физические бетонные диаграммы без numerical extension.
Private mSteelDiagram(1 To 3) As CMaterialDiagram ' Физические диаграммы арматуры без numerical extension.
Private mPurposeName(1 To 3) As String ' Каноническое имя физической цели для Results/diagnostics.
Private mLimitStateGroup(1 To 3) As String ' I или II ГПС, фактически использованная provider-ом.
Private mConcreteDiagramType(1 To 3) As String ' TwoLine или ThreeLine для бетона.
Private mConcreteTensionBehavior(1 To 3) As String ' Ignore или UseDiagram для растянутого бетона.
Private mSteelDiagramType(1 To 3) As String ' TwoLine или ThreeLine для арматуры.
Private mConcreteCompressionLimit(1 To 3) As Double ' Физический предел сжатой деформации бетона.
Private mConcreteTensionLimit(1 To 3) As Double ' Физический предел растянутой деформации бетона, если ветвь включена.
Private mConcreteTensionLimitEnabled(1 To 3) As Boolean ' True, если растянутый бетон имеет физический eps-предел.
Private mSteelCompressionLimit(1 To 3) As Double ' Физический предел сжатой деформации арматуры.
Private mSteelTensionLimit(1 To 3) As Double ' Физический предел растянутой деформации арматуры.
Private mConcreteCompressionResistance(1 To 3) As Double ' Rb/Rb,ser выбранной ГПС, МПа.
Private mConcreteTensionResistance(1 To 3) As Double ' Rbt/Rbt,ser выбранной ГПС, МПа.
Private mSteelCompressionResistance(1 To 3) As Double ' Rsc/Rsc,ser выбранной ГПС, МПа.
Private mSteelTensionResistance(1 To 3) As Double ' Rs/Rs,ser выбранной ГПС, МПа.
Private mConcreteElasticModulus(1 To 3) As Double ' Eb выбранной бетонной модели, МПа.
Private mSteelElasticModulus(1 To 3) As Double ' Es выбранной модели арматуры, МПа.
Private mDirectStateDiagramExtension As Boolean ' Разрешает numerical extension только для StateSolution.
Private mInitialized As Boolean ' Защита от использования provider-а до чтения параметров.

' Читает Config, создает параметры материалов и строит физические диаграммы.
Public Sub Initialize(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3240, "CMaterialModelProvider", "Настройки материалов не переданы."

    Set mConcreteParameters = New CConcreteMaterialParameters
    mConcreteParameters.LoadFromSettings settings, units

    Set mSteelParameters = New CSteelMaterialParameters
    mSteelParameters.LoadFromSettings settings, units

    mDirectStateDiagramExtension = settings.GetBoolean("Solver.DirectState.DiagramExtension", True)

    BuildPurpose cpStrength, "I", _
        settings.GetString("Diagram.Strength.Concrete", "TwoLine"), _
        settings.GetString("Diagram.Strength.ConcreteTension", "Ignore"), _
        settings.GetString("Diagram.Strength.Steel", "TwoLine")

    BuildPurpose cpMcrc, "II", _
        settings.GetString("Diagram.Mcrc.Concrete", "ThreeLine"), _
        "UseDiagram", _
        settings.GetString("Diagram.Mcrc.Steel", "TwoLine")

    BuildPurpose cpCrackedNDS, "II", _
        settings.GetString("Diagram.CrackedNDS.Concrete", "TwoLine"), _
        "Ignore", _
        settings.GetString("Diagram.CrackedNDS.Steel", "TwoLine")

    mInitialized = True
End Sub

' Строит полный физический материальный набор одной расчетной цели.
Private Sub BuildPurpose(ByVal purpose As ECalculationPurpose, ByVal limitStateGroup As String, _
        ByVal concreteDiagramType As String, ByVal concreteTensionBehavior As String, _
        ByVal steelDiagramType As String)
    Dim slot As Long
    slot = PhysicalPurposeSlot(purpose)

    mPurposeName(slot) = PurposeToText(purpose)
    mLimitStateGroup(slot) = NormalizeLimitStateGroup(limitStateGroup)
    mConcreteDiagramType(slot) = NormalizeDiagramType(concreteDiagramType)
    mConcreteTensionBehavior(slot) = NormalizeTensionBehavior(concreteTensionBehavior)
    mSteelDiagramType(slot) = NormalizeDiagramType(steelDiagramType)

    ConcreteStrengths mLimitStateGroup(slot), mConcreteCompressionResistance(slot), mConcreteTensionResistance(slot)
    SteelStrengths mLimitStateGroup(slot), mSteelCompressionResistance(slot), mSteelTensionResistance(slot)
    mConcreteElasticModulus(slot) = mConcreteParameters.Eb
    mSteelElasticModulus(slot) = mSteelParameters.Es

    Set mConcreteDiagram(slot) = BuildConcreteDiagram(mLimitStateGroup(slot), _
        mConcreteDiagramType(slot), mConcreteTensionBehavior(slot))
    Set mSteelDiagram(slot) = BuildSteelDiagram(mLimitStateGroup(slot), mSteelDiagramType(slot))

    mConcreteCompressionLimit(slot) = mConcreteDiagram(slot).UltimateCompressionStrain
    mConcreteTensionLimit(slot) = mConcreteDiagram(slot).UltimateTensionStrain
    mConcreteTensionLimitEnabled(slot) = (mConcreteDiagram(slot).HasPhysicalTensionLimit And _
        StrComp(mConcreteTensionBehavior(slot), "UseDiagram", vbTextCompare) = 0)
    mSteelCompressionLimit(slot) = mSteelDiagram(slot).UltimateCompressionStrain
    mSteelTensionLimit(slot) = mSteelDiagram(slot).UltimateTensionStrain
End Sub

' Выбирает формулу бетонной диаграммы.
Private Function BuildConcreteDiagram(ByVal limitStateGroup As String, ByVal diagramType As String, _
        ByVal tensionBehavior As String) As CMaterialDiagram
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        Set BuildConcreteDiagram = BuildConcreteThreeLine(limitStateGroup, tensionBehavior)
    Else
        Set BuildConcreteDiagram = BuildConcreteTwoLine(limitStateGroup, tensionBehavior)
    End If
End Function

' Строит двухлинейную бетонную диаграмму. Параметры eps задает Config,
' provider добавляет знак, сопротивление выбранной ГПС и опциональную
' растянутую ветвь.
Private Function BuildConcreteTwoLine(ByVal limitStateGroup As String, ByVal tensionBehavior As String) As CMaterialDiagram
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths limitStateGroup, rb, rbt

    Dim count As Long
    If StrComp(tensionBehavior, "UseDiagram", vbTextCompare) = 0 Then
        count = 5
    Else
        count = 3
    End If

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To count)
    ReDim sig(1 To count)

    eps(1) = -mConcreteParameters.Eb2: sig(1) = -rb
    eps(2) = -mConcreteParameters.Eb1Red: sig(2) = -rb
    eps(3) = 0#: sig(3) = 0#
    If count = 5 Then
        eps(4) = mConcreteParameters.Ebt1Red: sig(4) = rbt
        eps(5) = mConcreteParameters.Ebt2: sig(5) = rbt
    End If

    Set BuildConcreteTwoLine = DiagramFromArrays(eps, sig, count)
End Function

' Строит трехлинейную бетонную диаграмму. Средняя точка 0.6*R/E
' вычисляется здесь, чтобы расчетное ядро не знало формулы построения.
Private Function BuildConcreteThreeLine(ByVal limitStateGroup As String, ByVal tensionBehavior As String) As CMaterialDiagram
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths limitStateGroup, rb, rbt

    Dim count As Long
    If StrComp(tensionBehavior, "UseDiagram", vbTextCompare) = 0 Then
        count = 7
    Else
        count = 4
    End If

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To count)
    ReDim sig(1 To count)

    eps(1) = -mConcreteParameters.Eb2: sig(1) = -rb
    eps(2) = -mConcreteParameters.Eb0: sig(2) = -rb
    eps(3) = -0.6 * rb / mConcreteParameters.Eb: sig(3) = -0.6 * rb
    eps(4) = 0#: sig(4) = 0#
    If count = 7 Then
        eps(5) = 0.6 * rbt / mConcreteParameters.Ebt: sig(5) = 0.6 * rbt
        eps(6) = mConcreteParameters.Ebt0: sig(6) = rbt
        eps(7) = mConcreteParameters.Ebt2: sig(7) = rbt
    End If

    Set BuildConcreteThreeLine = DiagramFromArrays(eps, sig, count)
End Function

' Выбирает формулу диаграммы арматуры.
Private Function BuildSteelDiagram(ByVal limitStateGroup As String, ByVal diagramType As String) As CMaterialDiagram
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        Set BuildSteelDiagram = BuildSteelThreeLine(limitStateGroup)
    Else
        Set BuildSteelDiagram = BuildSteelTwoLine(limitStateGroup)
    End If
End Function

' Строит двухлинейную диаграмму арматуры.
Private Function BuildSteelTwoLine(ByVal limitStateGroup As String) As CMaterialDiagram
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths limitStateGroup, rsc, rs

    Dim eps(1 To 5) As Double
    Dim sig(1 To 5) As Double
    eps(1) = -mSteelParameters.TwoLineEsc2: sig(1) = -rsc
    eps(2) = -rsc / mSteelParameters.Esc: sig(2) = -rsc
    eps(3) = 0#: sig(3) = 0#
    eps(4) = rs / mSteelParameters.Es: sig(4) = rs
    eps(5) = mSteelParameters.TwoLineEs2: sig(5) = rs

    Set BuildSteelTwoLine = DiagramFromArrays(eps, sig, 5)
End Function

' Строит трехлинейную диаграмму арматуры с расчетными точками перелома.
Private Function BuildSteelThreeLine(ByVal limitStateGroup As String) As CMaterialDiagram
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths limitStateGroup, rsc, rs

    Dim epsSc0 As Double
    Dim epsSc1 As Double
    Dim epsScPl As Double
    Dim epsS0 As Double
    Dim epsS1 As Double
    Dim epsSPl As Double

    epsSc0 = rsc / mSteelParameters.Esc + 0.002
    epsSc1 = 0.9 * rsc / mSteelParameters.Esc
    epsScPl = 2# * epsSc0 - epsSc1
    epsS0 = rs / mSteelParameters.Es + 0.002
    epsS1 = 0.9 * rs / mSteelParameters.Es
    epsSPl = 2# * epsS0 - epsS1

    Dim eps(1 To 9) As Double
    Dim sig(1 To 9) As Double
    eps(1) = -mSteelParameters.ThreeLineEsc2: sig(1) = -1.1 * rsc
    eps(2) = -epsScPl: sig(2) = -1.1 * rsc
    eps(3) = -epsSc0: sig(3) = -rsc
    eps(4) = -epsSc1: sig(4) = -0.9 * rsc
    eps(5) = 0#: sig(5) = 0#
    eps(6) = epsS1: sig(6) = 0.9 * rs
    eps(7) = epsS0: sig(7) = rs
    eps(8) = epsSPl: sig(8) = 1.1 * rs
    eps(9) = mSteelParameters.ThreeLineEs2: sig(9) = 1.1 * rs

    Set BuildSteelThreeLine = DiagramFromArrays(eps, sig, 9)
End Function

' Создает CMaterialDiagram из массивов физических точек.
Private Function DiagramFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long) As CMaterialDiagram
    Dim diagram As CMaterialDiagram
    Set diagram = New CMaterialDiagram
    diagram.InitializeFromArrays strains, stresses, pointCount
    Set DiagramFromArrays = diagram
End Function

' Создает копию физической диаграммы с дополнительными техническими точками
' для прямого StateSolution. Физические eps_ult передаются в CMaterialDiagram
' отдельно и поэтому не меняются.
Private Function DiagramWithExtension(ByVal baseDiagram As CMaterialDiagram, _
        ByVal compressionE As Double, ByVal tensionE As Double, _
        ByVal extendCompression As Boolean, ByVal extendTension As Boolean) As CMaterialDiagram
    Dim compressionStart As Double
    Dim tensionStart As Double
    compressionStart = baseDiagram.PhysicalCompressionStrain
    tensionStart = baseDiagram.PhysicalTensionStrain

    Dim extraCompression As Boolean
    Dim extraTension As Boolean
    extraCompression = extendCompression And _
        (-STATE_EXTENSION_STRAIN_LIMIT < baseDiagram.PhysicalCompressionStrain - EXTENSION_TOLERANCE)
    extraTension = extendTension And baseDiagram.HasPhysicalTensionLimit And _
        (STATE_EXTENSION_STRAIN_LIMIT > baseDiagram.PhysicalTensionStrain + EXTENSION_TOLERANCE)

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To baseDiagram.PointCount + 2)
    ReDim sig(1 To baseDiagram.PointCount + 2)

    Dim i As Long
    Dim targetIndex As Long
    targetIndex = 1
    If extraCompression Then
        eps(targetIndex) = -STATE_EXTENSION_STRAIN_LIMIT
        sig(targetIndex) = baseDiagram.GetStress(compressionStart) + _
            (eps(targetIndex) - compressionStart) * STATE_EXTENSION_E_FACTOR * compressionE
        targetIndex = targetIndex + 1
    End If

    For i = 1 To baseDiagram.PointCount
        eps(targetIndex) = baseDiagram.PointStrain(i)
        sig(targetIndex) = baseDiagram.PointStress(i)
        targetIndex = targetIndex + 1
    Next i

    If extraTension Then
        eps(targetIndex) = STATE_EXTENSION_STRAIN_LIMIT
        sig(targetIndex) = baseDiagram.GetStress(tensionStart) + _
            (eps(targetIndex) - tensionStart) * STATE_EXTENSION_E_FACTOR * tensionE
        targetIndex = targetIndex + 1
    End If

    Dim count As Long
    count = targetIndex - 1

    Dim diagram As CMaterialDiagram
    Set diagram = New CMaterialDiagram
    diagram.InitializeFromArrays eps, sig, count, _
        baseDiagram.PhysicalCompressionStrain, baseDiagram.PhysicalTensionStrain, _
        extraCompression, extraTension, compressionStart, tensionStart
    Set DiagramWithExtension = diagram
End Function

' Возвращает бетонные сопротивления нужной ГПС.
Private Sub ConcreteStrengths(ByVal limitStateGroup As String, ByRef rb As Double, ByRef rbt As Double)
    If StrComp(limitStateGroup, "II", vbTextCompare) = 0 Then
        rb = mConcreteParameters.RbSLS
        rbt = mConcreteParameters.RbtSLS
    Else
        rb = mConcreteParameters.RbULS
        rbt = mConcreteParameters.RbtULS
    End If
End Sub

' Возвращает сопротивления арматуры нужной ГПС.
Private Sub SteelStrengths(ByVal limitStateGroup As String, ByRef rsc As Double, ByRef rs As Double)
    If StrComp(limitStateGroup, "II", vbTextCompare) = 0 Then
        rsc = mSteelParameters.RscSLS
        rs = mSteelParameters.RsSLS
    Else
        rsc = mSteelParameters.RscULS
        rs = mSteelParameters.RsULS
    End If
End Sub

' Возвращает индекс внутренних массивов для физической цели.
Private Function PhysicalPurposeSlot(ByVal purpose As ECalculationPurpose) As Long
    Select Case purpose
        Case cpStrength
            PhysicalPurposeSlot = SLOT_STRENGTH
        Case cpMcrc
            PhysicalPurposeSlot = SLOT_MCRC
        Case cpCrackedNDS
            PhysicalPurposeSlot = SLOT_CRACKED_NDS
        Case Else
            Err.Raise vbObjectError + 3241, "CMaterialModelProvider", _
                "Для material provider требуется физическая расчетная цель, а не " & PurposeToText(purpose) & "."
    End Select
End Function

' Проверяет, что StateSolution питается допустимой физической базой.
Private Sub ValidateStateBasePurpose(ByVal basePurpose As ECalculationPurpose)
    If Not IsStateBasePurpose(basePurpose) Then
        Err.Raise vbObjectError + 3246, "CMaterialModelProvider", _
            "StateSolution может использовать только Strength или CrackedNDS как физическую базу."
    End If
End Sub

' Нормализует группу предельных состояний.
Private Function NormalizeLimitStateGroup(ByVal limitStateGroup As String) As String
    Select Case LCase$(Trim$(limitStateGroup))
        Case "ii", "2", "sls", "group2"
            NormalizeLimitStateGroup = "II"
        Case "i", "1", "uls", "group1"
            NormalizeLimitStateGroup = "I"
        Case Else
            Err.Raise vbObjectError + 3245, "CMaterialModelProvider", "ГПС material provider должна быть I или II."
    End Select
End Function

' Нормализует тип диаграммы.
Private Function NormalizeDiagramType(ByVal diagramType As String) As String
    Select Case LCase$(Trim$(diagramType))
        Case "threeline", "three"
            NormalizeDiagramType = "ThreeLine"
        Case Else
            NormalizeDiagramType = "TwoLine"
    End Select
End Function

' Нормализует режим растянутого бетона.
Private Function NormalizeTensionBehavior(ByVal tensionBehavior As String) As String
    Select Case LCase$(Trim$(tensionBehavior))
        Case "usediagram", "use"
            NormalizeTensionBehavior = "UseDiagram"
        Case Else
            NormalizeTensionBehavior = "Ignore"
    End Select
End Function

' Возвращает физическую бетонную диаграмму без extension.
Public Function ConcreteMaterial(ByVal purpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    Set ConcreteMaterial = mConcreteDiagram(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую диаграмму арматуры без extension.
Public Function SteelMaterial(ByVal purpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    Set SteelMaterial = mSteelDiagram(PhysicalPurposeSlot(purpose))
End Function

' Возвращает бетонную диаграмму для прямого StateSolution.
Public Function ConcreteStateMaterial(ByVal basePurpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    ValidateStateBasePurpose basePurpose
    If mDirectStateDiagramExtension Then
        Set ConcreteStateMaterial = DiagramWithExtension(mConcreteDiagram(PhysicalPurposeSlot(basePurpose)), _
            mConcreteParameters.Eb, mConcreteParameters.Ebt, True, False)
    Else
        Set ConcreteStateMaterial = mConcreteDiagram(PhysicalPurposeSlot(basePurpose))
    End If
End Function

' Возвращает диаграмму арматуры для прямого StateSolution.
Public Function SteelStateMaterial(ByVal basePurpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    ValidateStateBasePurpose basePurpose
    If mDirectStateDiagramExtension Then
        Set SteelStateMaterial = DiagramWithExtension(mSteelDiagram(PhysicalPurposeSlot(basePurpose)), _
            mSteelParameters.Esc, mSteelParameters.Es, True, True)
    Else
        Set SteelStateMaterial = mSteelDiagram(PhysicalPurposeSlot(basePurpose))
    End If
End Function

' Возвращает физическую бетонную диаграмму по спецификации профиля.
' Это новый основной вход provider-а: профиль уже определил ValueSet,
' TwoLine/ThreeLine и режим растянутого бетона, а solver получает готовую
' CMaterialDiagram без знания этих настроек.
Public Function ConcreteMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec
    Set ConcreteMaterialFromSpec = BuildConcreteDiagram(LimitStateGroupFromSpec(spec), _
        spec.ConcreteDiagramText, spec.ConcreteTensionText)
End Function

' Возвращает физическую диаграмму арматуры по спецификации профиля.
Public Function SteelMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec
    Set SteelMaterialFromSpec = BuildSteelDiagram(LimitStateGroupFromSpec(spec), spec.SteelDiagramText)
End Function

' Возвращает бетонную диаграмму прямого StateSolution по спецификации профиля.
' Numerical extension добавляется только здесь и только если соответствующая
' настройка включена. Capacity, Mcrc и crack-состояния вызывают физический
' ConcreteMaterialFromSpec без extension.
Public Function ConcreteStateMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    Dim physical As CMaterialDiagram
    Set physical = ConcreteMaterialFromSpec(spec)
    If mDirectStateDiagramExtension Then
        Set ConcreteStateMaterialFromSpec = DiagramWithExtension(physical, _
            mConcreteParameters.Eb, mConcreteParameters.Ebt, True, False)
    Else
        Set ConcreteStateMaterialFromSpec = physical
    End If
End Function

' Возвращает диаграмму арматуры прямого StateSolution по спецификации профиля.
Public Function SteelStateMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    Dim physical As CMaterialDiagram
    Set physical = SteelMaterialFromSpec(spec)
    If mDirectStateDiagramExtension Then
        Set SteelStateMaterialFromSpec = DiagramWithExtension(physical, _
            mSteelParameters.Esc, mSteelParameters.Es, True, True)
    Else
        Set SteelStateMaterialFromSpec = physical
    End If
End Function

Public Function ConcreteCompressionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ConcreteCompressionLimitFromSpec = ConcreteMaterialFromSpec(spec).UltimateCompressionStrain
End Function

Public Function ConcreteTensionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ConcreteTensionLimitFromSpec = ConcreteMaterialFromSpec(spec).UltimateTensionStrain
End Function

' Возвращает eps_bt0 из исходных параметров бетона для нормативного критерия
' Mcrc по СП 63 п. 8.1.30, формула (8.54). Значение берется из той же
' профильной спецификации, что и диаграмма Mcrc, чтобы crack-калькулятор не
' читал Config напрямую и не создавал параллельный источник материалов.
Public Function ConcreteTensionEbt0FromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteTensionEbt0FromSpec = mConcreteParameters.Ebt0
End Function

Public Function ConcreteTensionLimitEnabledFromSpec(ByVal spec As CMaterialModelSpec) As Boolean
    Dim diagram As CMaterialDiagram
    Set diagram = ConcreteMaterialFromSpec(spec)
    ConcreteTensionLimitEnabledFromSpec = (diagram.HasPhysicalTensionLimit And spec.ConcreteTension = ctkUseDiagram)
End Function

Public Function SteelCompressionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    SteelCompressionLimitFromSpec = SteelMaterialFromSpec(spec).UltimateCompressionStrain
End Function

Public Function SteelTensionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    SteelTensionLimitFromSpec = SteelMaterialFromSpec(spec).UltimateTensionStrain
End Function

Public Function ConcreteCompressionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths LimitStateGroupFromSpec(spec), rb, rbt
    ConcreteCompressionResistanceFromSpec = rb
End Function

Public Function ConcreteTensionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths LimitStateGroupFromSpec(spec), rb, rbt
    ConcreteTensionResistanceFromSpec = rbt
End Function

Public Function ConcreteElasticModulusFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteElasticModulusFromSpec = mConcreteParameters.Eb
End Function

Public Function SteelElasticModulusFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    SteelElasticModulusFromSpec = mSteelParameters.Es
End Function

' ГПС определяется спецификацией профиля, а не расчетным классом.
Private Function LimitStateGroupFromSpec(ByVal spec As CMaterialModelSpec) As String
    ValidateSpec spec
    If spec.ValueSet = mvsSLS Then
        LimitStateGroupFromSpec = "II"
    Else
        LimitStateGroupFromSpec = "I"
    End If
End Function

' Общая проверка контракта профильной спецификации.
Private Sub ValidateSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Err.Raise vbObjectError + 3250, "CMaterialModelProvider", "MaterialModelSpec не передан."
    If Not spec.IsComplete Then Err.Raise vbObjectError + 3251, "CMaterialModelProvider", "MaterialModelSpec заполнен не полностью."
End Sub

Public Function LimitStateGroup(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    LimitStateGroup = mLimitStateGroup(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteDiagramType(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    ConcreteDiagramType = mConcreteDiagramType(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteTensionBehavior(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    ConcreteTensionBehavior = mConcreteTensionBehavior(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelDiagramType(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    SteelDiagramType = mSteelDiagramType(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteCompressionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteCompressionLimit = mConcreteCompressionLimit(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteTensionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteTensionLimit = mConcreteTensionLimit(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteTensionLimitEnabled(ByVal purpose As ECalculationPurpose) As Boolean
    EnsureInitialized
    ConcreteTensionLimitEnabled = mConcreteTensionLimitEnabled(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelCompressionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelCompressionLimit = mSteelCompressionLimit(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelTensionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelTensionLimit = mSteelTensionLimit(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteCompressionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteCompressionResistance = mConcreteCompressionResistance(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteTensionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteTensionResistance = mConcreteTensionResistance(PhysicalPurposeSlot(purpose))
End Function

Public Function ConcreteElasticModulus(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteElasticModulus = mConcreteElasticModulus(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelCompressionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelCompressionResistance = mSteelCompressionResistance(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelTensionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelTensionResistance = mSteelTensionResistance(PhysicalPurposeSlot(purpose))
End Function

Public Function SteelElasticModulus(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelElasticModulus = mSteelElasticModulus(PhysicalPurposeSlot(purpose))
End Function

Public Property Get DirectStateDiagramExtensionEnabled() As Boolean
    EnsureInitialized
    DirectStateDiagramExtensionEnabled = mDirectStateDiagramExtension
End Property

' Проверяет, что provider уже построил диаграммы.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3242, "CMaterialModelProvider", "Material provider не инициализирован."
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Этот API нужен unit-тестам материальных диаграмм и расчетных сценариев без
' чтения листа Config. В рабочем пользовательском сценарии provider должен
' инициализироваться через Initialize(settings, units).

Public Sub InitializeFromParameters(ByVal concreteParameters As CConcreteMaterialParameters, _
        ByVal steelParameters As CSteelMaterialParameters, _
        Optional ByVal strengthConcreteDiagram As String = "TwoLine", _
        Optional ByVal strengthConcreteTension As String = "Ignore", _
        Optional ByVal strengthSteelDiagram As String = "TwoLine", _
        Optional ByVal mcrcConcreteDiagram As String = "ThreeLine", _
        Optional ByVal mcrcSteelDiagram As String = "TwoLine", _
        Optional ByVal crackedConcreteDiagram As String = "TwoLine", _
        Optional ByVal crackedSteelDiagram As String = "TwoLine", _
        Optional ByVal directStateDiagramExtension As Boolean = True)
    If concreteParameters Is Nothing Then Err.Raise vbObjectError + 3243, "CMaterialModelProvider", "Параметры бетона не переданы."
    If steelParameters Is Nothing Then Err.Raise vbObjectError + 3244, "CMaterialModelProvider", "Параметры арматуры не переданы."

    Set mConcreteParameters = concreteParameters
    Set mSteelParameters = steelParameters
    mDirectStateDiagramExtension = directStateDiagramExtension
    BuildPurpose cpStrength, "I", strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram
    BuildPurpose cpMcrc, "II", mcrcConcreteDiagram, "UseDiagram", mcrcSteelDiagram
    BuildPurpose cpCrackedNDS, "II", crackedConcreteDiagram, "Ignore", crackedSteelDiagram
    mInitialized = True
End Sub

Public Property Get ConcreteParameters() As CConcreteMaterialParameters
    EnsureInitialized
    Set ConcreteParameters = mConcreteParameters
End Property

Public Property Get SteelParameters() As CSteelMaterialParameters
    EnsureInitialized
    Set SteelParameters = mSteelParameters
End Property
