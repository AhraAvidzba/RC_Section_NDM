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
' Единая точка построения расчетных диаграмм материалов
' ==========================================================================
' Provider получает исходные параметры бетона и арматуры, читает режимы
' расчетов с листа Config и строит готовые CMaterialDiagram для трех целей:
' Strength, Mcrc и CrackedNDS. Именно здесь выбираются I/II группа предельных
' состояний, сопротивления R/Rser, типы TwoLine/ThreeLine и режим растянутого
' бетона.
'
' Решатели не знают этих решений. CSectionSolver и CCapacitySolver получают
' только готовые диаграммы и вызывают общий материалный контракт GetStress,
' GetTangentModulus и GetSecantModulus. Такой слой нужен, чтобы не размазывать
' нормативный выбор по batch, crack-калькулятору, writer-ам и export-у.

Private Const PURPOSE_STRENGTH As String = "Strength" ' I ГПС, расчет прочности.
Private Const PURPOSE_MCRC As String = "Mcrc" ' II ГПС, состояние образования трещины.
Private Const PURPOSE_CRACKED_NDS As String = "CrackedNDS" ' II ГПС, НДС раскрытой трещины.

Private mConcreteParameters As CConcreteMaterialParameters ' Исходные параметры бетона из Config.
Private mSteelParameters As CSteelMaterialParameters ' Исходные параметры арматуры из Config.
Private mConcreteDiagram(1 To 3) As CMaterialDiagram ' Готовые бетонные диаграммы по расчетным целям.
Private mSteelDiagram(1 To 3) As CMaterialDiagram ' Готовые диаграммы арматуры по расчетным целям.
Private mPurposeName(1 To 3) As String ' Нормализованное имя расчетной цели.
Private mLimitStateGroup(1 To 3) As String ' I или II ГПС, фактически использованная provider-ом.
Private mConcreteDiagramType(1 To 3) As String ' TwoLine или ThreeLine для бетона.
Private mConcreteTensionBehavior(1 To 3) As String ' Ignore или UseDiagram для растянутого бетона.
Private mSteelDiagramType(1 To 3) As String ' TwoLine или ThreeLine для арматуры.
Private mConcreteCompressionLimit(1 To 3) As Double ' Предельная сжатая деформация бетонной диаграммы.
Private mConcreteTensionLimit(1 To 3) As Double ' Предельная растянутая деформация бетонной диаграммы.
Private mConcreteTensionLimitEnabled(1 To 3) As Boolean ' True, если растянутая ветвь бетона реально включена.
Private mSteelCompressionLimit(1 To 3) As Double ' Предельная сжатая деформация диаграммы арматуры.
Private mSteelTensionLimit(1 To 3) As Double ' Предельная растянутая деформация диаграммы арматуры.
Private mConcreteCompressionResistance(1 To 3) As Double ' Rb/Rb,ser выбранной ГПС, МПа.
Private mConcreteTensionResistance(1 To 3) As Double ' Rbt/Rbt,ser выбранной ГПС, МПа.
Private mSteelCompressionResistance(1 To 3) As Double ' Rsc/Rsc,ser выбранной ГПС, МПа.
Private mSteelTensionResistance(1 To 3) As Double ' Rs/Rs,ser выбранной ГПС, МПа.
Private mConcreteElasticModulus(1 To 3) As Double ' Eb выбранной бетонной модели, МПа.
Private mSteelElasticModulus(1 To 3) As Double ' Es выбранной модели арматуры, МПа.
Private mInitialized As Boolean ' Защита от использования provider-а до чтения параметров.

' Читает Config, создает параметры материалов и строит все расчетные диаграммы.
Public Sub Initialize(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3240, "CMaterialModelProvider", "Настройки материалов не переданы."

    Set mConcreteParameters = New CConcreteMaterialParameters
    mConcreteParameters.LoadFromSettings settings, units

    Set mSteelParameters = New CSteelMaterialParameters
    mSteelParameters.LoadFromSettings settings, units

    BuildPurpose PURPOSE_STRENGTH, "I", _
        settings.GetString("Diagram.Strength.Concrete", "TwoLine"), _
        settings.GetString("Diagram.Strength.ConcreteTension", "Ignore"), _
        settings.GetString("Diagram.Strength.Steel", "TwoLine")

    BuildPurpose PURPOSE_MCRC, "II", _
        settings.GetString("Diagram.Mcrc.Concrete", "ThreeLine"), _
        "UseDiagram", _
        settings.GetString("Diagram.Mcrc.Steel", "TwoLine")

    BuildPurpose PURPOSE_CRACKED_NDS, "II", _
        settings.GetString("Diagram.CrackedNDS.Concrete", "TwoLine"), _
        "Ignore", _
        settings.GetString("Diagram.CrackedNDS.Steel", "TwoLine")

    mInitialized = True
End Sub

' Строит полный материалный набор одной расчетной цели: две диаграммы и все
' численные свойства, которые потом нужны capacity, crack и Results snapshot.
Private Sub BuildPurpose(ByVal purpose As String, ByVal limitStateGroup As String, _
        ByVal concreteDiagramType As String, ByVal concreteTensionBehavior As String, _
        ByVal steelDiagramType As String)
    Dim slot As Long
    slot = PurposeSlot(purpose)

    mPurposeName(slot) = NormalizePurpose(purpose)
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
    mConcreteTensionLimitEnabled(slot) = (mConcreteTensionLimit(slot) > 0# And _
        StrComp(mConcreteTensionBehavior(slot), "UseDiagram", vbTextCompare) = 0)
    mSteelCompressionLimit(slot) = mSteelDiagram(slot).UltimateCompressionStrain
    mSteelTensionLimit(slot) = mSteelDiagram(slot).UltimateTensionStrain
End Sub

' Выбирает формулу бетонной диаграммы. Здесь остаются только нормативные
' зависимости построения точек; сама диаграмма ниже уже не знает, что она бетонная.
Private Function BuildConcreteDiagram(ByVal limitStateGroup As String, ByVal diagramType As String, _
        ByVal tensionBehavior As String) As CMaterialDiagram
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        Set BuildConcreteDiagram = BuildConcreteThreeLine(limitStateGroup, tensionBehavior)
    Else
        Set BuildConcreteDiagram = BuildConcreteTwoLine(limitStateGroup, tensionBehavior)
    End If
End Function

' Строит двухлинейную бетонную диаграмму. Белые строки Config задают eps_b1,red
' и eps_b2; provider добавляет только знаки, сопротивления выбранной ГПС и
' опциональную растянутую ветвь.
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

' Строит трехлинейную бетонную диаграмму. eps_b0/eps_bt0 и eps_b2/eps_bt2
' задает пользователь, а промежуточная точка 0.6*R/E вычисляется provider-ом.
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

' Строит двухлинейную диаграмму арматуры. Предельные eps_s2/eps_sc2 берутся
' из Config, а начало плато R/E вычисляется из сопротивления и модуля.
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

' Строит трехлинейную диаграмму арматуры. Конечные eps_s2/eps_sc2 задает
' пользователь, а eps_s0, eps_s1 и eps_s,pl рассчитываются по формулам СП.
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

' Создает CMaterialDiagram из массивов и дает единую точку для будущих проверок
' порядка точек, если они понадобятся после нормативной трассировки.
Private Function DiagramFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long) As CMaterialDiagram
    Dim diagram As CMaterialDiagram
    Set diagram = New CMaterialDiagram
    diagram.InitializeFromArrays strains, stresses, pointCount
    Set DiagramFromArrays = diagram
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

' Приводит расчетную цель к каноническому имени.
Private Function NormalizePurpose(ByVal purpose As String) As String
    Select Case LCase$(Trim$(purpose))
        Case "strength", "group1", "uls"
            NormalizePurpose = PURPOSE_STRENGTH
        Case "mcrc"
            NormalizePurpose = PURPOSE_MCRC
        Case "crackednds", "a_crc", "crack", "group2", "sls"
            NormalizePurpose = PURPOSE_CRACKED_NDS
        Case Else
            Err.Raise vbObjectError + 3241, "CMaterialModelProvider", "Неизвестная расчетная цель material provider: " & purpose
    End Select
End Function

' Возвращает индекс массивов по расчетной цели.
Private Function PurposeSlot(ByVal purpose As String) As Long
    Select Case NormalizePurpose(purpose)
        Case PURPOSE_STRENGTH: PurposeSlot = 1
        Case PURPOSE_MCRC: PurposeSlot = 2
        Case PURPOSE_CRACKED_NDS: PurposeSlot = 3
    End Select
End Function

' Нормализует группу предельных состояний. Для неизвестной строки безопаснее
' остановиться, чем молча построить диаграмму не той группы.
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

' Нормализует тип диаграммы для snapshot и диагностики.
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

Public Function ConcreteMaterial(ByVal purpose As String) As CMaterialDiagram
    EnsureInitialized
    Set ConcreteMaterial = mConcreteDiagram(PurposeSlot(purpose))
End Function

Public Function SteelMaterial(ByVal purpose As String) As CMaterialDiagram
    EnsureInitialized
    Set SteelMaterial = mSteelDiagram(PurposeSlot(purpose))
End Function

Public Function LimitStateGroup(ByVal purpose As String) As String
    EnsureInitialized
    LimitStateGroup = mLimitStateGroup(PurposeSlot(purpose))
End Function

Public Function ConcreteDiagramType(ByVal purpose As String) As String
    EnsureInitialized
    ConcreteDiagramType = mConcreteDiagramType(PurposeSlot(purpose))
End Function

Public Function ConcreteTensionBehavior(ByVal purpose As String) As String
    EnsureInitialized
    ConcreteTensionBehavior = mConcreteTensionBehavior(PurposeSlot(purpose))
End Function

Public Function SteelDiagramType(ByVal purpose As String) As String
    EnsureInitialized
    SteelDiagramType = mSteelDiagramType(PurposeSlot(purpose))
End Function

Public Function ConcreteCompressionLimit(ByVal purpose As String) As Double
    EnsureInitialized
    ConcreteCompressionLimit = mConcreteCompressionLimit(PurposeSlot(purpose))
End Function

Public Function ConcreteTensionLimit(ByVal purpose As String) As Double
    EnsureInitialized
    ConcreteTensionLimit = mConcreteTensionLimit(PurposeSlot(purpose))
End Function

Public Function ConcreteTensionLimitEnabled(ByVal purpose As String) As Boolean
    EnsureInitialized
    ConcreteTensionLimitEnabled = mConcreteTensionLimitEnabled(PurposeSlot(purpose))
End Function

Public Function SteelCompressionLimit(ByVal purpose As String) As Double
    EnsureInitialized
    SteelCompressionLimit = mSteelCompressionLimit(PurposeSlot(purpose))
End Function

Public Function SteelTensionLimit(ByVal purpose As String) As Double
    EnsureInitialized
    SteelTensionLimit = mSteelTensionLimit(PurposeSlot(purpose))
End Function

Public Function ConcreteCompressionResistance(ByVal purpose As String) As Double
    EnsureInitialized
    ConcreteCompressionResistance = mConcreteCompressionResistance(PurposeSlot(purpose))
End Function

Public Function ConcreteTensionResistance(ByVal purpose As String) As Double
    EnsureInitialized
    ConcreteTensionResistance = mConcreteTensionResistance(PurposeSlot(purpose))
End Function

Public Function ConcreteElasticModulus(ByVal purpose As String) As Double
    EnsureInitialized
    ConcreteElasticModulus = mConcreteElasticModulus(PurposeSlot(purpose))
End Function

Public Function SteelCompressionResistance(ByVal purpose As String) As Double
    EnsureInitialized
    SteelCompressionResistance = mSteelCompressionResistance(PurposeSlot(purpose))
End Function

Public Function SteelTensionResistance(ByVal purpose As String) As Double
    EnsureInitialized
    SteelTensionResistance = mSteelTensionResistance(PurposeSlot(purpose))
End Function

Public Function SteelElasticModulus(ByVal purpose As String) As Double
    EnsureInitialized
    SteelElasticModulus = mSteelElasticModulus(PurposeSlot(purpose))
End Function

Public Property Get RebarProfile() As String
    EnsureInitialized
    RebarProfile = mSteelParameters.RebarProfile
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
' инициализироваться через Initialize(settings, units), чтобы источник данных
' оставался одним и совпадал с настройками книги.

Public Sub InitializeFromParameters(ByVal concreteParameters As CConcreteMaterialParameters, _
        ByVal steelParameters As CSteelMaterialParameters, _
        Optional ByVal strengthConcreteDiagram As String = "TwoLine", _
        Optional ByVal strengthConcreteTension As String = "Ignore", _
        Optional ByVal strengthSteelDiagram As String = "TwoLine", _
        Optional ByVal mcrcConcreteDiagram As String = "ThreeLine", _
        Optional ByVal mcrcSteelDiagram As String = "TwoLine", _
        Optional ByVal crackedConcreteDiagram As String = "TwoLine", _
        Optional ByVal crackedSteelDiagram As String = "TwoLine")
    If concreteParameters Is Nothing Then Err.Raise vbObjectError + 3243, "CMaterialModelProvider", "Параметры бетона не переданы."
    If steelParameters Is Nothing Then Err.Raise vbObjectError + 3244, "CMaterialModelProvider", "Параметры арматуры не переданы."

    Set mConcreteParameters = concreteParameters
    Set mSteelParameters = steelParameters
    BuildPurpose PURPOSE_STRENGTH, "I", strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram
    BuildPurpose PURPOSE_MCRC, "II", mcrcConcreteDiagram, "UseDiagram", mcrcSteelDiagram
    BuildPurpose PURPOSE_CRACKED_NDS, "II", crackedConcreteDiagram, "Ignore", crackedSteelDiagram
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
