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
' Единая точка получения расчетных материальных моделей
' ==========================================================================
' Provider читает параметры материалов и настройки диаграмм из Config,
' строит три готовых набора материалов и дальше отдает их batch-расчету,
' writer-ам и расчету трещин.
'
' Это единственное место, где выбираются I/II ГПС, TwoLine/ThreeLine и
' учет растянутого бетона для конкретной расчетной цели:
' Strength, Mcrc и CrackedNDS. Solver-ы получают уже готовые материалы и не
' знают, из каких нормативных характеристик они построены.

Private Const PURPOSE_STRENGTH As String = "Strength" ' I ГПС, проверка прочности.
Private Const PURPOSE_MCRC As String = "Mcrc" ' II ГПС, состояние образования трещины.
Private Const PURPOSE_CRACKED_NDS As String = "CrackedNDS" ' II ГПС, раскрытая трещина/a_crc.

Private mConcreteParameters As CConcreteMaterialParameters ' Исходные характеристики бетона.
Private mSteelParameters As CSteelMaterialParameters ' Исходные характеристики арматуры.
Private mStrengthSet As CCalculationMaterialSet ' Материалы для прочности I ГПС.
Private mMcrcSet As CCalculationMaterialSet ' Материалы для образования трещины II ГПС.
Private mCrackedNDSSet As CCalculationMaterialSet ' Материалы для НДС раскрытой трещины II ГПС.
Private mInitialized As Boolean ' Защита от использования provider-а до чтения Config.

' Читает Config, строит параметры и готовит все расчетные наборы материалов.
Public Sub Initialize(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Err.Raise vbObjectError + 3240, "CMaterialModelProvider", "Настройки материалов не переданы."

    Set mConcreteParameters = New CConcreteMaterialParameters
    mConcreteParameters.LoadFromSettings settings, units

    Set mSteelParameters = New CSteelMaterialParameters
    mSteelParameters.LoadFromSettings settings, units

    Set mStrengthSet = BuildSet(PURPOSE_STRENGTH, "I", _
        settings.GetString("Diagram.Strength.Concrete", "TwoLine"), _
        settings.GetString("Diagram.Strength.ConcreteTension", "Ignore"), _
        settings.GetString("Diagram.Strength.Steel", "TwoLine"))

    Set mMcrcSet = BuildSet(PURPOSE_MCRC, "II", _
        settings.GetString("Diagram.Mcrc.Concrete", "ThreeLine"), _
        "UseDiagram", _
        settings.GetString("Diagram.Mcrc.Steel", "TwoLine"))

    Set mCrackedNDSSet = BuildSet(PURPOSE_CRACKED_NDS, "II", _
        settings.GetString("Diagram.CrackedNDS.Concrete", "TwoLine"), _
        "Ignore", _
        settings.GetString("Diagram.CrackedNDS.Steel", "TwoLine"))

    mInitialized = True
End Sub

' Готовит provider из заранее созданных параметров материалов.
' Это тот же путь построения диаграмм, что и Initialize(Config), но без чтения
' Excel. Метод используется тестами, чтобы проверять расчетные классы на
' контролируемых характеристиках материалов.
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
    Set mStrengthSet = BuildSet(PURPOSE_STRENGTH, "I", strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram)
    Set mMcrcSet = BuildSet(PURPOSE_MCRC, "II", mcrcConcreteDiagram, "UseDiagram", mcrcSteelDiagram)
    Set mCrackedNDSSet = BuildSet(PURPOSE_CRACKED_NDS, "II", crackedConcreteDiagram, "Ignore", crackedSteelDiagram)
    mInitialized = True
End Sub

' Создает один расчетный набор: бетонный материал, стальной материал и
' связанные с ними предельные деформации.
Private Function BuildSet(ByVal purpose As String, _
        ByVal limitStateGroup As String, ByVal concreteDiagramType As String, _
        ByVal concreteTensionBehavior As String, ByVal steelDiagramType As String) As CCalculationMaterialSet
    Dim concreteBuilder As CConcreteDiagramBuilder
    Set concreteBuilder = New CConcreteDiagramBuilder
    Dim concreteMaterial As CConcreteDiagramMaterial
    Set concreteMaterial = concreteBuilder.BuildMaterial(mConcreteParameters, limitStateGroup, concreteDiagramType, concreteTensionBehavior)

    Dim steelBuilder As CSteelDiagramBuilder
    Set steelBuilder = New CSteelDiagramBuilder
    Dim steelMaterial As CSteelDiagramMaterial
    Set steelMaterial = steelBuilder.BuildMaterial(mSteelParameters, limitStateGroup, steelDiagramType)

    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths limitStateGroup, rb, rbt

    Dim rsc As Double
    Dim rs As Double
    SteelStrengths limitStateGroup, rsc, rs

    Dim result As CCalculationMaterialSet
    Set result = New CCalculationMaterialSet
    result.Initialize purpose, limitStateGroup, NormalizeDiagramType(concreteDiagramType), _
        NormalizeTensionBehavior(concreteTensionBehavior), NormalizeDiagramType(steelDiagramType), _
        concreteMaterial, steelMaterial, concreteBuilder.CompressionLimit, concreteBuilder.TensionLimit, _
        concreteBuilder.TensionEnabled, steelBuilder.CompressionLimit, steelBuilder.TensionLimit, _
        rb, rbt, mConcreteParameters.Eb, rsc, rs, mSteelParameters.Es, mSteelParameters.RebarProfile

    Set BuildSet = result
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

' Нормализует тип диаграммы для записи в snapshot и диагностический отчет.
Private Function NormalizeDiagramType(ByVal diagramType As String) As String
    Select Case LCase$(Trim$(diagramType))
        Case "threeline", "three"
            NormalizeDiagramType = "ThreeLine"
        Case Else
            NormalizeDiagramType = "TwoLine"
    End Select
End Function

' Нормализует режим работы растянутого бетона.
Private Function NormalizeTensionBehavior(ByVal tensionBehavior As String) As String
    Select Case LCase$(Trim$(tensionBehavior))
        Case "usediagram", "use"
            NormalizeTensionBehavior = "UseDiagram"
        Case Else
            NormalizeTensionBehavior = "Ignore"
    End Select
End Function

' Возвращает набор материалов по имени расчетной цели.
Public Function MaterialSet(ByVal purpose As String) As CCalculationMaterialSet
    EnsureInitialized
    Select Case LCase$(Trim$(purpose))
        Case "strength", "group1", "uls"
            Set MaterialSet = mStrengthSet
        Case "mcrc"
            Set MaterialSet = mMcrcSet
        Case "crackednds", "a_crc", "crack", "group2", "sls"
            Set MaterialSet = mCrackedNDSSet
        Case Else
            Err.Raise vbObjectError + 3241, "CMaterialModelProvider", "Неизвестное назначение material set: " & purpose
    End Select
End Function

Public Property Get StrengthSet() As CCalculationMaterialSet
    EnsureInitialized
    Set StrengthSet = mStrengthSet
End Property

Public Property Get McrcSet() As CCalculationMaterialSet
    EnsureInitialized
    Set McrcSet = mMcrcSet
End Property

Public Property Get CrackedNDSSet() As CCalculationMaterialSet
    EnsureInitialized
    Set CrackedNDSSet = mCrackedNDSSet
End Property

Public Property Get ConcreteParameters() As CConcreteMaterialParameters
    EnsureInitialized
    Set ConcreteParameters = mConcreteParameters
End Property

Public Property Get SteelParameters() As CSteelMaterialParameters
    EnsureInitialized
    Set SteelParameters = mSteelParameters
End Property

' Защищает вызовы provider-а до Initialize.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3242, "CMaterialModelProvider", "Material provider не инициализирован."
End Sub
