VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalculationMaterialSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Готовый набор материалов для одного расчетного назначения
' ==========================================================================
' Объект связывает выбранный режим расчета с конкретными диаграммами бетона
' и арматуры и хранит производные пределы деформаций. Это snapshot
' материалной модели, который batch передает solver-ам, writer-ам и
' crack-калькулятору.
'
' Важно: solver не выбирает Rb/Rb,ser, Rs/Rs,ser, TwoLine/ThreeLine и учет
' растянутого бетона. Он получает ConcreteMaterial и SteelMaterial из этого
' набора и решает только равновесие.

Private mPurpose As String ' Strength, Mcrc или CrackedNDS.
Private mLimitStateGroup As String ' I или II ГПС для справки и snapshot.
Private mConcreteDiagramType As String ' TwoLine или ThreeLine для бетона.
Private mConcreteTensionBehavior As String ' Ignore или UseDiagram.
Private mSteelDiagramType As String ' TwoLine или ThreeLine для арматуры.
Private mConcrete As Object ' Готовый материал бетона для solver-а.
Private mSteel As Object ' Готовый материал арматуры для solver-а.
Private mConcreteCompressionLimit As Double ' Предельная сжатая деформация бетона.
Private mConcreteTensionLimit As Double ' Предельная растянутая деформация бетона, если она включена.
Private mConcreteTensionLimitEnabled As Boolean ' True, когда растянутая ветвь бетона реально входит в диаграмму.
Private mSteelCompressionLimit As Double ' Предельная сжатая деформация арматуры.
Private mSteelTensionLimit As Double ' Предельная растянутая деформация арматуры.
Private mConcreteCompressionResistance As Double ' Использованное сопротивление бетона сжатию, МПа.
Private mConcreteTensionResistance As Double ' Использованное сопротивление бетона растяжению, МПа.
Private mConcreteElasticModulus As Double ' Использованный модуль бетона, МПа.
Private mSteelTensionResistance As Double ' Использованное сопротивление арматуры растяжению, МПа.
Private mSteelCompressionResistance As Double ' Использованное сопротивление арматуры сжатию, МПа.
Private mSteelElasticModulus As Double ' Использованный модуль арматуры растяжения, МПа.
Private mRebarProfile As String ' Ribbed/Smooth для расчета трещин.

' Заполняет набор готовыми материалами и расчетными характеристиками.
Public Sub Initialize(ByVal purpose As String, ByVal limitStateGroup As String, _
        ByVal concreteDiagramType As String, ByVal concreteTensionBehavior As String, _
        ByVal steelDiagramType As String, ByVal concreteMaterial As Object, ByVal steelMaterial As Object, _
        ByVal concreteCompressionLimit As Double, ByVal concreteTensionLimit As Double, _
        ByVal concreteTensionLimitEnabled As Boolean, ByVal steelCompressionLimit As Double, _
        ByVal steelTensionLimit As Double, ByVal concreteCompressionResistance As Double, _
        ByVal concreteTensionResistance As Double, ByVal concreteElasticModulus As Double, _
        ByVal steelCompressionResistance As Double, ByVal steelTensionResistance As Double, _
        ByVal steelElasticModulus As Double, ByVal rebarProfile As String)
    If concreteMaterial Is Nothing Then Err.Raise vbObjectError + 3210, "CCalculationMaterialSet", "Материал бетона не передан."
    If steelMaterial Is Nothing Then Err.Raise vbObjectError + 3211, "CCalculationMaterialSet", "Материал арматуры не передан."

    mPurpose = purpose
    mLimitStateGroup = limitStateGroup
    mConcreteDiagramType = concreteDiagramType
    mConcreteTensionBehavior = concreteTensionBehavior
    mSteelDiagramType = steelDiagramType
    Set mConcrete = concreteMaterial
    Set mSteel = steelMaterial
    mConcreteCompressionLimit = concreteCompressionLimit
    mConcreteTensionLimit = concreteTensionLimit
    mConcreteTensionLimitEnabled = concreteTensionLimitEnabled
    mSteelCompressionLimit = steelCompressionLimit
    mSteelTensionLimit = steelTensionLimit
    mConcreteCompressionResistance = concreteCompressionResistance
    mConcreteTensionResistance = concreteTensionResistance
    mConcreteElasticModulus = concreteElasticModulus
    mSteelCompressionResistance = steelCompressionResistance
    mSteelTensionResistance = steelTensionResistance
    mSteelElasticModulus = steelElasticModulus
    mRebarProfile = rebarProfile
End Sub

Public Property Get Purpose() As String
    Purpose = mPurpose
End Property

Public Property Get LimitStateGroup() As String
    LimitStateGroup = mLimitStateGroup
End Property

Public Property Get ConcreteDiagramType() As String
    ConcreteDiagramType = mConcreteDiagramType
End Property

Public Property Get ConcreteTensionBehavior() As String
    ConcreteTensionBehavior = mConcreteTensionBehavior
End Property

Public Property Get SteelDiagramType() As String
    SteelDiagramType = mSteelDiagramType
End Property

Public Property Get ConcreteMaterial() As Object
    Set ConcreteMaterial = mConcrete
End Property

Public Property Get SteelMaterial() As Object
    Set SteelMaterial = mSteel
End Property

Public Property Get ConcreteCompressionLimit() As Double
    ConcreteCompressionLimit = mConcreteCompressionLimit
End Property

Public Property Get ConcreteTensionLimit() As Double
    ConcreteTensionLimit = mConcreteTensionLimit
End Property

Public Property Get ConcreteTensionLimitEnabled() As Boolean
    ConcreteTensionLimitEnabled = mConcreteTensionLimitEnabled
End Property

Public Property Get SteelCompressionLimit() As Double
    SteelCompressionLimit = mSteelCompressionLimit
End Property

Public Property Get SteelTensionLimit() As Double
    SteelTensionLimit = mSteelTensionLimit
End Property

Public Property Get ConcreteCompressionResistance() As Double
    ConcreteCompressionResistance = mConcreteCompressionResistance
End Property

Public Property Get ConcreteTensionResistance() As Double
    ConcreteTensionResistance = mConcreteTensionResistance
End Property

Public Property Get ConcreteElasticModulus() As Double
    ConcreteElasticModulus = mConcreteElasticModulus
End Property

Public Property Get SteelCompressionResistance() As Double
    SteelCompressionResistance = mSteelCompressionResistance
End Property

Public Property Get SteelTensionResistance() As Double
    SteelTensionResistance = mSteelTensionResistance
End Property

Public Property Get SteelElasticModulus() As Double
    SteelElasticModulus = mSteelElasticModulus
End Property

Public Property Get RebarProfile() As String
    RebarProfile = mRebarProfile
End Property
