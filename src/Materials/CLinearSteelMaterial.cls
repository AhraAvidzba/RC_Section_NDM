VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinearSteelMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Упрощенный линейный материал арматуры для тестов
' ==========================================================================
' Класс нужен для малых эталонных задач, где важно проверить равновесие и
' матрицы solver-а без нелинейной диаграммы. В пользовательском сценарии
' арматура обычно задается CSteelDiagramMaterial.

Private mElasticModulus As Double

Public Sub Initialize(ByVal elasticModulus As Double)
    If elasticModulus <= 0# Then Err.Raise vbObjectError + 3010, "CLinearSteelMaterial", "Steel modulus must be positive."
    mElasticModulus = elasticModulus
End Sub

Public Property Get ElasticModulus() As Double
    ElasticModulus = mElasticModulus
End Property

Public Function GetStress(ByVal strain As Double) As Double
    GetStress = mElasticModulus * strain
End Function

Public Function GetTangentModulus(ByVal strain As Double) As Double
    GetTangentModulus = mElasticModulus
End Function

Public Function GetSecantModulus(ByVal strain As Double) As Double
    GetSecantModulus = mElasticModulus
End Function







