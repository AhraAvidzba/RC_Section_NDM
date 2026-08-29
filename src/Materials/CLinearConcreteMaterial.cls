VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinearConcreteMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Упрощенный линейный бетонный материал для тестов и проверок
' ==========================================================================
' Класс реализует простой закон sigma = E * epsilon с ограничениями, где это нужно
' старым тестам. Основной пользовательский расчет должен идти через
' CMaterialDiagram и таблицу параметров построения диаграммы.

Private mElasticModulus As Double

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Весь публичный API класса предназначен для unit-тестов solver-а. Он дает
' простую линейную связь stress/strain, чтобы тестировать равновесие без
' влияния нелинейных диаграмм и нормативного provider-а.

Public Sub Initialize(ByVal elasticModulus As Double)
    If elasticModulus <= 0# Then Err.Raise vbObjectError + 3000, "CLinearConcreteMaterial", "Модуль упругости бетона должен быть положительным."
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







