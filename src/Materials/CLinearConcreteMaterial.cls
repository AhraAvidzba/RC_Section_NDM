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
' Класс задает линейный закон sigma = E * epsilon для эталонных solver-задач,
' без нелинейных участков и физических пределов. Пользовательский расчет
' использует CMaterialDiagram и видимые параметры построения диаграммы.

Private mElasticModulus As Double ' Постоянный модуль линейной тестовой модели, МПа.

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Весь публичный API класса предназначен для unit-тестов solver-а. Он дает
' простую линейную связь stress/strain, чтобы тестировать равновесие без
' влияния нелинейных диаграмм и нормативного provider-а.

' Задает модуль упругости линейного тестового бетона.
Public Sub Initialize(ByVal elasticModulus As Double)
    If elasticModulus <= 0# Then Err.Raise vbObjectError + 3000, "CLinearConcreteMaterial", "Модуль упругости бетона должен быть положительным."
    mElasticModulus = elasticModulus
End Sub

Public Property Get ElasticModulus() As Double
    ElasticModulus = mElasticModulus
End Property

' Возвращает напряжение по линейному закону sigma = E * epsilon.
Public Function GetStress(ByVal strain As Double) As Double
    GetStress = mElasticModulus * strain
End Function

' Возвращает постоянный касательный модуль для линейного тестового материала.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    GetTangentModulus = mElasticModulus
End Function

' Возвращает постоянный секущий модуль для линейного тестового материала.
Public Function GetSecantModulus(ByVal strain As Double) As Double
    GetSecantModulus = mElasticModulus
End Function







