VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestLinearConcreteMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: Линейная модель бетона для контрольных задач
' ==========================================================================
' Задаёт sigma = E * epsilon без нелинейных участков и физических пределов.
' Позволяет проверять равновесие на простых эталонных задачах.
' В пользовательском расчёте применяется CMaterialDiagram.

Private mElasticModulus As Double ' ДЛЯ ТЕСТОВ: Постоянный модуль линейной тестовой модели, МПа.

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Задаёт простую линейную зависимость напряжения от деформации
' для проверки равновесия без нелинейных диаграмм и нормативных параметров материалов.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задает модуль упругости линейного тестового бетона.
Public Sub TestInitialize(ByVal elasticModulus As Double)
    If elasticModulus <= 0# Then Err.Raise vbObjectError + 3000, "CTestLinearConcreteMaterial", "Модуль упругости бетона должен быть положительным."
    mElasticModulus = elasticModulus
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Постоянный модуль линейной тестовой модели, МПа.
Public Property Get TestElasticModulus() As Double
    TestElasticModulus = mElasticModulus
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает напряжение по линейному закону sigma = E * epsilon.
Public Function GetStress(ByVal strain As Double) As Double
    GetStress = mElasticModulus * strain
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает постоянный касательный модуль для линейного тестового материала.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    GetTangentModulus = mElasticModulus
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает постоянный секущий модуль для линейного тестового материала.
Public Function GetSecantModulus(ByVal strain As Double) As Double
    GetSecantModulus = mElasticModulus
End Function






