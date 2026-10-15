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
' Линейная модель бетона для контрольных задач
' ==========================================================================
' Задаёт sigma = E * epsilon без нелинейных участков и физических пределов.
' Позволяет проверять равновесие на простых эталонных задачах.
' В пользовательском расчёте применяется CMaterialDiagram.

Private mElasticModulus As Double ' Постоянный модуль линейной тестовой модели, МПа.

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Задаёт простую линейную зависимость напряжения от деформации
' для проверки равновесия без нелинейных диаграмм и нормативных параметров материалов.

' Задает модуль упругости линейного тестового бетона.
Public Sub Initialize(ByVal elasticModulus As Double)
    If elasticModulus <= 0# Then Err.Raise vbObjectError + 3000, "CLinearConcreteMaterial", "Модуль упругости бетона должен быть положительным."
    mElasticModulus = elasticModulus
End Sub

' Постоянный модуль линейной тестовой модели, МПа.
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






