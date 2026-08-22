Attribute VB_Name = "modGeometryTypes"
Option Explicit

' ==========================================================================
' Общие геометрические константы и типы
' ==========================================================================
' Модуль содержит небольшие общие определения, которыми пользуются геометрия,
' раскладчики арматуры и расчетные модели. Здесь не должно быть логики
' конкретной формы сечения или обращения к Excel: только простые константы и
' функции, одинаково полезные всем геометрическим слоям.

Public Const GEOM_PI As Double = 3.1415926535897931
Public Const GEOM_TOLERANCE As Double = 0.000000001

Public Type TPoint2D
    X As Double
    Y As Double
End Type

Public Type TFiber
    X As Double
    Y As Double
    Width As Double
    Height As Double
    Area As Double
    FillFactor As Double
    MaterialID As Long
End Type

Public Type TGeometryProperties
    Area As Double
    StaticMomentX As Double
    StaticMomentY As Double
    CentroidX As Double
    CentroidY As Double
    Ix As Double
    Iy As Double
    Ixy As Double
    Ixc As Double
    Iyc As Double
    Ixyc As Double
    PrincipalI1 As Double
    PrincipalI2 As Double
    PrincipalAngleRad As Double
    RadiusX As Double
    RadiusY As Double
    PrincipalRadius1 As Double
    PrincipalRadius2 As Double
End Type

Public Function GeomMax(ByVal A As Double, ByVal B As Double) As Double
    If A >= B Then
        GeomMax = A
    Else
        GeomMax = B
    End If
End Function

Public Function GeomMin(ByVal A As Double, ByVal B As Double) As Double
    If A <= B Then
        GeomMin = A
    Else
        GeomMin = B
    End If
End Function










' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildGeneratedSectionModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, Optional ByVal sourceType As String = "Generated") As CSectionModel
    Dim builder As CSectionModelBuilder
    Set builder = New CSectionModelBuilder
    Set BuildGeneratedSectionModel = builder.BuildFromGenerated(mesh, rebars, sourceType)
End Function


