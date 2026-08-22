VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISectionTypeProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Контракт provider-а типа сечения
' ==========================================================================
' Provider отвечает за создание полного специализированного набора для одной
' формы: Geometry, RebarBuilder и AnnotationBuilder. Универсальный код
' вызывает только этот интерфейс, поэтому добавление нового сечения не должно
' менять writer, reader, layout или plotter.

Public Function TypeName() As String
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildGeometry(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As ISectionGeometry
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildRebars(ByVal geometry As ISectionGeometry, ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CRebarLayout
End Function

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Sub BuildAnnotations(ByVal model As CSectionModel, ByVal geometry As ISectionGeometry, ByVal rebars As CRebarLayout)
End Sub

