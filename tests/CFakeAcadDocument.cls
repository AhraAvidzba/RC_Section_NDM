VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeAcadDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: граница документа AutoCAD перед чтением геометрии
' ==========================================================================
' Возвращает заданный WORLDUCS или ошибку COM. Обращение к ModelSpace/Layers
' запрещено и подсчитывается, чтобы отказ при локальной/непроверенной ПСК
' не мог незаметно прочитать геометрию. Не заменяет живые CAD-проверки.

Private mWorldUCS As Variant
Private mReadFails As Boolean
Private mGeometryReads As Long
Private mVariableReads As Long

' ДЛЯ ТЕСТОВ: задает ответ машинного признака и сбрасывает счетчики чтений.
Public Sub Initialize(ByVal worldUCS As Variant, Optional ByVal readFails As Boolean = False)
    mWorldUCS = worldUCS: mReadFails = readFails
    mGeometryReads = 0: mVariableReads = 0
End Sub

' ДЛЯ ТЕСТОВ: иной системный параметр означает ошибочный контракт проверки.
Public Function GetVariable(ByVal name As String) As Variant
    mVariableReads = mVariableReads + 1
    If name <> "WORLDUCS" Or mReadFails Then Err.Raise vbObjectError + 9105, "CFakeAcadDocument", "Variable unavailable."
    GetVariable = mWorldUCS
End Function

' ДЛЯ ТЕСТОВ: импорт не должен доходить сюда при локальной ПСК.
Public Property Get ModelSpace() As Object
    mGeometryReads = mGeometryReads + 1
    Err.Raise vbObjectError + 9106, "CFakeAcadDocument", "ModelSpace accessed."
End Property

' ДЛЯ ТЕСТОВ: даже чтение слоев запрещено до успешной проверки МСК.
Public Property Get Layers() As Object
    mGeometryReads = mGeometryReads + 1
    Err.Raise vbObjectError + 9106, "CFakeAcadDocument", "Layers accessed."
End Property

' ДЛЯ ТЕСТОВ: счетчики отделяют отказ проверки от отказа чтения геометрии.
Public Property Get GeometryReads() As Long
    GeometryReads = mGeometryReads
End Property

Public Property Get VariableReads() As Long
    VariableReads = mVariableReads
End Property
