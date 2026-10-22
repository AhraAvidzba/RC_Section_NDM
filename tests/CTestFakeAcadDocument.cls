VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestFakeAcadDocument"
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

Private mWorldUCS As Variant ' Подставное значение признака мировой системы координат.
Private mReadFails As Boolean ' Нужно ли имитировать ошибку чтения документа.
Private mGeometryReads As Long ' Число чтений геометрии тестового документа.
Private mVariableReads As Long ' Число обращений к переменным тестового документа.
Private mCleanupLayers As Object ' Подставная таблица слоёв для проверки раннего выхода очистки.
Private mModelSpaceReads As Long ' Отдельный счётчик обращений к объектам чертежа.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' задает ответ машинного признака и сбрасывает счетчики чтений.
Public Sub TestInitialize(ByVal worldUCS As Variant, Optional ByVal readFails As Boolean = False)
    mWorldUCS = worldUCS: mReadFails = readFails
    mGeometryReads = 0: mVariableReads = 0: mModelSpaceReads = 0
    Set mCleanupLayers = Nothing
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запрос другого системного параметра выявляет неверный вызов проверки мировой
' системы координат.
Public Function GetVariable(ByVal name As String) As Variant
    mVariableReads = mVariableReads + 1
    If name <> "WORLDUCS" Or mReadFails Then Err.Raise vbObjectError + 9105, "CTestFakeAcadDocument", "Variable unavailable."
    GetVariable = mWorldUCS
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' импорт не должен доходить сюда при локальной ПСК.
Public Property Get ModelSpace() As Object
    mModelSpaceReads = mModelSpaceReads + 1
    mGeometryReads = mGeometryReads + 1
    Err.Raise vbObjectError + 9106, "CTestFakeAcadDocument", "ModelSpace accessed."
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' даже чтение слоев запрещено до успешной проверки МСК.
Public Property Get Layers() As Object
    mGeometryReads = mGeometryReads + 1
    If Not mCleanupLayers Is Nothing Then Set Layers = mCleanupLayers: Exit Property
    Err.Raise vbObjectError + 9106, "CTestFakeAcadDocument", "Layers accessed."
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' разрешает чтение только заданных слоёв; ModelSpace остаётся запрещённым.
Public Sub TestInitializeCleanupLayers(ByVal layers As Object)
    TestInitialize True
    Set mCleanupLayers = layers
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' отсутствие ModelSpace-чтений доказывает ранний выход очистки.
Public Property Get TestModelSpaceReads() As Long
    TestModelSpaceReads = mModelSpaceReads
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' счетчики отделяют отказ проверки от отказа чтения геометрии.
Public Property Get TestGeometryReads() As Long
    TestGeometryReads = mGeometryReads
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' число чтений системных переменных имитируемого документа.
Public Property Get TestVariableReads() As Long
    TestVariableReads = mVariableReads
End Property
