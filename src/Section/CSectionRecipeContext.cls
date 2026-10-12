VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Ввод рецепта: неизменяемые таблицы и единицы. Готовое definition не хранит
' context, reader или таблицы; к расчетному ядру передаются только числа.
Private mTables As Object
Private mUnits As CUnitSystem
Private mSealed As Boolean

' Открывает контекст с подготовленной системой единиц один раз.
Public Sub Initialize(ByVal units As CUnitSystem)
    If Not mTables Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Контекст уже создан."
    If units Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Система единиц не передана."
    Set mUnits = units
    Set mTables = CreateObject("Scripting.Dictionary")
    mTables.CompareMode = vbTextCompare
End Sub

' Добавляет immutable input snapshot с уникальным именем до Seal.
Public Sub AddTable(ByVal table As CSectionInputTable)
    If mTables Is Nothing Or mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст не открыт для заполнения."
    If table Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Таблица не передана."
    If table.RowCount = 0 Then SectionRecipeError "CSectionRecipeContext", "Таблица не инициализирована."
    If mTables.Exists(table.Name) Then SectionRecipeError "CSectionRecipeContext", "Повторная таблица '" & table.Name & "'."
    mTables.Add table.Name, table
End Sub

' Закрывает состав таблиц; после этого добавление невозможно.
Public Sub Seal()
    If mTables Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Контекст не инициализирован."
    mSealed = True
End Sub

' Возвращает снимок зарегистрированного блока по имени; отсутствие — ошибка, не fallback.
Public Function Table(ByVal tableName As String) As CSectionInputTable
    If Not mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст еще не опубликован."
    If Not mTables.Exists(tableName) Then SectionRecipeError "CSectionRecipeContext", "Не передана таблица '" & tableName & "'."
    Set Table = mTables(tableName)
End Function

' Система единиц внешнего workflow; не перезагружайте её из Excel в рецепте.
Public Property Get Units() As CUnitSystem
    If Not mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст еще не опубликован."
    Set Units = mUnits
End Property
