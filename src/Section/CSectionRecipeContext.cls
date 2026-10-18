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

' ==========================================================================
' Входные данные рецепта сечения
' ==========================================================================
' Передаёт неизменяемые входные таблицы и правила перевода единиц.
' Готовое описание сечения не хранит этот контекст или ссылки на Excel;
' в расчётное ядро поступают подготовленные геометрические данные.

Private mTables As Object ' Копии входных таблиц рецепта, доступные по именам.
Private mUnits As CUnitSystem ' Отдельная копия настроек единиц; последующее изменение переданного объекта units не влияет на вход рецепта.
Private mSealed As Boolean ' После завершения подготовки добавлять таблицы и менять параметры запрещено.

' Один раз задаёт систему единиц для чтения входных таблиц рецепта.
Public Sub Initialize(ByVal units As CUnitSystem)
    If Not mTables Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Контекст уже создан."
    If units Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Система единиц не передана."
    Set mUnits = units.Clone
    mUnits.Freeze
    Set mTables = CreateObject("Scripting.Dictionary")
    mTables.CompareMode = vbTextCompare
End Sub

' Добавляет сохранённую входную таблицу с уникальным именем до завершения контекста.
Public Sub AddTable(ByVal table As CSectionInputTable)
    If mTables Is Nothing Or mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст не открыт для заполнения."
    If table Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Таблица не передана."
    If table.RowCount = 0 Then SectionRecipeError "CSectionRecipeContext", "Таблица не инициализирована."
    If mTables.Exists(table.Name) Then SectionRecipeError "CSectionRecipeContext", "Повторная таблица '" & table.Name & "'."
    mTables.Add table.Name, table
End Sub

' Завершает состав входных таблиц; после этого добавлять таблицы нельзя.
Public Sub Seal()
    If mTables Is Nothing Then SectionRecipeError "CSectionRecipeContext", "Контекст не инициализирован."
    mSealed = True
End Sub

' Возвращает зарегистрированную входную таблицу по имени.
' Если таблица отсутствует, сообщает об ошибке без подстановки другого источника.
Public Function Table(ByVal tableName As String) As CSectionInputTable
    If Not mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст еще не опубликован."
    If Not mTables.Exists(tableName) Then SectionRecipeError "CSectionRecipeContext", "Не передана таблица '" & tableName & "'."
    Set Table = mTables(tableName)
End Function

' Создаёт ридер с защищёнными единицами: по ключам блоков либо по позиции ValueAt.
' rowPrefix ограничивает чтение строками с общей начальной частью ключа; для общей таблицы это имя типа и точка.
' Таблица произвольной структуры также доступна через Table.
Public Function Reader(ByVal tableName As String, Optional ByVal rowPrefix As String = vbNullString) As CSectionInputReader
    Dim result As CSectionInputReader: Set result = New CSectionInputReader
    result.Initialize Table(tableName), Units, rowPrefix
    Set Reader = result
End Function

' Возвращает защищённую от изменений копию единиц ввода.
' Рецепт использует её для перевода размеров, не меняя настройки.
Public Property Get Units() As CUnitSystem
    If Not mSealed Then SectionRecipeError "CSectionRecipeContext", "Контекст еще не опубликован."
    Set Units = mUnits
End Property
