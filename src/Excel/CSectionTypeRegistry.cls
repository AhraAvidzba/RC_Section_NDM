VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionTypeRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Единственная регистрация рецептов и имен их входных таблиц.
' Реестр не разбирает ячейки и не строит сетку, арматуру или аннотации.
Private mRecipes As Object, mRecipeTables As Object

Private Sub Class_Initialize()
    Set mRecipes = CreateObject("Scripting.Dictionary"): mRecipes.CompareMode = vbTextCompare
    Set mRecipeTables = CreateObject("Scripting.Dictionary"): mRecipeTables.CompareMode = vbTextCompare
    RegisterRecipe "Circle", New CSectionRecipeCircle, Array("rngCircleGeometry")
    RegisterRecipe "RectSet", New CSectionRecipeRectSet, Array("rngRectSetGeometry")
    RegisterRecipe "RoundedRectangle", New CSectionRecipeRoundedRectangle, Array("rngRoundedRectangleGeometry")
    RegisterRecipe "HollowRectangle", New CSectionRecipeHollowRectangle, Array("rngHollowRectangleGeometry")
End Sub

' Пустой массив таблиц допустим для программно заданного рецепта.
' Массив строк копируется; live Excel-объекты в метаданные не попадают.
Public Sub RegisterRecipe(ByVal typeID As String, ByVal recipe As ISectionRecipe, ByVal tableNames As Variant)
    typeID = Trim$(typeID)
    If Len(typeID) = 0 Or recipe Is Nothing Then SectionRecipeError "CSectionTypeRegistry", "Не заданы тип и рецепт."
    If mRecipes.Exists(typeID) Then SectionRecipeError "CSectionTypeRegistry", "Повторная регистрация типа '" & typeID & "'."
    If Not IsArray(tableNames) Then SectionRecipeError "CSectionTypeRegistry", "Имена таблиц должны быть массивом строк."
    Dim name As Variant, seen As Object
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each name In tableNames
        If VarType(name) <> vbString Then SectionRecipeError "CSectionTypeRegistry", "Имя таблицы должно быть строкой."
        If Len(Trim$(CStr(name))) = 0 Then SectionRecipeError "CSectionTypeRegistry", "Имя таблицы не может быть пустым."
        If seen.Exists(CStr(name)) Then SectionRecipeError "CSectionTypeRegistry", "Повторная входная таблица '" & CStr(name) & "'."
        seen.Add CStr(name), True
    Next name
    mRecipes.Add typeID, recipe: mRecipeTables.Add typeID, tableNames
End Sub

Public Function CreateRecipe(ByVal typeID As String) As ISectionRecipe
    If Not mRecipes.Exists(typeID) Then SectionRecipeError "CSectionTypeRegistry", "Не зарегистрирован рецепт '" & typeID & "'."
    Set CreateRecipe = mRecipes(typeID)
End Function

Public Function RecipeTableNames(ByVal typeID As String) As Variant
    If Not mRecipeTables.Exists(typeID) Then SectionRecipeError "CSectionTypeRegistry", "Не зарегистрирован ввод типа '" & typeID & "'."
    RecipeTableNames = mRecipeTables(typeID)
End Function

Public Function RegisteredRecipeTypes() As Variant
    RegisteredRecipeTypes = mRecipes.Keys
End Function

Public Function SelectedRecipeType(ByVal settings As CSystemSettingsReader) As String
    If settings Is Nothing Then SectionRecipeError "CSectionTypeRegistry", "Настройки Config не переданы."
    SelectedRecipeType = settings.GetRequiredChoice("Geometry.Type", RegisteredRecipeTypes)
End Function

' Один вызов выбранного рецепта; готовое definition не хранит ввод или reader.
Public Function CreateDefinition(ByVal typeID As String, ByVal context As CSectionRecipeContext) As CSectionDefinition
    If context Is Nothing Then SectionRecipeError "CSectionTypeRegistry", "Снимок входных таблиц не передан."
    Dim recipe As ISectionRecipe
    Set recipe = CreateRecipe(typeID)
    Set CreateDefinition = recipe.CreateDefinition(context)
End Function
