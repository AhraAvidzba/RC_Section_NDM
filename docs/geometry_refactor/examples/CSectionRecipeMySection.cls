VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionRecipeMySection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Учебный рецепт прямоугольного сечения с пятью верхними стержнями
' ==========================================================================
' Показывает описание бетонной границы, армирования и аннотаций
' через CSectionDefinitionBuilder без отдельной расстановки стержней.
Implements ISectionRecipe

' Размеры заданы постоянными в мм. Контекст передаётся по интерфейсу рецепта; входные таблицы в
' этом примере не используются.
Private Function ISectionRecipe_CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
    Dim facade As CSectionDefinitionBuilder, path As CGeometryPath
    Dim options As CRebarRuleOptions
    Set facade = New CSectionDefinitionBuilder
    facade.Initialize "MySection"
    Set path = facade.RectanglePath("Boundary", 0#, 0#, 600#, 1000#)
    facade.Outer "OUTER", path.ID
    Set path = facade.LinePath("TopAxis", 50#, 950#, 550#, 950#)
    Set options = New CRebarRuleOptions
    options.AnnotationGroup = "Top"
    facade.RebarByCount "Top.Main", path.ID, 5, 20#, options
    facade.RebarRowLabel "Top", "Top.Main", 0#, 1#
    facade.HorizontalDistance "DIM_B", "B", 0#, 600#, 1000#
    Set ISectionRecipe_CreateDefinition = facade.Finish
End Function
