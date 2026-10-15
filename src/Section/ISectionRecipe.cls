VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISectionRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Интерфейс описания параметрического типа сечения
' ==========================================================================
' Каждый рецепт читает параметры своего типа и задаёт геометрию, армирование
' и аннотации через CSectionDefinitionBuilder. Сетку и отдельные стержни
' создают общие построители; рецепт определяет, что требуется построить.

' Возвращает готовое описание геометрии, армирования и аннотаций
' по входным таблицам и системе единиц переданного контекста.
Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
End Function
