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

' Один класс на тип: интерпретирует свой ввод и описывает определение
' через общий фасад. Не строит сетку, физические стержни или Excel Shapes.
Public Function CreateDefinition(ByVal context As CSectionRecipeContext) As CSectionDefinition
End Function
