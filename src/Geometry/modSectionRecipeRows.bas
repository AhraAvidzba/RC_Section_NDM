Attribute VB_Name = "modSectionRecipeRows"
Option Explicit

' Подготавливает контактное смещение дополнительного ряда для штатных рецептов.
' Чтение таблиц, направление линии и политики вывода остаются в самом рецепте;
' положения стержней по этим настройкам вычисляет общий CRebarRuleExecutor.

' Возвращает новые независимые настройки дополнительного ряда. Stacked смещает стержни
' по нормали, SideBySide — вдоль линии; знаки обоих направлений задаёт рецепт.
' binding выбирает все исходные позиции или нечётные. interveningDiameterSum учитывает
' полные диаметры промежуточных рядов, независимо от фактического выбора их позиций.
Public Function RecipeContactOptions(ByVal sourceReference As String, ByVal location As String, _
        ByVal binding As String, ByVal baseDiameter As Double, ByVal diameter As Double, _
        ByVal normalSign As Double, Optional ByVal tangentSign As Double = 1#, _
        Optional ByVal interveningDiameterSum As Double = 0#) As CRebarRuleOptions
    Dim options As CRebarRuleOptions, axis As ERebarOffsetAxis, directionSign As Double
    Set options = New CRebarRuleOptions
    options.SourceReference = sourceReference ' Адреса активных полей ряда для сообщения об ошибке.
    axis = RowNormal
    directionSign = normalSign
    If location = "SideBySide" Then
        axis = RowTangent
        directionSign = tangentSign
    End If
    options.ConfigureDependentContact baseDiameter, diameter, axis, directionSign, _
        (binding = "EverySecondBar"), interveningDiameterSum ' Контактное расстояние и выбор исходных стержней.
    Set RecipeContactOptions = options
End Function
