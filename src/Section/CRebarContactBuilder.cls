VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarContactBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подготовка дополнительных контактных рядов арматуры
' ==========================================================================
' Один экземпляр обслуживает дополнительные ряды одного основного ряда.
' Сохраняет отдельные общие настройки этих рядов и накапливает номинальные
' диаметры по двум направлениям. Геометрию и стержни создают общие построители;
' правила и аннотации по подготовленным настройкам объявляет сам рецепт.

Private mBaseDiameter As Double ' Диаметр стержней основного ряда, мм.
Private mNormalSign As Double ' Знак смещения Stacked по нормали к линии основного ряда.
Private mTangentSign As Double ' Знак смещения SideBySide вдоль линии основного ряда.
Private mDefaults As CRebarRuleOptions ' Независимый снимок общих настроек только дополнительных рядов.
Private mNormalDiameters As Double ' Сумма диаметров успешно подготовленных рядов Stacked, мм.
Private mTangentDiameters As Double ' Сумма диаметров успешно подготовленных рядов SideBySide, мм.
Private mInitialized As Boolean ' Основной диаметр, направления и общие настройки уже заданы.

' Один раз задаёт основной ряд и сохраняет копию настроек его дополнительных рядов.
' Основные options сюда не передаются автоматически; оба знака должны быть равны +1 или -1.
Public Sub Initialize(ByVal baseDiameter As Double, ByVal normalSign As Double, ByVal tangentSign As Double, ByVal defaults As CRebarRuleOptions)
    Dim sourceReference As String, snapshot As CRebarRuleOptions
    If Not defaults Is Nothing Then sourceReference = defaults.SourceReference
    If mInitialized Then SectionRecipeError "CRebarContactBuilder", "Подготовщик контактных рядов уже инициализирован. Создайте отдельный экземпляр для другого основного ряда.", sourceReference
    If defaults Is Nothing Then SectionRecipeError "CRebarContactBuilder", "Не переданы общие настройки дополнительных рядов."
    If baseDiameter <= 0# Then SectionRecipeError "CRebarContactBuilder", "Диаметр основного ряда должен быть положительным.", sourceReference
    If normalSign <> -1# And normalSign <> 1# Then SectionRecipeError "CRebarContactBuilder", "Знак смещения Stacked должен быть -1 или 1.", sourceReference
    If tangentSign <> -1# And tangentSign <> 1# Then SectionRecipeError "CRebarContactBuilder", "Знак смещения SideBySide должен быть -1 или 1.", sourceReference
    Set snapshot = defaults.Clone
    mBaseDiameter = baseDiameter: mNormalSign = normalSign: mTangentSign = tangentSign
    Set mDefaults = snapshot
    mInitialized = True
End Sub

' Возвращает независимые настройки одного активного ряда и учитывает его диаметр для следующих.
' Вызывается один раз на ряд в порядке описания. EverySecondBar не уменьшает накопленные диаметры.
' При ошибке общие настройки и накопленные диаметры остаются прежними.
Public Function AppendRow(ByVal diameter As Double, ByVal location As String, ByVal binding As String, ByVal sourceReference As String) As CRebarRuleOptions
    On Error GoTo Failed
    If Not mInitialized Then SectionRecipeError "CRebarContactBuilder", "Сначала задайте основной ряд через Initialize.", sourceReference
    Dim axis As ERebarOffsetAxis, directionSign As Double, intermediate As Double, nextSum As Double
    Dim options As CRebarRuleOptions, everySecond As Boolean
    Select Case location
        Case "Stacked": axis = RowNormal: directionSign = mNormalSign: intermediate = mNormalDiameters
        Case "SideBySide": axis = RowTangent: directionSign = mTangentSign: intermediate = mTangentDiameters
        Case Else: SectionRecipeError "CRebarContactBuilder", "Неизвестное расположение ряда '" & location & "'. Выберите Stacked или SideBySide.", sourceReference
    End Select
    Select Case binding
        Case "EachBar": everySecond = False
        Case "EverySecondBar": everySecond = True
        Case Else: SectionRecipeError "CRebarContactBuilder", "Неизвестная привязка ряда '" & binding & "'. Выберите EachBar или EverySecondBar.", sourceReference
    End Select
    Set options = mDefaults.Clone
    options.SourceReference = sourceReference ' Адрес активного ряда имеет приоритет над адресом общих настроек.
    options.ConfigureDependentContact mBaseDiameter, diameter, axis, directionSign, everySecond, intermediate
    nextSum = intermediate + diameter ' Готовим новое состояние до изменения накопителя; переполнение не оставляет частичного результата.
    Set AppendRow = options
    If axis = RowNormal Then mNormalDiameters = nextSum Else mTangentDiameters = nextSum
    Exit Function
Failed:
    Dim number As Long, owner As String, reason As String
    number = Err.Number: owner = Err.Source: reason = Err.Description
    If number = 6 Then SectionRecipeError "CRebarContactBuilder", "Контактное смещение или сумма диаметров выходит за числовой диапазон. Уменьшите заданные диаметры.", sourceReference
    If Len(sourceReference) > 0 And InStr(1, reason, sourceReference, vbBinaryCompare) = 0 Then reason = sourceReference & ": " & reason
    Err.Raise number, owner, reason
End Function
