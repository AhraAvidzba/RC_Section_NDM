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
' Подготовка дополнительных рядов арматуры и их контактных смещений
' ==========================================================================
' Один экземпляр обслуживает дополнительные ряды одного основного ряда.
' Отступ каждого ряда отсчитывает от основного. Промежуточные расстояния накапливает
' отдельно для совпадающих базовых отступов и двух направлений. Стержни создают общие построители;
' правила и аннотации по подготовленным настройкам объявляет сам рецепт.
' Между приставляемыми стержнями оставляет 2,5 мм на выступы арматуры периодического профиля;
' расчётные диаметры и площади стержней эта добавка не меняет.

Private Const PROFILE_GAP_MM As Double = 2.5 ' Добавка 2,5 мм учитывает выступы арматуры периодического профиля.
Private mNormalSign As Double ' Направление базового отступа и смещения Stacked по нормали к основному ряду.
Private mTangentSign As Double ' Знак смещения SideBySide вдоль линии основного ряда.
Private mDefaults As CRebarRuleOptions ' Независимый снимок общих настроек только дополнительных рядов.
Private Type TContactGroup
    Offset As Double ' Расстояние базовой линии группы от основного ряда, мм.
    Diameter As Double ' Диаметр первого стержня этой группы, мм.
    NormalSpacing As Double ' Сумма промежуточных диаметров и добавок на выступы арматуры периодического профиля по нормали только этой группы, мм.
    TangentSpacing As Double ' Сумма промежуточных диаметров и добавок на выступы арматуры периодического профиля вдоль линии только этой группы, мм.
End Type
Private mGroups() As TContactGroup ' Контактные группы одного семейства; первая соответствует основному ряду.
Private mGroupCount As Long ' Число уже подготовленных контактных групп.
Private mInitialized As Boolean ' Основной диаметр, направления и общие настройки уже заданы.

' Один раз задаёт основной ряд и сохраняет копию настроек его дополнительных рядов.
' Основные options сюда не передаются автоматически; оба знака должны быть равны +1 или -1.
Public Sub Initialize(ByVal baseDiameter As Double, ByVal normalSign As Double, ByVal tangentSign As Double, ByVal defaults As CRebarRuleOptions)
    Dim sourceReference As String, snapshot As CRebarRuleOptions
    If Not defaults Is Nothing Then sourceReference = defaults.SourceReference
    If mInitialized Then SectionRecipeError "CRebarContactBuilder", "Подготовщик контактных рядов уже инициализирован. Создайте отдельный экземпляр для другого основного ряда.", sourceReference
    If defaults Is Nothing Then SectionRecipeError "CRebarContactBuilder", "Не переданы общие настройки дополнительных рядов."
    If baseDiameter <= 0# Then SectionRecipeError "CRebarContactBuilder", "Диаметр основного ряда должен быть положительным.", sourceReference
    If normalSign <> -1# And normalSign <> 1# Then SectionRecipeError "CRebarContactBuilder", "Для размещения ряда по нормали к линии (Stacked) задайте знак направления -1 или 1.", sourceReference
    If tangentSign <> -1# And tangentSign <> 1# Then SectionRecipeError "CRebarContactBuilder", "Для размещения ряда вдоль линии (SideBySide) задайте знак направления -1 или 1.", sourceReference
    Set snapshot = defaults.Clone
    mNormalSign = normalSign: mTangentSign = tangentSign
    ReDim mGroups(1 To 1)
    mGroups(1).Diameter = baseDiameter: mGroupCount = 1
    Set mDefaults = snapshot
    mInitialized = True
End Sub

' Возвращает независимые настройки активного ряда: базовый отступ от Main плюс контактная поправка.
' Вызывается один раз на ряд в порядке описания. EverySecondBar не уменьшает накопленные расстояния.
' При ошибке общие настройки и накопленные расстояния остаются прежними.
Public Function AppendRow(ByVal diameter As Double, ByVal location As String, ByVal binding As String, ByVal sourceReference As String, Optional ByVal baseOffset As Double = 0#) As CRebarRuleOptions
    On Error GoTo Failed
    If Not mInitialized Then SectionRecipeError "CRebarContactBuilder", "Сначала задайте основной ряд через Initialize.", sourceReference
    Dim axis As ERebarOffsetAxis, directionSign As Double, intermediate As Double, nextSum As Double
    Dim options As CRebarRuleOptions, everySecond As Boolean
    Dim group As Long, i As Long, anchor As Double, groupOffset As Double
    If diameter <= 0# Then SectionRecipeError "CRebarContactBuilder", "Диаметр активного ряда должен быть положительным.", sourceReference
    If baseOffset < 0# Then SectionRecipeError "CRebarContactBuilder", "Отступ от основного ряда не может быть отрицательным.", sourceReference
    For i = 1 To mGroupCount
        If Abs(mGroups(i).Offset - baseOffset) <= GEOM_TOLERANCE Then group = i: Exit For
    Next i
    groupOffset = baseOffset
    If group > 0 Then groupOffset = mGroups(group).Offset: anchor = mGroups(group).Diameter ' Контакт отсчитывается от первого стержня группы с тем же базовым отступом.
    Select Case location
        Case "Stacked"
            axis = RowNormal: directionSign = mNormalSign
            If group > 0 Then intermediate = mGroups(group).NormalSpacing
        Case "SideBySide"
            axis = RowTangent: directionSign = mTangentSign
            If group > 0 Then intermediate = mGroups(group).TangentSpacing
        Case Else: SectionRecipeError "CRebarContactBuilder", "Неизвестное расположение ряда '" & location & "'. Выберите Stacked или SideBySide.", sourceReference
    End Select
    Select Case binding
        Case "EachBar": everySecond = False
        Case "EverySecondBar": everySecond = True
        Case Else: SectionRecipeError "CRebarContactBuilder", "Неизвестная привязка ряда '" & binding & "'. Выберите EachBar или EverySecondBar.", sourceReference
    End Select
    Set options = mDefaults.Clone
    options.SourceReference = sourceReference ' Адрес активного ряда имеет приоритет над адресом общих настроек.
    If group > 0 Then
        options.ConfigureDependentContact anchor, diameter, axis, directionSign, everySecond, intermediate + PROFILE_GAP_MM
        nextSum = intermediate + diameter + PROFILE_GAP_MM ' Следующий ряд учитывает полный диаметр уже добавляемого стержня и зазор 2,5 мм.
    Else
        options.NormalOffset = 0#: options.TangentOffset = 0#: options.EverySecond = everySecond
    End If
    options.NormalOffset = options.NormalOffset + mNormalSign * groupOffset ' Базовый отступ от Main добавляется к контактной поправке только по нормали.
    options.ValidateRowOffset = (baseOffset > 0#)
    If group = 0 Then ' Сначала завершаем все вычисления и проверки. Только затем меняем накопители.
        ReDim Preserve mGroups(1 To mGroupCount + 1)
        mGroupCount = mGroupCount + 1: group = mGroupCount
        mGroups(group).Offset = groupOffset: mGroups(group).Diameter = diameter
    ElseIf axis = RowNormal Then
        mGroups(group).NormalSpacing = nextSum
    Else
        mGroups(group).TangentSpacing = nextSum
    End If
    Set AppendRow = options
    Exit Function
Failed:
    Dim number As Long, owner As String, reason As String
    number = Err.Number: owner = Err.Source: reason = Err.Description
    If number = 6 Then SectionRecipeError "CRebarContactBuilder", "Контактное смещение или сумма диаметров выходит за числовой диапазон. Уменьшите заданные диаметры.", sourceReference
    If Len(sourceReference) > 0 And InStr(1, reason, sourceReference, vbBinaryCompare) = 0 Then reason = sourceReference & ": " & reason
    Err.Raise number, owner, reason
End Function
