Attribute VB_Name = "modTestGeometryConfig"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: параметры геометрии через фактический Config
' ==========================================================================
' Проверяет передачу пользовательских геометрических полей от ячеек через
' CSystemSettingsReader/CUnitSystem и CSectionTypeRegistry к готовой форме
' и раскладке. Не ищет равновесие и не подменяет инженерный расчет fixtures.
' Измененные таблицы восстанавливаются по исходным формулам даже при ошибке.

Private Type TGeometryConfigStats
    Passed As Long
    Failed As Long
    Report As String
End Type

' Выполняет активные/неактивные варианты восьми параметров круга, границы
' и восстановление после ошибочного ввода. Счетчики доступны полному UI-suite.
Public Function RunAudit03CircleConfigTests(Optional ByRef passed As Long = 0, _
        Optional ByRef failed As Long = 0) As String
    Dim stats As TGeometryConfigStats
    Dim systemRange As Object, circleRange As Object, unitRange As Object
    Dim savedSystem As Variant, savedCircle As Variant, savedUnits As Variant
    On Error GoTo FailedRun
    Set systemRange = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    Set circleRange = ThisWorkbook.Names.Item("rngCircleGeometry").RefersToRange
    Set unitRange = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
    savedSystem = systemRange.Formula: savedCircle = circleRange.Formula: savedUnits = unitRange.Formula
    SetKey systemRange, "Geometry.Type", "Circle"
    unitRange.Cells(2, 2).Value2 = "mm"
    ConfigureCircle circleRange
    CheckCircleEffects stats, circleRange
    CheckCircleBoundaries stats, circleRange
    CheckCircleInactive stats, systemRange, circleRange
    GoTo Restore
FailedRun:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Restore:
    On Error GoTo FailedRestore
    If Not systemRange Is Nothing Then systemRange.Formula = savedSystem
    If Not circleRange Is Nothing Then circleRange.Formula = savedCircle
    If Not unitRange Is Nothing Then unitRange.Formula = savedUnits
    If Not systemRange Is Nothing Then Check stats, "circleConfig.restore.system", SameFormula(systemRange.Formula, savedSystem)
    If Not circleRange Is Nothing Then Check stats, "circleConfig.restore.circle", SameFormula(circleRange.Formula, savedCircle)
    If Not unitRange Is Nothing Then Check stats, "circleConfig.restore.units", SameFormula(unitRange.Formula, savedUnits)
    GoTo Finish
FailedRestore:
    stats.Failed = stats.Failed + 1
    stats.Report = stats.Report & "FAIL: circleConfig.restore.runtime; " & CStr(Err.Number) & "; " & Err.Description & vbCrLf
Finish:
    passed = stats.Passed: failed = stats.Failed
    RunAudit03CircleConfigTests = stats.Report & "TOTAL_CIRCLE_CONFIG: passed=" & CStr(passed) & "; failed=" & CStr(failed) & vbCrLf
End Function

' Задает достаточное защитное расстояние и три активных ряда. Каждая следующая
' проверка меняет только одно поле относительно этого физически допустимого набора.
Private Sub ConfigureCircle(ByVal target As Object)
    SetKey target, "Circle.Diameter", 1000#
    SetKey target, "Rebar.AxisDistance", 50#
    SetKey target, "Rebar.Count", 8
    SetKey target, "Rebar.Diameter", 16#
    SetKey target, "Rebar.Diameter2", 12#
    SetKey target, "Rebar.Diameter3", 10#
    SetKey target, "Rebar.Loc2row", "Stacked"
    SetKey target, "Rebar.Loc3row", "Stacked"
End Sub

' Oracle первой позиции не использует builder: ось основного ряда лежит на +X,
' Stacked сдвигает внутрь на полусумму диаметров, третий ряд перескакивает второй.
Private Sub CheckCircleEffects(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, geom As ISectionGeometry, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        Select Case CStr(key)
            Case "Circle.Diameter": SetKey target, CStr(key), 1100#
            Case "Rebar.AxisDistance": SetKey target, CStr(key), 60#
            Case "Rebar.Count": SetKey target, CStr(key), 10
            Case "Rebar.Diameter": SetKey target, CStr(key), 20#
            Case "Rebar.Diameter2": SetKey target, CStr(key), 20#
            Case "Rebar.Diameter3": SetKey target, CStr(key), 14#
            Case Else: SetKey target, CStr(key), "SideBySide"
        End Select
        ReadGeometry geom, bars
        Dim prefix As String
        prefix = "circleConfig.effect." & CStr(key)
        Select Case CStr(key)
            Case "Circle.Diameter"
                Near stats, prefix & ".concreteExtent", geom.MaxX, 550#
                Near stats, prefix & ".barAxis", bars.X(1), 500#
            Case "Rebar.AxisDistance"
                Near stats, prefix & ".barAxis", bars.X(1), 440#
                Near stats, prefix & ".row2Axis", bars.X(2), 426#
            Case "Rebar.Count"
                Check stats, prefix & ".allRows", bars.Count = 30
                Near stats, prefix & ".secondAngleX", bars.X(4), 450# * Cos(2# * GEOM_PI / 10#)
            Case "Rebar.Diameter"
                Near stats, prefix & ".diameter", bars.Diameter(1), 20#
                Near stats, prefix & ".row2Axis", bars.X(2), 434#
            Case "Rebar.Diameter2"
                Near stats, prefix & ".diameter", bars.Diameter(2), 20#
                Near stats, prefix & ".row3Skip", bars.X(3), 417#
            Case "Rebar.Diameter3"
                Near stats, prefix & ".diameter", bars.Diameter(3), 14#
                Near stats, prefix & ".row3Axis", bars.X(3), 423#
            Case "Rebar.Loc2row"
                Near stats, prefix & ".tangentX", bars.X(2), 450#
                Near stats, prefix & ".clockwiseY", bars.Y(2), -14#
                Near stats, prefix & ".row3NoSkip", bars.X(3), 437#
            Case "Rebar.Loc3row"
                Near stats, prefix & ".tangentX", bars.X(3), 450#
                Near stats, prefix & ".clockwiseY", bars.Y(3), -13#
                Near stats, prefix & ".row2Unchanged", bars.X(2), 436#
        End Select
        stats.Report = stats.Report & "CIRCLE_CONFIG_EFFECT: key=" & CStr(key) & "; bars=" & CStr(bars.Count) & vbCrLf
    Next key
End Sub

' Проверяет ошибочные численные/enum значения на том же reader/registry-маршруте,
' затем обычную раскладку. Пустые/нулевые диаметры сохраняют утвержденное выключение ряда.
Private Sub CheckCircleBoundaries(ByRef stats As TGeometryConfigStats, ByVal target As Object)
    Dim key As Variant, value As Variant, errorNumber As Long, description As String
    Dim geom As ISectionGeometry, bars As CRebarLayout
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then value = "Unknown" Else value = -1
        SetKey target, CStr(key), value
        errorNumber = 0: description = vbNullString
        On Error Resume Next
        ReadGeometry geom, bars
        errorNumber = Err.Number: description = Err.Description
        Err.Clear
        On Error GoTo 0
        Check stats, "circleConfig.invalid." & CStr(key), errorNumber <> 0 And Len(description) > 0
        ConfigureCircle target
        ReadGeometry geom, bars
        Check stats, "circleConfig.recovery." & CStr(key), bars.Count = 24
    Next key
    For Each key In Array("Rebar.Diameter", "Rebar.Diameter2", "Rebar.Diameter3")
        For Each value In Array(vbNullString, 0#)
            ConfigureCircle target
            SetKey target, CStr(key), value
            ReadGeometry geom, bars
            If CStr(key) = "Rebar.Diameter" Then
                Check stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 0
            Else
                Check stats, "circleConfig.optional." & CStr(key) & "." & CStr(value), bars.Count = 16
            End If
        Next value
    Next key
    ConfigureCircle target
    SetKey target, "Rebar.Count", 0
    ReadGeometry geom, bars
    Check stats, "circleConfig.count.zero", bars.Count = 0
    SetKey target, "Rebar.Count", 2
    ReadGeometry geom, bars
    Check stats, "circleConfig.count.minimum", bars.Count = 6
End Sub

' При другом Geometry.Type изменение каждого поля круга не должно менять
' форму и арматуру выбранного RectSet. Проверяем весь снимок, а не один счетчик.
Private Sub CheckCircleInactive(ByRef stats As TGeometryConfigStats, ByVal systemRange As Object, ByVal target As Object)
    SetKey systemRange, "Geometry.Type", "RectSet"
    Dim originalGeom As ISectionGeometry, originalBars As CRebarLayout
    Dim geom As ISectionGeometry, bars As CRebarLayout, key As Variant, index As Long, equal As Boolean
    ReadGeometry originalGeom, originalBars
    For Each key In Array("Circle.Diameter", "Rebar.AxisDistance", "Rebar.Count", "Rebar.Diameter", _
            "Rebar.Diameter2", "Rebar.Diameter3", "Rebar.Loc2row", "Rebar.Loc3row")
        ConfigureCircle target
        If InStr(1, CStr(key), "Loc", vbBinaryCompare) > 0 Then
            SetKey target, CStr(key), "SideBySide"
        Else
            SetKey target, CStr(key), 30#
        End If
        ReadGeometry geom, bars
        equal = geom.MinX = originalGeom.MinX And geom.MaxX = originalGeom.MaxX And _
            geom.MinY = originalGeom.MinY And geom.MaxY = originalGeom.MaxY And bars.Count = originalBars.Count
        If bars.Count = originalBars.Count Then
            For index = 1 To bars.Count
                If bars.X(index) <> originalBars.X(index) Or bars.Y(index) <> originalBars.Y(index) Or _
                        bars.Diameter(index) <> originalBars.Diameter(index) Then equal = False
            Next index
        End If
        Check stats, "circleConfig.inactive." & CStr(key), equal
    Next key
End Sub

' Читает текущие ячейки заново и использует штатные единицы/registry;
' прямые вызовы чистого Build здесь не могли бы обнаружить обрыв Config-передачи.
Private Sub ReadGeometry(ByRef geom As ISectionGeometry, ByRef bars As CRebarLayout)
    Dim settings As CSystemSettingsReader, units As CUnitSystem, registry As CSectionTypeRegistry
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    Set registry = New CSectionTypeRegistry
    Set geom = registry.CreateGeometry(settings, units)
    Set bars = registry.CreateRebars(geom, settings, units)
End Sub

' Меняет реальную ячейку значения по ключу таблицы; отсутствие ключа является
' ошибкой fixture, а не скрытым созданием новой настройки или default.
Private Sub SetKey(ByVal target As Object, ByVal key As String, ByVal value As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To target.Rows.Count
        If CStr(target.Cells(rowIndex, 1).Value2) = key Then
            target.Cells(rowIndex, 2).Value2 = value
            Exit Sub
        End If
    Next rowIndex
    Err.Raise vbObjectError + 5960, "modTestGeometryConfig", "Не найдена строка fixture: " & key
End Sub

' Сравнивает восстановленные двумерные Formula-массивы без округления,
' включая пустые значения; формы исходных диапазонов не изменяются тестом.
Private Function SameFormula(ByVal actual As Variant, ByVal expected As Variant) As Boolean
    Dim rowIndex As Long, columnIndex As Long
    For rowIndex = LBound(expected, 1) To UBound(expected, 1)
        For columnIndex = LBound(expected, 2) To UBound(expected, 2)
            If CStr(actual(rowIndex, columnIndex)) <> CStr(expected(rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    SameFormula = True
End Function

' Фиксирует абсолютный геометрический допуск 1e-6 мм и оба значения oracle;
' этот допуск не меняет настройки рабочего solver-а.
Private Sub Near(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal actual As Double, ByVal expected As Double)
    Check stats, name, Abs(actual - expected) <= 0.000001
    stats.Report = stats.Report & "CIRCLE_CONFIG_NUMBER: " & name & "; actual=" & CStr(actual) & "; expected=" & CStr(expected) & vbCrLf
End Sub

' Сохраняет счет и строку каждого assertion, не останавливая набор при несовпадении.
Private Sub Check(ByRef stats As TGeometryConfigStats, ByVal name As String, ByVal condition As Boolean)
    If condition Then
        stats.Passed = stats.Passed + 1
        stats.Report = stats.Report & "OK: " & name & vbCrLf
    Else
        stats.Failed = stats.Failed + 1
        stats.Report = stats.Report & "FAIL: " & name & vbCrLf
    End If
End Sub
