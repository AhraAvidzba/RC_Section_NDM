Attribute VB_Name = "modGeometryTypes"
Option Explicit

' ==========================================================================
' Общие геометрические константы и типы
' ==========================================================================
' Модуль содержит небольшие общие определения, которыми пользуются геометрия,
' раскладчики арматуры и расчетные модели. Здесь не должно быть логики
' конкретной формы сечения или обращения к Excel: только простые константы и
' функции, одинаково полезные всем геометрическим слоям.

' Сумма сохраняет исходное значение Double после округления литералов редактором
' VBA до 15 цифр; повторная компиляция не должна менять точки контура на границе.
Public Const GEOM_PI As Double = 3.14159265358979 + 3.10862446895044E-15
Public Const GEOM_TOLERANCE As Double = 0.000000001
Public Const GEOM_MAX_REBAR_COUNT As Double = 2147483646# ' Long-счетчик должен сохранять представимость следующего шага обхода.

Public Type TPoint2D
    X As Double
    Y As Double
End Type

' Аналитическая ориентированная граница области: бетон находится слева при
' обходе outer против часовой стрелки и opening по часовой стрелке.
' Sweep=0 обозначает отрезок; ненулевой Sweep хранит signed угол дуги.
Public Type TRegionEdge
    X1 As Double
    Y1 As Double
    X2 As Double
    Y2 As Double
    CenterX As Double
    CenterY As Double
    Sweep As Double
    LoopID As Long ' Номер замкнутого контура в готовом геометрическом снимке.
End Type

Public Type TFiber
    X As Double
    Y As Double
    Width As Double
    Height As Double
    Area As Double
    FillFactor As Double
    MaterialID As Long
End Type

Public Type TGeometryProperties
    Area As Double
    StaticMomentX As Double
    StaticMomentY As Double
    CentroidX As Double
    CentroidY As Double
    Ix As Double
    Iy As Double
    Ixy As Double
    Ixc As Double
    Iyc As Double
    Ixyc As Double
    PrincipalI1 As Double
    PrincipalI2 As Double
    PrincipalAngleRad As Double
    RadiusX As Double
    RadiusY As Double
    PrincipalRadius1 As Double
    PrincipalRadius2 As Double
End Type

' Возвращает большее из двух чисел в текущих внутренних единицах,
' без обращения к Excel или перевода единиц.
Public Function GeomMax(ByVal A As Double, ByVal B As Double) As Double
    If A >= B Then
        GeomMax = A
    Else
        GeomMax = B
    End If
End Function

' Возвращает меньшее из двух чисел в текущих внутренних единицах.
' Перевод пользовательских величин выполняет CUnitSystem.
Public Function GeomMin(ByVal A As Double, ByVal B As Double) As Double
    If A <= B Then
        GeomMin = A
    Else
        GeomMin = B
    End If
End Function

' Проверяет инженерное число целиком: точка или запятая означают десятичный разделитель,
' E/e задаёт степень десяти; группировка разрядов, валюты, суффиксы и числовые префиксы не принимаются.
' Числовой Variant не проходит через строку и не теряет точность. Excel и COM здесь не вызываются.
Public Function TryReadStrictDouble(ByVal value As Variant, ByRef parsed As Double, _
        Optional ByVal allowDecimalComma As Boolean = True) As Boolean
    parsed = 0#
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then Exit Function
    If VarType(value) <> vbString Then
        If Not IsNumeric(value) Then Exit Function
        parsed = CDbl(value)
        TryReadStrictDouble = True
        Exit Function
    End If
    Dim text As String, index As Long, digits As Long, character As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function
    If Not allowDecimalComma And InStr(text, ",") > 0 Then Exit Function
    index = 1: character = Mid$(text, index, 1)
    If character = "+" Or character = "-" Then index = index + 1
    Do While index <= Len(text)
        character = Mid$(text, index, 1)
        If character < "0" Or character > "9" Then Exit Do
        digits = digits + 1: index = index + 1
    Loop
    character = Mid$(text, index, 1)
    If character = "." Or character = "," Then
        index = index + 1
        Do While index <= Len(text)
            character = Mid$(text, index, 1)
            If character < "0" Or character > "9" Then Exit Do
            digits = digits + 1: index = index + 1
        Loop
    End If
    If digits = 0 Then Exit Function
    character = Mid$(text, index, 1)
    If character = "E" Or character = "e" Then
        index = index + 1: character = Mid$(text, index, 1)
        If character = "+" Or character = "-" Then index = index + 1
        digits = 0
        Do While index <= Len(text)
            character = Mid$(text, index, 1)
            If character < "0" Or character > "9" Then Exit Do
            digits = digits + 1: index = index + 1
        Loop
        If digits = 0 Then Exit Function
    End If
    If index <= Len(text) Then Exit Function
    text = Replace$(text, ",", ".")
    parsed = CDbl(Replace$(text, ".", Mid$(CStr(0.5), 2, 1))) ' После полной проверки переводим только десятичный знак в локаль функций VBA, без распознавания группировки.
    TryReadStrictDouble = True
    Exit Function
InvalidValue:
    parsed = 0#
End Function

' Читает весь текст угла обхода дуги со знаком. Принимает точку, запятую
' и запись с E, но не извлекает численный префикс из произвольного текста.
' Угол меньше полного оборота; для окружности используется CONTOUR_CIRCLE.
' Нулевой угол задаёт прямой участок без дуги. Место ввода передаёт
' вызывающий класс; листы Excel здесь не читаются.
Public Function ReadContourArcSweep(ByVal value As Variant, ByVal inputLocation As String) As Double
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then GoTo InvalidValue
    Dim text As String, index As Long, decimalSeparator As String
    text = Replace$(Trim$(CStr(value)), ",", ".")
    If Len(text) = 0 Then GoTo InvalidValue
    If Len(text) - Len(Replace$(text, ".", vbNullString)) > 1 Then GoTo InvalidValue
    For index = 1 To Len(text)
        If InStr(1, "0123456789+-.Ee", Mid$(text, index, 1), vbBinaryCompare) = 0 Then GoTo InvalidValue
    Next index
    decimalSeparator = Mid$(CStr(0.5), 2, 1)
    text = Replace$(text, ".", decimalSeparator)
    If Not IsNumeric(text) Then GoTo InvalidValue
    ReadContourArcSweep = CDbl(text)
    If Abs(ReadContourArcSweep) >= 2# * GEOM_PI Then GoTo InvalidValue
    Exit Function
InvalidValue:
    Err.Raise vbObjectError + 5283, "ReadContourArcSweep", _
        "В сохраненном контуре угол дуги (" & inputLocation & ") должен быть полностью задан числом в радианах: " & _
        "десятичное число с точкой или запятой, по модулю меньше полного оборота 2*pi. " & _
        "Исправьте указанную ячейку или повторите расчет/импорт для восстановления снимка."
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' создаёт CSectionModel для коротких тестов и контрольных примеров.
' Обычный запуск книги использует CSectionModelBuilder напрямую
' как общее место сборки расчётной модели.
Public Function TestBuildGeneratedSectionModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, Optional ByVal sourceType As String = "Generated") As CSectionModel
    Dim builder As CSectionModelBuilder
    Set builder = New CSectionModelBuilder
    Set TestBuildGeneratedSectionModel = builder.BuildFromGenerated(mesh, rebars, sourceType)
End Function


