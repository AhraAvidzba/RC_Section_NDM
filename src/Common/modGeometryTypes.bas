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

' Скалярные helpers выбирают крайнее значение в геометрических расчетах.
' Не зависят от Excel и не приводят значения к пользовательским единицам.
Public Function GeomMax(ByVal A As Double, ByVal B As Double) As Double
    If A >= B Then
        GeomMax = A
    Else
        GeomMax = B
    End If
End Function

' Выбирает меньший геометрический скаляр в той же системе единиц;
' преобразование пользовательских единиц остается обязанностью CUnitSystem.
Public Function GeomMin(ByVal A As Double, ByVal B As Double) As Double
    If A <= B Then
        GeomMin = A
    Else
        GeomMin = B
    End If
End Function

' Читает signed sweep semantic-дуги целиком, без приема числового префикса.
' Точка/запятая и десятичная E-нотация допустимы; локаль преобразования
' берется у VBA, а не у Excel. Дуга меньше полного оборота: для окружности
' существует CONTOUR_CIRCLE. Нулевой sweep сохраняет вырожденный отрезок.
' Место ввода передает вызывающий reader; обращения к листам здесь нет.
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

' Считает запрошенные позиции активной линии до удаления геометрических дублей.
' Double исключает переполнение суммы рядов; EverySecondBar берет нечетные
' позиции первого ряда. Это техническая оценка счетчиков, не физический лимит
' армирования. Выключенная линия не создает и зависимые дополнительные ряды.
Public Function RebarRequestedPositionCount(ByVal baseCount As Long, ByVal baseDiameter As Double, _
        ByVal row2Diameter As Double, ByVal row3Diameter As Double, _
        ByVal row2Binding As String, ByVal row3Binding As String) As Double
    If baseCount <= 0 Or baseDiameter <= 0# Then Exit Function
    Dim countValue As Double, rowCount As Double
    countValue = CDbl(baseCount)
    RebarRequestedPositionCount = countValue
    If row2Diameter > 0# Then
        rowCount = countValue
        If StrComp(row2Binding, "EverySecondBar", vbTextCompare) = 0 Then rowCount = Fix((countValue + 1#) / 2#)
        RebarRequestedPositionCount = RebarRequestedPositionCount + rowCount
    End If
    If row3Diameter > 0# Then
        rowCount = countValue
        If StrComp(row3Binding, "EverySecondBar", vbTextCompare) = 0 Then rowCount = Fix((countValue + 1#) / 2#)
        RebarRequestedPositionCount = RebarRequestedPositionCount + rowCount
    End If
End Function

' Создает CSectionModel коротким путем для тестов и небольших расчетных примеров.
' Рабочий Excel-pipeline собирает модель через CSectionModelBuilder напрямую,
' чтобы каноническая точка сборки сечения была видна в production-коде.
Public Function BuildGeneratedSectionModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout, Optional ByVal sourceType As String = "Generated") As CSectionModel
    Dim builder As CSectionModelBuilder
    Set builder = New CSectionModelBuilder
    Set BuildGeneratedSectionModel = builder.BuildFromGenerated(mesh, rebars, sourceType)
End Function


