VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthFormulaCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чистая формула ширины раскрытия нормальной трещины
' ==========================================================================
' CCrackWidthFormulaCalculator выполняет только последнюю формульную операцию
' расчета a_crc по уже подготовленным численным данным. Класс не получает
' State-объекты, не выбирает растянутую зону, не выбирает арматуру, не считает
' Abt/As/ds/ls/sigma_s из геометрии и не запускает solver.
'
' Вся инженерная подготовка данных остается в orchestration-слое трещин:
' CurrentCrackedState/PostCrackState -> подготовленные Abt, As, ds_eq, ls,
' sigma_s, sigma_s,crc, psi_s и phi -> формула -> a_crc.

' Возвращает ширину раскрытия трещины по СП 63 через уже подготовленные
' коэффициенты и расстояние между трещинами. Проверки применимости выполнены
' вызывающим слоем; здесь остается только математическая операция.
Public Function CrackWidth(ByVal phi1 As Double, ByVal phi2 As Double, _
        ByVal phi3 As Double, ByVal psiS As Double, ByVal sigmaS As Double, _
        ByVal steelEs As Double, ByVal crackSpacing As Double) As Double
    If steelEs <= 0# Then Exit Function
    CrackWidth = phi1 * phi2 * phi3 * psiS * (sigmaS / steelEs) * crackSpacing
End Function

' Возвращает использование по ширине раскрытия. Нулевой предел означает, что
' проверочное отношение не определено, но сама формула a_crc могла быть посчитана.
Public Function Utilization(ByVal crackWidthValue As Double, _
        ByVal allowableCrackWidth As Double) As Double
    If allowableCrackWidth <= 0# Then Exit Function
    Utilization = crackWidthValue / allowableCrackWidth
End Function
