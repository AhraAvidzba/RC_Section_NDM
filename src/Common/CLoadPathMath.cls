VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadPathMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общая математика lambda-траектории нагрузки
' ==========================================================================
' CLoadPathMath является единым низкоуровневым расчетным классом для задач, где
' расчетная точка должна лежать на прямой нагрузки:
'
'     Target(lambda) = Offset + lambda * Base
'
' Класс не знает, какой предел ищется: несущая способность, образование
' трещины или другой критерий. Он только выбирает устойчивую опорную
' компоненту lambda, строит две независимые невязки попадания на траекторию
' и нормирует их через допуски равновесия solver-а. Поэтому capacity и crack
' не должны держать собственные копии этой математики.

Private Const DEFAULT_BASE_TOLERANCE As Double = 0.000000001 ' Численный ноль для масштабируемой компоненты Base.
Private Const RELATIVE_COMPONENT_TOLERANCE As Double = 0.000000001 ' Минимальная относительная добавка к solver tolerance.

Private mToleranceN As Double ' Допуск равновесия по N, Н.
Private mToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mBaseTolerance As Double ' Численный ноль для определения ненулевой компоненты Base.

' Инициализирует расчетные допуски равновесия значениями, совместимыми с default solver.
Private Sub Class_Initialize()
    mToleranceN = 5#
    mToleranceMx = 5000#
    mToleranceMy = 5000#
    mBaseTolerance = DEFAULT_BASE_TOLERANCE
End Sub

' Передает допуски того solver-а, для которого будут строиться безразмерные
' невязки load-path. Разные расчетные ветки могут иметь разные настройки, но
' формула нормировки остается общей.
Public Sub Configure(ByVal toleranceN As Double, ByVal toleranceMx As Double, _
        ByVal toleranceMy As Double, Optional ByVal baseTolerance As Double = DEFAULT_BASE_TOLERANCE)
    mToleranceN = Abs(toleranceN)
    mToleranceMx = Abs(toleranceMx)
    mToleranceMy = Abs(toleranceMy)
    mBaseTolerance = Abs(baseTolerance)
    If mToleranceN <= 0# Then mToleranceN = 5#
    If mToleranceMx <= 0# Then mToleranceMx = 5000#
    If mToleranceMy <= 0# Then mToleranceMy = 5000#
    If mBaseTolerance <= 0# Then mBaseTolerance = DEFAULT_BASE_TOLERANCE
End Sub

' Возвращает lambda по опорной масштабируемой компоненте. Опорная компонента
' выбирается не по абсолютной величине, а по отношению к допуску своей
' размерности, чтобы N и моменты сравнивались корректно.
Public Function Lambda(ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal loadPath As CLoadPathVector) As Double
    Select Case AnchorComponent(loadPath)
        Case 1
            If Abs(loadPath.NBase) <= mBaseTolerance Then _
                Err.Raise vbObjectError + 3841, "CLoadPathMath", "Траектория нагружения не содержит ненулевой опорной компоненты N."
            Lambda = (nValue - loadPath.NOffset) / loadPath.NBase
        Case 2
            If Abs(loadPath.MxBase) <= mBaseTolerance Then _
                Err.Raise vbObjectError + 3842, "CLoadPathMath", "Траектория нагружения не содержит ненулевой опорной компоненты Mx."
            Lambda = (mxValue - loadPath.MxOffset) / loadPath.MxBase
        Case 3
            If Abs(loadPath.MyBase) <= mBaseTolerance Then _
                Err.Raise vbObjectError + 3843, "CLoadPathMath", "Траектория нагружения не содержит ненулевой опорной компоненты My."
            Lambda = (myValue - loadPath.MyOffset) / loadPath.MyBase
    End Select
End Function

' Возвращает номер опорной компоненты траектории:
' 1 - N, 2 - Mx, 3 - My. Через эту компоненту вычисляется lambda; две
' остальные компоненты дают независимые уравнения попадания на прямую нагрузки.
Public Function AnchorComponent(ByVal loadPath As CLoadPathVector) As Long
    Dim nScore As Double
    Dim mxScore As Double
    Dim myScore As Double
    nScore = Abs(loadPath.NBase) / MaxDouble(mToleranceN, 1#)
    mxScore = Abs(loadPath.MxBase) / MaxDouble(mToleranceMx, 1#)
    myScore = Abs(loadPath.MyBase) / MaxDouble(mToleranceMy, 1#)

    If mxScore >= nScore And mxScore >= myScore And Abs(loadPath.MxBase) > mBaseTolerance Then
        AnchorComponent = 2
    ElseIf myScore >= nScore And myScore >= mxScore And Abs(loadPath.MyBase) > mBaseTolerance Then
        AnchorComponent = 3
    Else
        AnchorComponent = 1
    End If
End Function

' Собирает две безразмерные невязки, которые удерживают внутренние усилия
' solver-а на прямой Target = Offset + lambda*Base. Ненулевые компоненты
' сверяются со своим target при том же lambda, нулевые Base-компоненты
' контролируются как постоянная часть Offset.
Public Sub BuildResiduals(ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double, ByRef rPath1 As Double, ByRef rPath2 As Double)
    Dim residualCount As Long
    Dim anchor As Long
    Dim targetNValue As Double
    Dim targetMxValue As Double
    Dim targetMyValue As Double

    rPath1 = 0#
    rPath2 = 0#
    anchor = AnchorComponent(loadPath)
    loadPath.Target lambdaValue, targetNValue, targetMxValue, targetMyValue

    ' Масштаб определяется заданной нагрузкой, не пробной невязкой.
    ' Иначе при нулевом target отношение M / (|M|*relativeTolerance)
    ' становится константой и численный Якобиан теряет уравнение момента.

    If Abs(loadPath.NBase) <= mBaseTolerance Then _
        AddResidual residualCount, rPath1, rPath2, _
            (nValue - loadPath.NOffset) / ComponentTolerance(1, loadPath.NOffset, loadPath.NOffset)
    If Abs(loadPath.MxBase) <= mBaseTolerance Then _
        AddResidual residualCount, rPath1, rPath2, _
            (mxValue - loadPath.MxOffset) / ComponentTolerance(2, loadPath.MxOffset, loadPath.MxOffset)
    If Abs(loadPath.MyBase) <= mBaseTolerance Then _
        AddResidual residualCount, rPath1, rPath2, _
            (myValue - loadPath.MyOffset) / ComponentTolerance(3, loadPath.MyOffset, loadPath.MyOffset)

    If residualCount < 2 And Abs(loadPath.NBase) > mBaseTolerance And anchor <> 1 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (nValue - targetNValue) / ComponentTolerance(1, targetNValue, targetNValue)
    If residualCount < 2 And Abs(loadPath.MxBase) > mBaseTolerance And anchor <> 2 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (mxValue - targetMxValue) / ComponentTolerance(2, targetMxValue, targetMxValue)
    If residualCount < 2 And Abs(loadPath.MyBase) > mBaseTolerance And anchor <> 3 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (myValue - targetMyValue) / ComponentTolerance(3, targetMyValue, targetMyValue)
End Sub

' Возвращает норму системы load-path + limit. Последний множитель сохраняет
' принятую в capacity/crack практику: предельный критерий важнее мелкой
' балансировки пути, но не должен скрывать грубую потерю направления нагрузки.
Public Function ResidualNorm(ByVal rPath1 As Double, ByVal rPath2 As Double, _
        ByVal rLimit As Double, Optional ByVal limitWeight As Double = 10#) As Double
    ResidualNorm = Abs(rPath1) + Abs(rPath2) + Abs(rLimit) * limitWeight
End Function

' Возвращает размерный допуск компоненты как максимум абсолютного допуска
' solver-а и маленькой относительной доли от масштаба текущей компоненты.
Public Function ComponentTolerance(ByVal componentIndex As Long, _
        ByVal targetValue As Double, ByVal actualValue As Double) As Double
    Select Case componentIndex
        Case 1
            ComponentTolerance = MaxDouble(mToleranceN, MaxDouble(Abs(targetValue), Abs(actualValue)) * RELATIVE_COMPONENT_TOLERANCE)
        Case 2
            ComponentTolerance = MaxDouble(mToleranceMx, MaxDouble(Abs(targetValue), Abs(actualValue)) * RELATIVE_COMPONENT_TOLERANCE)
        Case 3
            ComponentTolerance = MaxDouble(mToleranceMy, MaxDouble(Abs(targetValue), Abs(actualValue)) * RELATIVE_COMPONENT_TOLERANCE)
    End Select
End Function

' Добавляет очередную независимую невязку в компактную пару rPath1/rPath2.
Private Sub AddResidual(ByRef residualCount As Long, ByRef rPath1 As Double, _
        ByRef rPath2 As Double, ByVal residualValue As Double)
    If residualCount >= 2 Then Exit Sub
    residualCount = residualCount + 1
    If residualCount = 1 Then
        rPath1 = residualValue
    Else
        rPath2 = residualValue
    End If
End Sub

' Возвращает большее из двух чисел. Локальный helper нужен, потому что этот
' класс находится в общем слое и не должен зависеть от одноименных приватных
' функций расчетных классов.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        MaxDouble = a
    Else
        MaxDouble = b
    End If
End Function
