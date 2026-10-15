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
' Общие уравнения траектории изменения нагрузки
' ==========================================================================
' Для Target(lambda) = Offset + lambda*Base выбирает устойчивую опорную
' компоненту множителя lambda и строит две независимые невязки попадания
' на заданную траекторию. Нормирует их по допускам равновесия.
' Эта математика одинакова для несущей способности и образования трещины.

Private Const DEFAULT_BASE_TOLERANCE As Double = 0# ' Любая ненулевая часть Base масштабируется; фильтр малых моментов применяется до поиска предела.
Private Const RELATIVE_COMPONENT_TOLERANCE As Double = 0.000000001 ' Минимальная относительная добавка к допуску равновесия.

Private mToleranceN As Double ' Допуск равновесия по N, Н.
Private mToleranceMx As Double ' Допуск равновесия по Mx, Н*мм.
Private mToleranceMy As Double ' Допуск равновесия по My, Н*мм.
Private mBaseTolerance As Double ' Необязательный порог для программного вызова; рабочий расчёт проверяет точное отличие Base от нуля.
Private mNumericalFrame As CSectionNumericalFrame ' Начало координат и масштабы, используемые для устойчивого численного решения.
Private mNumericalPath As CLoadPathVector, mSourcePath As CLoadPathVector ' Траектории нагрузки в численных и исходных координатах соответственно.
Private mPathSection As CSectionModel, mPathRevision As Long ' Модель и её версия, для которых подготовлена траектория нагрузки.
Private mSourceTolerances(1 To 3) As Double ' Исходные допуски равновесия по N, Mx и My: Н, Н*мм, Н*мм.
Private mCommonNumericalScale As Boolean, mNumericalReady As Boolean ' Признак общего масштаба уравнений и признак готовности численного представления.
Private mForceScale As Double, mPathTolerance As Double ' Масштаб усилий и допуск отклонения от траектории в численных координатах.

' Задаёт начальные допуски равновесия, согласованные с начальными настройками решателя НДС.
Private Sub Class_Initialize()
    mToleranceN = 5#
    mToleranceMx = 5000#
    mToleranceMy = 5000#
    mBaseTolerance = DEFAULT_BASE_TOLERANCE
End Sub

' Проверяет, является ли заданное изменение нагрузки осевым относительно выбранной точки.
' Учитывает только погрешность переноса усилий Double; способ поиска
' и инженерный порог малых моментов этим методом не задаются.
Public Function IsPureAxialLoadPath(ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, ByVal myOffset As Double, ByVal myBase As Double, _
        ByVal referenceX As Double, ByVal referenceY As Double) As Boolean
    If Abs(nBase) <= 0.000000001 Or Abs(nOffset) > 0.000000001 Then Exit Function
    IsPureAxialLoadPath = IsZeroMomentAboutPoint(nOffset, mxOffset, referenceY) And _
        IsZeroMomentAboutPoint(nBase, mxBase, referenceY) And _
        IsZeroMomentAboutPoint(nOffset, myOffset, referenceX) And _
        IsZeroMomentAboutPoint(nBase, myBase, referenceX)
End Function

' Проверяет оба момента относительно заданной точки приложения с численным масштабом задачи.
Private Function IsZeroMomentAboutPoint(ByVal n As Double, ByVal moment As Double, ByVal coordinate As Double) As Boolean
    Dim transfer As Double, roundoff As Double
    transfer = n * coordinate
    ' Восемь машинных epsilon Double для произведения и разности моментов.
    roundoff = 0.000000001 + 0.00000000000000177635683940025 * (Abs(moment) + Abs(transfer))
    IsZeroMomentAboutPoint = Abs(moment - transfer) <= roundoff
End Function

' Принимает допуски решателя для нормирования уравнений траектории нагрузки.
' Разные проверки могут использовать разные допуски при одной формуле.
' Допуск равновесия не обнуляет заданную ненулевую масштабируемую нагрузку.
Public Sub Configure(ByVal toleranceN As Double, ByVal toleranceMx As Double, _
        ByVal toleranceMy As Double, Optional ByVal baseTolerance As Double = DEFAULT_BASE_TOLERANCE)
    mNumericalReady = False
    mToleranceN = Abs(toleranceN)
    mToleranceMx = Abs(toleranceMx)
    mToleranceMy = Abs(toleranceMy)
    mBaseTolerance = Abs(baseTolerance)
    If mToleranceN <= 0# Then mToleranceN = 5#
    If mToleranceMx <= 0# Then mToleranceMx = 5000#
    If mToleranceMy <= 0# Then mToleranceMy = 5000#
End Sub

' Подготавливает нагрузку относительно центра бетонной сетки для общих поисковых задач.
' Нормирует моменты по размеру сечения, сохраняя исходные допуски равновесия.
' Точное совпадение проверяется по значениям, а не по ссылке на изменяемый объект нагрузки.
Friend Sub PrepareNumericalPath(ByVal section As CSectionModel, ByVal loadPath As CLoadPathVector, _
        ByVal toleranceN As Double, ByVal toleranceMx As Double, ByVal toleranceMy As Double, _
        Optional ByVal commonForceScale As Boolean = False)
    ' Технический журнал сохраняет стандартный текст VBA; потребитель дает понятное пояснение.
    If section Is Nothing Or loadPath Is Nothing Then Err.Raise 91, "CLoadPathMath"
    If NumericalPathMatches(section, loadPath, toleranceN, toleranceMx, toleranceMy, commonForceScale) Then Exit Sub
    mNumericalReady = False
    If mNumericalFrame Is Nothing Then Set mNumericalFrame = New CSectionNumericalFrame
    mNumericalFrame.Initialize section
    Set mNumericalPath = New CLoadPathVector
    mNumericalPath.Initialize loadPath.NOffset, loadPath.NBase, _
        mNumericalFrame.CenterMx(loadPath.NOffset, loadPath.MxOffset) / mNumericalFrame.LengthScale, _
        mNumericalFrame.CenterMx(loadPath.NBase, loadPath.MxBase) / mNumericalFrame.LengthScale, _
        mNumericalFrame.CenterMy(loadPath.NOffset, loadPath.MyOffset) / mNumericalFrame.LengthScale, _
        mNumericalFrame.CenterMy(loadPath.NBase, loadPath.MyBase) / mNumericalFrame.LengthScale
    mForceScale = MaxDouble(1#, MaxDouble(Abs(mNumericalPath.NOffset), _
        MaxDouble(Abs(mNumericalPath.NBase), MaxDouble(Abs(mNumericalPath.MxOffset), _
        MaxDouble(Abs(mNumericalPath.MxBase), MaxDouble(Abs(mNumericalPath.MyOffset), Abs(mNumericalPath.MyBase)))))))
    If commonForceScale Then
        Configure mForceScale, mForceScale, mForceScale
    Else
        Configure toleranceN, toleranceMx / mNumericalFrame.LengthScale, toleranceMy / mNumericalFrame.LengthScale
    End If
    mPathTolerance = MinDouble(toleranceN, MinDouble(toleranceMx, toleranceMy) / mNumericalFrame.LengthScale) / mForceScale
    Set mSourcePath = New CLoadPathVector
    mSourcePath.Initialize loadPath.NOffset, loadPath.NBase, loadPath.MxOffset, loadPath.MxBase, loadPath.MyOffset, loadPath.MyBase
    Set mPathSection = section: mPathRevision = section.Revision
    mSourceTolerances(1) = toleranceN: mSourceTolerances(2) = toleranceMx: mSourceTolerances(3) = toleranceMy
    mCommonNumericalScale = commonForceScale: mNumericalReady = True
End Sub

' Проверяет принадлежность состояния заданной нагрузочной траектории в нормированной системе
' координат.
Private Function NumericalPathMatches(ByVal section As CSectionModel, ByVal loadPath As CLoadPathVector, _
        ByVal toleranceN As Double, ByVal toleranceMx As Double, ByVal toleranceMy As Double, _
        ByVal commonForceScale As Boolean) As Boolean
    If Not mNumericalReady Then Exit Function
    If Not section Is mPathSection Then Exit Function
    If section.Revision <> mPathRevision Or commonForceScale <> mCommonNumericalScale Then Exit Function
    NumericalPathMatches = loadPath.NOffset = mSourcePath.NOffset And loadPath.NBase = mSourcePath.NBase And _
        loadPath.MxOffset = mSourcePath.MxOffset And loadPath.MxBase = mSourcePath.MxBase And _
        loadPath.MyOffset = mSourcePath.MyOffset And loadPath.MyBase = mSourcePath.MyBase And _
        toleranceN = mSourceTolerances(1) And toleranceMx = mSourceTolerances(2) And toleranceMy = mSourceTolerances(3)
End Function

' Берет центрированные интегралы напрямую: вычитание N*координата из уже
' восстановленного большого момента потеряло бы точность удаленного сечения.
Friend Sub BuildNumericalResiduals(ByVal solver As CSectionSolver, ByRef lambdaValue As Double, _
        ByRef rPath1 As Double, ByRef rPath2 As Double)
    RequireNumericalPath
    Dim n As Double, mx As Double, my As Double, residualScale As Double
    solver.NumericalEfforts n, mx, my
    mx = mx / mNumericalFrame.LengthScale: my = my / mNumericalFrame.LengthScale
    lambdaValue = Lambda(n, mx, my, mNumericalPath)
    If mCommonNumericalScale Then residualScale = mForceScale
    BuildResiduals n, mx, my, mNumericalPath, lambdaValue, rPath1, rPath2, residualScale
End Sub

' Определяет множитель нагрузки по нормированной траектории; не подменяет проверку равновесия.
Friend Function NumericalLambda(ByVal solver As CSectionSolver) As Double
    RequireNumericalPath
    Dim n As Double, mx As Double, my As Double
    solver.NumericalEfforts n, mx, my
    NumericalLambda = Lambda(n, mx / mNumericalFrame.LengthScale, my / mNumericalFrame.LengthScale, mNumericalPath)
End Function

' Отклоняет вызов до подготовки численной траектории и масштабов усилий.
Private Sub RequireNumericalPath()
    If Not mNumericalReady Then _
        Err.Raise vbObjectError + 3845, "CLoadPathMath", "Численная траектория нагружения не подготовлена."
End Sub

' Подготовленная численная система координат и масштабов усилий.
Friend Property Get NumericalFrame() As CSectionNumericalFrame
    RequireNumericalPath
    Set NumericalFrame = mNumericalFrame
End Property

' Положительный масштаб осевой силы для численной траектории, Н.
Friend Property Get NumericalForceScale() As Double
    RequireNumericalPath
    NumericalForceScale = mForceScale
End Property

' Допуск совпадения состояния с нормированной нагрузочной траекторией.
Friend Property Get NumericalPathTolerance() As Double
    RequireNumericalPath
    NumericalPathTolerance = mPathTolerance
End Property

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

' Вычисляет две безразмерные невязки, проверяющие соответствие внутренних усилий
' зависимости Target = Offset + lambda*Base. Масштабируемые компоненты
' проверяются при одном lambda; постоянные — относительно своих значений Offset.
Public Sub BuildResiduals(ByVal nValue As Double, ByVal mxValue As Double, _
        ByVal myValue As Double, ByVal loadPath As CLoadPathVector, _
        ByVal lambdaValue As Double, ByRef rPath1 As Double, ByRef rPath2 As Double, _
        Optional ByVal commonResidualScale As Double = 0#)
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
            (nValue - loadPath.NOffset) / ResidualScale(1, loadPath.NOffset, commonResidualScale)
    If Abs(loadPath.MxBase) <= mBaseTolerance Then _
        AddResidual residualCount, rPath1, rPath2, _
            (mxValue - loadPath.MxOffset) / ResidualScale(2, loadPath.MxOffset, commonResidualScale)
    If Abs(loadPath.MyBase) <= mBaseTolerance Then _
        AddResidual residualCount, rPath1, rPath2, _
            (myValue - loadPath.MyOffset) / ResidualScale(3, loadPath.MyOffset, commonResidualScale)

    If residualCount < 2 And Abs(loadPath.NBase) > mBaseTolerance And anchor <> 1 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (nValue - targetNValue) / ResidualScale(1, targetNValue, commonResidualScale)
    If residualCount < 2 And Abs(loadPath.MxBase) > mBaseTolerance And anchor <> 2 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (mxValue - targetMxValue) / ResidualScale(2, targetMxValue, commonResidualScale)
    If residualCount < 2 And Abs(loadPath.MyBase) > mBaseTolerance And anchor <> 3 Then _
        AddResidual residualCount, rPath1, rPath2, _
            (myValue - targetMyValue) / ResidualScale(3, targetMyValue, commonResidualScale)
End Sub

' В численном пространстве N/Mx/L/My/L все компоненты имеют размерность силы.
' Общий масштаб отделяет направление шага от допусков окончательной приемки.
Private Function ResidualScale(ByVal componentIndex As Long, ByVal target As Double, _
        ByVal commonScale As Double) As Double
    If commonScale > 0# Then
        ResidualScale = commonScale
    Else
        ResidualScale = ComponentTolerance(componentIndex, target, target)
    End If
End Function

' Вычисляет общую меру отклонения от траектории нагрузки и от предельного критерия.
' Критерию придаётся принятый вес, но большое отклонение направления нагрузки
' не должно скрываться малой невязкой деформаций.
Public Function ResidualNorm(ByVal rPath1 As Double, ByVal rPath2 As Double, _
        ByVal rLimit As Double, Optional ByVal limitWeight As Double = 10#) As Double
    ResidualNorm = Abs(rPath1) + Abs(rPath2) + Abs(rLimit) * limitWeight
End Function

' Определяет размерный допуск компоненты усилий как максимум
' абсолютного допуска равновесия и малой относительной доли её масштаба.
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

' Возвращает меньшее из двух значений Double без изменения размерности.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
