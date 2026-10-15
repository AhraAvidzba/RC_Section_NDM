VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStateGuessBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Начальное приближение поля деформаций сечения
' ==========================================================================
' Использует ту же модель и диаграммы материалов, что основной расчёт.
' Строит упругую оценку, уточняет работающие участки диаграмм и при необходимости
' подбирает выход из плато. Полученная плоскость служит только начальным приближением:
' равновесие и физическую допустимость проверяет CSectionSolver.

Private mFrame As CSectionNumericalFrame ' Общая центрированная система стартовых оценок.

Private Const STATE_NEAR_ZERO As Double = 0.000000000001 ' Допуск для почти нулевой стартовой плоскости.
Private Const EXTENSION_START_FACTOR As Double = 1.02 ' Небольшой заход за физический eps_ult для старта повторного solve в extension.
Private Const EXTENSION_GUESS_ATTEMPTS As Long = 3 ' Количество независимых стартовых оценок для повторного StateSolution.
Private Const EXTENSION_ACTIVE_SET_ITERATIONS As Long = 8 ' Сколько раз уточнять кусочно-линейный набор участков диаграмм.
Private Const ACTIVE_SET_EPSILON_TOLERANCE As Double = 0.0000000001 ' Допуск изменения epsilon0, при котором уточнение работающих участков диаграмм можно завершить.
Private Const LIMIT_GUESS_UTILIZATION As Double = 0.9 ' Старт немного внутри первого предела, вне вырожденного плато.

' Стартовые оценки используют ту же систему, что и основные итерации.
' Наружу возвращается epsilon0 в исходном начале координат; модель не сдвигается.
Public Function BuildForDirectState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If section Is Nothing Then Exit Function
    InitializeFrame section
    BuildForDirectState = BuildCenteredDirectState(section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, eps0, kx, ky)
    If BuildForDirectState Then eps0 = mFrame.OriginalStrain(eps0, kx, ky)
End Function

' Уменьшает только начальное поле деформаций поиска предела. Расчётные усилия
' не меняет; коэффициент использования всех физических пределов передаёт задача.
Public Sub ScaleToPhysicalLimit(ByVal section As CSectionModel, ByVal utilization As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double)
    If utilization <= 0# Then Exit Sub
    InitializeFrame section
    Dim strainScale As Double, centerStrain As Double
    strainScale = LIMIT_GUESS_UTILIZATION / utilization
    centerStrain = mFrame.CenterStrain(eps0, kx, ky)
    kx = kx * strainScale: ky = ky * strainScale
    eps0 = mFrame.OriginalStrain(centerStrain * strainScale, kx, ky)
End Sub

' Повторные начальные приближения задаются в тех же исходных координатах сечения.
Public Function BuildAttemptForExtensionState(ByVal attemptIndex As Long, _
        ByVal section As CSectionModel, ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If section Is Nothing Then Exit Function
    If concreteMaterial Is Nothing Or steelMaterial Is Nothing Then Exit Function
    InitializeFrame section
    BuildAttemptForExtensionState = BuildCenteredExtensionAttempt(attemptIndex, section, failedService, _
        concreteMaterial, steelMaterial, targetN, targetMx, targetMy, eps0, kx, ky)
    If BuildAttemptForExtensionState Then eps0 = mFrame.OriginalStrain(eps0, kx, ky)
End Function

' Готовит численную систему координат и масштабов для начального приближения НДС.
Private Sub InitializeFrame(ByVal section As CSectionModel)
    If mFrame Is Nothing Then Set mFrame = New CSectionNumericalFrame
    mFrame.Initialize section
End Sub

' Строит начальное поле для обычного НДС. Для чистого изгиба старт из нуля
' попадает в излом диаграммы бетона. Сначала получает упругую оценку,
' затем уточняет, какие участки диаграмм работают в элементах.
' Это только начальное приближение; равновесие подтверждает CSectionSolver.
Private Function BuildCenteredDirectState(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If section Is Nothing Then Exit Function
    If concreteMaterial Is Nothing Then Exit Function
    If steelMaterial Is Nothing Then Exit Function
    If Abs(targetN) + Abs(targetMx) + Abs(targetMy) <= STATE_NEAR_ZERO Then Exit Function

    Dim trialEps0 As Double
    Dim trialKx As Double
    Dim trialKy As Double
    If Not TryBuildNearZeroElasticPlane(section, concreteMaterial, steelMaterial, _
            targetN, targetMx, targetMy, trialEps0, trialKx, trialKy) Then
        Exit Function
    End If

    BuildCenteredDirectState = TryBuildFromTrialPlane(section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, trialEps0, trialKx, trialKy, eps0, kx, ky, False)
    If Not BuildCenteredDirectState Then
        ' Если уточнение работающих участков дало вырожденную систему, сохраняем
        ' ранее полученную упругую оценку как старт. Равновесие проверит основной решатель.
        eps0 = trialEps0
        kx = trialKx
        ky = trialKy
        BuildCenteredDirectState = True
    End If
End Function

' Возвращает число независимых вариантов стартовой плоскости.
' CStateSolutionRunner использует это, чтобы не останавливаться на первом варианте,
' если сам повторный solve на нем численно не сошелся.
Public Property Get ExtensionAttemptCount() As Long
    ExtensionAttemptCount = EXTENSION_GUESS_ATTEMPTS
End Property

' Дает короткое имя варианта старта для диагностического txt-отчета.
Public Function ExtensionAttemptName(ByVal attemptIndex As Long) As String
    Select Case attemptIndex
        Case 1: ExtensionAttemptName = "равновесие по продолженной диаграмме"
        Case 2: ExtensionAttemptName = "активная зона несошедшегося состояния"
        Case 3: ExtensionAttemptName = "равномерный сдвиг в продолжение диаграммы"
        Case Else: ExtensionAttemptName = "неизвестное приближение"
    End Select
End Function

' Готовит очередной вариант начального поля для повторного решения.
' Пробует способы от более точных к более простым, не меняя
' диаграмм материалов и физической постановки задачи.
Private Function BuildCenteredExtensionAttempt(ByVal attemptIndex As Long, _
        ByVal section As CSectionModel, ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    Select Case attemptIndex
        Case 1
            BuildCenteredExtensionAttempt = TryBuildGlobalExtensionEquilibrium(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, eps0, kx, ky)
        Case 2
            BuildCenteredExtensionAttempt = TryBuildFromFailedState(section, failedService, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, eps0, kx, ky)
        Case 3
            BuildCenteredExtensionAttempt = TryBuildUniformExtensionShift(failedService, concreteMaterial, steelMaterial, _
                targetN, eps0, kx, ky)

    End Select
End Function

' Использует последнее поле неудачной попытки как оценку формы изгиба.
' Пересобирает линейную оценку по элементам, которые при этой форме
' могут попасть на вспомогательное продолжение диаграмм.
Private Function TryBuildFromFailedState(ByVal section As CSectionModel, _
        ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    Dim trialEps0 As Double
    Dim trialKx As Double
    Dim trialKy As Double
    If Not BuildExtensionTrialPlane(failedService, concreteMaterial, steelMaterial, _
            trialEps0, trialKx, trialKy) Then
        Exit Function
    End If

    TryBuildFromFailedState = TryBuildFromTrialPlane(section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, trialEps0, trialKx, trialKy, eps0, kx, ky, True)
End Function

' Собирает линейную систему по полной касательной модели в пробной плоскости.
' Для изгиба важно учитывать не только элементы с ненулевой касательной, но и
' постоянные усилия от плато/обычных участков диаграммы. Иначе правая часть
' равновесия получается искаженной, и старт может быть хуже, чем сама первая
' несошедшаяся попытка.
Private Function TryBuildFromTrialPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal trialEps0 As Double, ByVal trialKx As Double, ByVal trialKy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double, _
        Optional ByVal requireExtension As Boolean = True, _
        Optional ByVal requireStable As Boolean = False) As Boolean
    Dim currentEps0 As Double
    Dim currentKx As Double
    Dim currentKy As Double
    Dim nextEps0 As Double
    Dim nextKx As Double
    Dim nextKy As Double
    Dim iter As Long
    Dim solvedAtLeastOnce As Boolean
    Dim stable As Boolean ' Прекратилась смена активных участков и плоскости.

    currentEps0 = trialEps0
    currentKx = trialKx
    currentKy = trialKy

    ' После линейного решения определяем работающие участки диаграмм и заново
    ' собираем систему по ним. Если набор участков перестал меняться,
    ' полученное поле можно передать как старт основному решателю.
    ' Правила одинаковы для всех форм сечения.
    For iter = 1 To EXTENSION_ACTIVE_SET_ITERATIONS
        If Not TrySolveLinearizedTrialPlane(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, currentEps0, currentKx, currentKy, _
                nextEps0, nextKx, nextKy) Then
            Exit For
        End If

        eps0 = nextEps0
        kx = nextKx
        ky = nextKy
        solvedAtLeastOnce = True

        If ExtensionPlaneClose(currentEps0, currentKx, currentKy, nextEps0, nextKx, nextKy) Then
            stable = True
            Exit For
        End If

        currentEps0 = nextEps0
        currentKx = nextKx
        currentKy = nextKy
    Next iter

    If requireStable And Not stable Then Exit Function
    If solvedAtLeastOnce Then
        If requireExtension Then
            TryBuildFromTrialPlane = CandidateUsesExtension(section, concreteMaterial, steelMaterial, eps0, kx, ky)
        Else
            TryBuildFromTrialPlane = Abs(eps0) + mFrame.LengthScale * (Abs(kx) + Abs(ky)) > STATE_NEAR_ZERO
        End If
    End If
End Function

' Решает одну линейную задачу равновесия относительно заданной пробной
' плоскости деформаций. Здесь нет итераций Newton: мы только берем текущие
' касательные участки диаграмм и представляем sigma(epsilon) как
' E_t * epsilon + intercept для каждого элемента.
Private Function TrySolveLinearizedTrialPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal trialEps0 As Double, ByVal trialKx As Double, ByVal trialKy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    Dim k11 As Double, k12 As Double, k13 As Double
    Dim k21 As Double, k22 As Double, k23 As Double
    Dim k31 As Double, k32 As Double, k33 As Double
    Dim n0 As Double, mx0 As Double, my0 As Double
    Dim activeCount As Long ' Сколько элементов реально попало в линейную extension-оценку.
    Dim i As Long
    Dim strain As Double

    ' Проходим по бетону полностью. В сжатой зоне может работать extension,
    ' на физическом плато будет только постоянное усилие, а растянутый бетон
    ' с Ignore даст нулевой вклад сам через диаграмму.
    For i = 1 To section.ConcreteCount
        strain = PlaneStrain(trialEps0, trialKx, trialKy, section.ConcreteX(i), section.ConcreteY(i))
        AccumulateMaterialElement concreteMaterial, strain, section.ConcreteArea(i), _
            section.ConcreteX(i), section.ConcreteY(i), activeCount, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
    Next i

    ' Учитываем все участки арматурной диаграммы: упругую жёсткость, постоянное
    ' усилие плато и малую жёсткость численного продолжения. Вклад стержня
    ' равен «сталь минус вытесненный бетон», поскольку бетонная сетка
    ' уже включает площадь, занятую арматурой.
    For i = 1 To section.RebarCount
        strain = PlaneStrain(trialEps0, trialKx, trialKy, section.RebarX(i), section.RebarY(i))
        AccumulateRebarElement concreteMaterial, steelMaterial, strain, section.RebarArea(i), _
            section.RebarX(i), section.RebarY(i), activeCount, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
    Next i

    If Not SolveLinearizedSystem(activeCount, targetN, targetMx, targetMy, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0, eps0, kx, ky) Then
        Exit Function
    End If

    TrySolveLinearizedTrialPlane = True
End Function

' Строит первую упругую оценку около нулевой деформации. Для бетона с отключённым
' растяжением берёт небольшую сжатую пробу, чтобы не получить нулевую жёсткость.
' Следующее уточнение исключает элементы, попавшие в растяжение.
Private Function TryBuildNearZeroElasticPlane(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    Dim k11 As Double, k12 As Double, k13 As Double
    Dim k21 As Double, k22 As Double, k23 As Double
    Dim k31 As Double, k32 As Double, k33 As Double
    Dim n0 As Double, mx0 As Double, my0 As Double
    Dim activeCount As Long
    Dim i As Long
    Dim sampleStrain As Double

    sampleStrain = -STATE_NEAR_ZERO
    For i = 1 To section.ConcreteCount
        AccumulateMaterialElement concreteMaterial, sampleStrain, section.ConcreteArea(i), _
            section.ConcreteX(i), section.ConcreteY(i), activeCount, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
    Next i

    For i = 1 To section.RebarCount
        AccumulateRebarElement concreteMaterial, steelMaterial, sampleStrain, section.RebarArea(i), _
            section.RebarX(i), section.RebarY(i), activeCount, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
    Next i

    TryBuildNearZeroElasticPlane = SolveLinearizedSystem(activeCount, targetN, targetMx, targetMy, _
        k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0, eps0, kx, ky)
End Function

' Строит грубую глобальную оценку по знаку продольной силы.
' Для растяжения работает арматура в tension-extension. Для сжатия работают
' бетонные волокна в compression-extension и сжатая арматура. Моменты Mx/My
' все равно входят в правую часть системы, поэтому оценка не ограничена только
' чистой осевой силой.
Private Function TryBuildGlobalExtensionEquilibrium(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If Abs(targetN) <= STATE_NEAR_ZERO Then Exit Function

    Dim k11 As Double, k12 As Double, k13 As Double
    Dim k21 As Double, k22 As Double, k23 As Double
    Dim k31 As Double, k32 As Double, k33 As Double
    Dim n0 As Double, mx0 As Double, my0 As Double
    Dim activeCount As Long
    Dim i As Long
    Dim sampleStrain As Double

    If targetN > 0# Then
        If Not TryTensionExtensionSample(steelMaterial, sampleStrain) Then Exit Function
        For i = 1 To section.RebarCount
            AccumulateRebarElement concreteMaterial, steelMaterial, sampleStrain, section.RebarArea(i), _
                section.RebarX(i), section.RebarY(i), activeCount, _
                k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
        Next i
    Else
        If TryCompressionExtensionSample(concreteMaterial, sampleStrain) Then
            For i = 1 To section.ConcreteCount
                AccumulateMaterialElement concreteMaterial, sampleStrain, section.ConcreteArea(i), _
                    section.ConcreteX(i), section.ConcreteY(i), activeCount, _
                    k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
            Next i
        End If
        If TryCompressionExtensionSample(steelMaterial, sampleStrain) Then
            For i = 1 To section.RebarCount
                AccumulateRebarElement concreteMaterial, steelMaterial, sampleStrain, section.RebarArea(i), _
                    section.RebarX(i), section.RebarY(i), activeCount, _
                    k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
            Next i
        End If
    End If

    If Not SolveLinearizedSystem(activeCount, targetN, targetMx, targetMy, _
            k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0, eps0, kx, ky) Then
        Exit Function
    End If

    Dim seedEps0 As Double
    Dim seedKx As Double
    Dim seedKy As Double
    seedEps0 = eps0
    seedKx = kx
    seedKy = ky

    If targetN > 0# Then PushRebarsIntoTensionExtension section, steelMaterial, seedEps0, seedKx, seedKy
    TryBuildGlobalExtensionEquilibrium = TryBuildFromTrialPlane(section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, seedEps0, seedKx, seedKy, eps0, kx, ky, True)
End Function

' Смещает равномерную деформацию epsilon0 к ближайшему продолжению диаграммы,
' сохраняя кривизны. Это даёт другой старт, если форма изгиба уже определена,
' но первая попытка остановилась на плато.
Private Function TryBuildUniformExtensionShift(ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If failedService Is Nothing Then Exit Function

    Dim shift As Double
    If targetN > 0# And steelMaterial.HasTensionExtension And _
            steelMaterial.PhysicalTensionStrain > 0# Then
        shift = steelMaterial.PhysicalTensionStrain * EXTENSION_START_FACTOR - failedService.MaxSteelStrain
    ElseIf targetN < 0# And concreteMaterial.HasCompressionExtension And _
            concreteMaterial.PhysicalCompressionStrain < 0# Then
        shift = concreteMaterial.PhysicalCompressionStrain * EXTENSION_START_FACTOR - failedService.MinConcreteStrain
    ElseIf targetN < 0# And steelMaterial.HasCompressionExtension And _
            steelMaterial.PhysicalCompressionStrain < 0# Then
        shift = steelMaterial.PhysicalCompressionStrain * EXTENSION_START_FACTOR - failedService.MinSteelStrain
    Else
        Exit Function
    End If

    If Abs(shift) <= STATE_NEAR_ZERO Then Exit Function

    eps0 = mFrame.CenterStrain(failedService.Epsilon0, failedService.KappaX, failedService.KappaY) + shift
    kx = failedService.KappaX
    ky = failedService.KappaY
    TryBuildUniformExtensionShift = True
End Function

' Готовит пробное поле из неудачного НДС: использует имеющийся выход
' за физический предел или немного увеличивает поле до такого выхода.
' Равновесие не решает; только выбирает элементы на продолжении диаграммы для нового старта.
Private Function BuildExtensionTrialPlane(ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If failedService Is Nothing Then Exit Function

    Dim maxAbsState As Double
    maxAbsState = MaxDouble(Abs(mFrame.CenterStrain(failedService.Epsilon0, failedService.KappaX, failedService.KappaY)), _
        mFrame.LengthScale * MaxDouble(Abs(failedService.KappaX), Abs(failedService.KappaY)))
    If maxAbsState <= STATE_NEAR_ZERO Then Exit Function

    If FailedStateAlreadyInExtension(failedService, concreteMaterial, steelMaterial) Then
        eps0 = mFrame.CenterStrain(failedService.Epsilon0, failedService.KappaX, failedService.KappaY)
        kx = failedService.KappaX
        ky = failedService.KappaY
        BuildExtensionTrialPlane = True
        Exit Function
    End If

    Dim scaleFactor As Double
    scaleFactor = 1#
    If steelMaterial.HasTensionExtension Then
        scaleFactor = MaxDouble(scaleFactor, ExtensionScaleForPositiveLimit( _
            failedService.MaxSteelStrain, steelMaterial.PhysicalTensionStrain))
    End If
    If steelMaterial.HasCompressionExtension Then
        scaleFactor = MaxDouble(scaleFactor, ExtensionScaleForNegativeLimit( _
            failedService.MinSteelStrain, steelMaterial.PhysicalCompressionStrain))
    End If
    If concreteMaterial.HasCompressionExtension Then
        scaleFactor = MaxDouble(scaleFactor, ExtensionScaleForNegativeLimit( _
            failedService.MinConcreteStrain, concreteMaterial.PhysicalCompressionStrain))
    End If
    If scaleFactor <= 1# Then Exit Function

    eps0 = mFrame.CenterStrain(failedService.Epsilon0, failedService.KappaX, failedService.KappaY) * scaleFactor
    kx = failedService.KappaX * scaleFactor
    ky = failedService.KappaY * scaleFactor
    BuildExtensionTrialPlane = True
End Function

' Добавляет вклад одного элемента в линейную систему равновесия.
' В любой кусочно-линейной точке sigma можно представить как
' sigma = E_t * epsilon + intercept. Если E_t = 0 на плато, элемент не добавляет
' жесткость в матрицу, но его постоянное усилие обязательно остается в правой
' части через n0/mx0/my0.
Private Sub AccumulateMaterialElement(ByVal material As CMaterialDiagram, _
        ByVal sampleStrain As Double, ByVal area As Double, ByVal x As Double, ByVal y As Double, _
        ByRef activeCount As Long, _
        ByRef k11 As Double, ByRef k12 As Double, ByRef k13 As Double, _
        ByRef k21 As Double, ByRef k22 As Double, ByRef k23 As Double, _
        ByRef k31 As Double, ByRef k32 As Double, ByRef k33 As Double, _
        ByRef n0 As Double, ByRef mx0 As Double, ByRef my0 As Double)
    Dim tangent As Double ' Касательная жесткость текущего участка диаграммы, МПа.
    Dim stress As Double ' Напряжение материала в точке линеаризации, МПа.
    material.EvaluateAtStrain sampleStrain, stress, tangent

    AccumulateLinearElement tangent, stress, sampleStrain, area, x, y, activeCount, _
        k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
End Sub

' Добавляет в систему один уже линеаризованный элемент. На вход приходит
' касательная E_t и напряжение sigma в sample-точке; дальше элемент дает
' жесткость E_t*A и постоянную часть от intercept = sigma - E_t*epsilon.
Private Sub AccumulateLinearElement(ByVal tangent As Double, ByVal stress As Double, ByVal sampleStrain As Double, _
        ByVal area As Double, ByVal x As Double, ByVal y As Double, _
        ByRef activeCount As Long, _
        ByRef k11 As Double, ByRef k12 As Double, ByRef k13 As Double, _
        ByRef k21 As Double, ByRef k22 As Double, ByRef k23 As Double, _
        ByRef k31 As Double, ByRef k32 As Double, ByRef k33 As Double, _
        ByRef n0 As Double, ByRef mx0 As Double, ByRef my0 As Double)
    Dim intercept As Double ' Свободный член линейной зависимости sigma(epsilon).
    Dim ea As Double ' Осевой коэффициент жесткости элемента E_t * A.
    Dim force0 As Double ' Постоянная часть усилия элемента от intercept * A.
    x = (x - mFrame.OriginX) / mFrame.LengthScale
    y = (y - mFrame.OriginY) / mFrame.LengthScale
    intercept = stress - tangent * sampleStrain
    ea = tangent * area
    force0 = intercept * area

    ' Замещение бетона арматурой может дать отрицательный вклад касательной.
    ' Он должен остаться в матрице, как при интегрировании CSectionSolver.
    If tangent <> 0# Then
        k11 = k11 + ea
        k12 = k12 + ea * y
        k13 = k13 + ea * x
        k21 = k21 + ea * y
        k22 = k22 + ea * y * y
        k23 = k23 + ea * y * x
        k31 = k31 + ea * x
        k32 = k32 + ea * x * y
        k33 = k33 + ea * x * x
        activeCount = activeCount + 1
    End If

    n0 = n0 + force0
    mx0 = mx0 + force0 * y
    my0 = my0 + force0 * x
End Sub

' Добавляет вклад арматурного стержня так же, как это делает CSectionSolver:
' из диаграммы стали вычитается диаграмма бетона в той же точке. Это нужно,
' потому что бетонные конечные элементы уже интегрируются по полной сетке, а
' площадь стержня заменяет бетон сталью, а не добавляется поверх бетона.
Private Sub AccumulateRebarElement(ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal sampleStrain As Double, ByVal area As Double, ByVal x As Double, ByVal y As Double, _
        ByRef activeCount As Long, _
        ByRef k11 As Double, ByRef k12 As Double, ByRef k13 As Double, _
        ByRef k21 As Double, ByRef k22 As Double, ByRef k23 As Double, _
        ByRef k31 As Double, ByRef k32 As Double, ByRef k33 As Double, _
        ByRef n0 As Double, ByRef mx0 As Double, ByRef my0 As Double)
    Dim steelStress As Double ' Напряжение стали в точке линеаризации.
    Dim concreteStress As Double ' Напряжение вытесненного бетона на той же деформации.
    Dim steelTangent As Double ' Касательный модуль стали на текущем участке.
    Dim concreteTangent As Double ' Касательный модуль вытесненного бетона.
    steelMaterial.EvaluateAtStrain sampleStrain, steelStress, steelTangent
    concreteMaterial.EvaluateAtStrain sampleStrain, concreteStress, concreteTangent

    AccumulateLinearElement steelTangent - concreteTangent, _
        steelStress - concreteStress, sampleStrain, area, x, y, activeCount, _
        k11, k12, k13, k21, k22, k23, k31, k32, k33, n0, mx0, my0
End Sub

' Решает линейную систему по выбранным работающим участкам диаграмм.
' Если работающих элементов нет или матрица вырождена,
' позволяет попробовать следующий способ начального приближения.
Private Function SolveLinearizedSystem(ByVal activeCount As Long, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal k11 As Double, ByVal k12 As Double, ByVal k13 As Double, _
        ByVal k21 As Double, ByVal k22 As Double, ByVal k23 As Double, _
        ByVal k31 As Double, ByVal k32 As Double, ByVal k33 As Double, _
        ByVal n0 As Double, ByVal mx0 As Double, ByVal my0 As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If activeCount <= 0 Then Exit Function

    Dim system As CLinearSystem3x3
    Set system = New CLinearSystem3x3
    If Not system.Solve(k11, k12, k13, k21, k22, k23, k31, k32, k33, _
            targetN - n0, mFrame.CenterMx(targetN, targetMx) / mFrame.LengthScale - mx0, _
            mFrame.CenterMy(targetN, targetMy) / mFrame.LengthScale - my0) Then
        Exit Function
    End If

    eps0 = system.X1
    kx = system.X2 / mFrame.LengthScale
    ky = system.X3 / mFrame.LengthScale
    SolveLinearizedSystem = (Abs(eps0) + mFrame.LengthScale * (Abs(kx) + Abs(ky)) > STATE_NEAR_ZERO)
End Function

' Немного увеличивает растянутую оценку, чтобы хотя бы крайние стержни
' вышли с плато на участок с ненулевым наклоном. Это даёт рабочий старт
' осевого растяжения, но не заменяет последующее решение равновесия.
Private Sub PushRebarsIntoTensionExtension(ByVal section As CSectionModel, _
        ByVal steelMaterial As CMaterialDiagram, ByRef eps0 As Double, ByVal kx As Double, ByVal ky As Double)
    If section.RebarCount <= 0 Then Exit Sub
    If Not steelMaterial.HasTensionExtension Then Exit Sub
    If steelMaterial.PhysicalTensionStrain <= 0# Then Exit Sub

    Dim minRebarStrain As Double
    minRebarStrain = 1E+100

    Dim i As Long
    For i = 1 To section.RebarCount
        minRebarStrain = MinDouble(minRebarStrain, _
            PlaneStrain(eps0, kx, ky, section.RebarX(i), section.RebarY(i)))
    Next i

    If minRebarStrain < steelMaterial.PhysicalTensionStrain * EXTENSION_START_FACTOR Then
        eps0 = eps0 + steelMaterial.PhysicalTensionStrain * EXTENSION_START_FACTOR - minRebarStrain
    End If
End Sub

' Проверяет, попал ли хотя бы один элемент начального поля
' на вспомогательное продолжение диаграммы. Иначе повторное решение
' началось бы на том же плато без новой касательной жёсткости.
Private Function CandidateUsesExtension(ByVal section As CSectionModel, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double) As Boolean
    Dim i As Long
    Dim strain As Double

    For i = 1 To section.ConcreteCount
        strain = PlaneStrain(eps0, kx, ky, section.ConcreteX(i), section.ConcreteY(i))
        If concreteMaterial.IsInExtensionRange(strain) Then
            CandidateUsesExtension = True
            Exit Function
        End If
    Next i

    For i = 1 To section.RebarCount
        strain = PlaneStrain(eps0, kx, ky, section.RebarX(i), section.RebarY(i))
        If steelMaterial.IsInExtensionRange(strain) Then
            CandidateUsesExtension = True
            Exit Function
        End If
    Next i
End Function

' Возвращает sample-деформацию на растянутой extension-ветви материала.
Private Function TryTensionExtensionSample(ByVal material As CMaterialDiagram, ByRef sampleStrain As Double) As Boolean
    If material Is Nothing Then Exit Function
    If Not material.HasTensionExtension Then Exit Function
    If material.PhysicalTensionStrain <= 0# Then Exit Function

    sampleStrain = material.PhysicalTensionStrain * EXTENSION_START_FACTOR
    TryTensionExtensionSample = (material.GetTangentModulus(sampleStrain) > 0#)
End Function

' Возвращает sample-деформацию на сжатой extension-ветви материала.
Private Function TryCompressionExtensionSample(ByVal material As CMaterialDiagram, ByRef sampleStrain As Double) As Boolean
    If material Is Nothing Then Exit Function
    If Not material.HasCompressionExtension Then Exit Function
    If material.PhysicalCompressionStrain >= 0# Then Exit Function

    sampleStrain = material.PhysicalCompressionStrain * EXTENSION_START_FACTOR
    TryCompressionExtensionSample = (material.GetTangentModulus(sampleStrain) > 0#)
End Function

' Показывает, вышло ли НДС неудачной попытки за физический предел деформации eps_ult.
Private Function FailedStateAlreadyInExtension(ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram) As Boolean
    If steelMaterial.HasTensionExtension And _
            failedService.MaxSteelStrain > steelMaterial.PhysicalTensionStrain Then
        FailedStateAlreadyInExtension = True
        Exit Function
    End If
    If steelMaterial.HasCompressionExtension And _
            failedService.MinSteelStrain < steelMaterial.PhysicalCompressionStrain Then
        FailedStateAlreadyInExtension = True
        Exit Function
    End If
    If concreteMaterial.HasCompressionExtension And _
            failedService.MinConcreteStrain < concreteMaterial.PhysicalCompressionStrain Then
        FailedStateAlreadyInExtension = True
    End If
End Function

' Масштаб до ближайшего растянутого physical limit.
Private Function ExtensionScaleForPositiveLimit(ByVal currentStrain As Double, _
        ByVal physicalLimit As Double) As Double
    If physicalLimit <= 0# Or currentStrain <= STATE_NEAR_ZERO Then Exit Function
    If currentStrain >= physicalLimit * EXTENSION_START_FACTOR Then Exit Function
    ExtensionScaleForPositiveLimit = physicalLimit * EXTENSION_START_FACTOR / currentStrain
End Function

' Масштаб до ближайшего сжатого physical limit.
Private Function ExtensionScaleForNegativeLimit(ByVal currentStrain As Double, _
        ByVal physicalLimit As Double) As Double
    If physicalLimit >= 0# Or currentStrain >= -STATE_NEAR_ZERO Then Exit Function
    If Abs(currentStrain) >= Abs(physicalLimit) * EXTENSION_START_FACTOR Then Exit Function
    ExtensionScaleForNegativeLimit = Abs(physicalLimit) * EXTENSION_START_FACTOR / Abs(currentStrain)
End Function

' Деформация точки в принятой НДМ-плоскости.
Private Function PlaneStrain(ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double, _
        ByVal x As Double, ByVal y As Double) As Double
    PlaneStrain = eps0 + kx * (y - mFrame.OriginY) + ky * (x - mFrame.OriginX)
End Function

' Проверяет изменение поля после уточнения работающих участков диаграмм.
' Если оно почти не изменилось, дальнейшие пересборки дадут те же участки;
' полученный старт можно передать обычному решателю.
Private Function ExtensionPlaneClose(ByVal oldEps0 As Double, ByVal oldKx As Double, ByVal oldKy As Double, _
        ByVal newEps0 As Double, ByVal newKx As Double, ByVal newKy As Double) As Boolean
    ExtensionPlaneClose = Abs(newEps0 - oldEps0) <= ACTIVE_SET_EPSILON_TOLERANCE And _
        mFrame.LengthScale * Abs(newKx - oldKx) <= ACTIVE_SET_EPSILON_TOLERANCE And _
        mFrame.LengthScale * Abs(newKy - oldKy) <= ACTIVE_SET_EPSILON_TOLERANCE
End Function

' Локальный минимум/максимум оставлены в классе, чтобы построение стартовых
' плоскостей не зависело от Excel WorksheetFunction и не создавало Variant.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Парная утилита к MaxDouble для ограничений и масштабов стартовой плоскости.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
