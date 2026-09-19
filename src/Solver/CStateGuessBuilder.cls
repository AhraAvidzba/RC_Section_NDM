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
' Построение стартовой плоскости для попыток StateSolution
' ==========================================================================
' CStateGuessBuilder строит только начальную плоскость деформаций
' epsilon0/kappaX/kappaY для очередной попытки решения равновесия. Он не
' выполняет Newton/Secant-итерации и не принимает решение о статусе расчета:
' после его работы CSectionSolver все равно заново решает полную систему НДС.
'
' У класса две практические роли. Первая - дать ненулевой старт штатному
' DirectState при чистом изгибе N = 0, когда нулевая плоскость попадает ровно
' в излом диаграммы бетона между сжатием и выключенным растяжением. Вторая -
' подготовить старт для повторной попытки StateSolution в technical extension,
' если первая попытка не смогла пройти горизонтальное плато диаграммы.
'
' Класс не выбирает материалы, не знает про расчетные профили, Capacity, Crack,
' Excel и СП. Он получает уже готовые CMaterialDiagram и CSectionModel, после
' чего пытается дать solver-у более осмысленный старт. Для физического
' DirectState это старт около рабочей стороны излома диаграммы; для extension
' это старт уже в области искусственного продолжения, где есть малая
' касательная жесткость и Newton/Secant снова получает управляемую матрицу.
'
' Это не меняет физический результат и не повышает несущую способность:
' CSectionSolver затем заново решает полное равновесие, а batch-слой отдельно
' помечает найденное состояние как FAIL, если оно вышло за физическую диаграмму.

Private Const STATE_NEAR_ZERO As Double = 0.000000000001 ' Допуск для почти нулевой стартовой плоскости.
Private Const EXTENSION_START_FACTOR As Double = 1.02 ' Небольшой заход за физический eps_ult для старта повторного solve в extension.
Private Const EXTENSION_GUESS_ATTEMPTS As Long = 5 ' Количество независимых стартовых оценок для повторного StateSolution.
Private Const EXTENSION_ACTIVE_SET_ITERATIONS As Long = 8 ' Сколько раз уточнять кусочно-линейный набор участков диаграмм.
Private Const ACTIVE_SET_EPSILON_TOLERANCE As Double = 0.0000000001 ' Допуск стабилизации epsilon0 в active-set уточнении.
Private Const ACTIVE_SET_KAPPA_TOLERANCE As Double = 0.000000000001 ' Допуск стабилизации кривизн в active-set уточнении.

' Строит стартовую плоскость для штатного прямого НДС.
' Метод особенно полезен для чистого изгиба N=0: старт из нулевой плоскости
' попадает ровно в излом диаграммы бетона, где сжатая производная ненулевая,
' а растянутая ветвь при Ignore уже выключена. Здесь сначала строится грубая
' линейно-упругая плоскость, затем несколько раз уточняется active-set
' кусочно-линейной диаграммы. Полученная плоскость не является результатом:
' она только помогает CSectionSolver начать итерации с правильной стороны
' излома диаграммы.
Public Function BuildForDirectState(ByVal section As CSectionModel, _
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

    BuildForDirectState = TryBuildFromTrialPlane(section, concreteMaterial, steelMaterial, _
        targetN, targetMx, targetMy, trialEps0, trialKx, trialKy, eps0, kx, ky, False)
End Function

' Строит стартовую плоскость для повторной попытки прямого StateSolution.
' Метод сначала использует форму деформаций, которую уже нашел несошедшийся
' solver: это важно для изгиба N + Mx + My, где отношение kappaX/kappaY несет
' информацию о направлении нейтральной линии. Если такая форма непригодна,
' используется более грубая линейная оценка по extension-ветвям материалов.
' Это не "следующая итерация" основного solver-а, а новый старт перед новым
' вызовом solve на той же нагрузке и той же расширенной StateSolution-диаграмме.
Public Function BuildForExtensionState(ByVal section As CSectionModel, _
        ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If section Is Nothing Then Exit Function
    If concreteMaterial Is Nothing Then Exit Function
    If steelMaterial Is Nothing Then Exit Function

    Dim attemptIndex As Long
    For attemptIndex = 1 To EXTENSION_GUESS_ATTEMPTS
        If BuildAttemptForExtensionState(attemptIndex, section, failedService, _
                concreteMaterial, steelMaterial, targetN, targetMx, targetMy, eps0, kx, ky) Then
            BuildForExtensionState = True
            Exit Function
        End If
    Next attemptIndex
End Function

' Возвращает число независимых вариантов стартовой плоскости.
' Batch-слой использует это, чтобы не останавливаться на первом варианте,
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
        Case 4: ExtensionAttemptName = "масштабированное несошедшееся состояние"
        Case 5: ExtensionAttemptName = "нулевое осевое приближение"
        Case Else: ExtensionAttemptName = "неизвестное приближение"
    End Select
End Function

' Строит конкретный вариант стартовой плоскости для повторного solve.
' Варианты идут от наиболее расчетно осмысленного к более простым fallback:
' если один старт плохо подошел для Newton/Secant, следующий может оказаться
' устойчивее без изменения диаграмм и без подмены расчетной физики.
Public Function BuildAttemptForExtensionState(ByVal attemptIndex As Long, _
        ByVal section As CSectionModel, ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    Select Case attemptIndex
        Case 1
            BuildAttemptForExtensionState = TryBuildGlobalExtensionEquilibrium(section, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, eps0, kx, ky)
        Case 2
            BuildAttemptForExtensionState = TryBuildFromFailedState(section, failedService, concreteMaterial, steelMaterial, _
                targetN, targetMx, targetMy, eps0, kx, ky)
        Case 3
            BuildAttemptForExtensionState = TryBuildUniformExtensionShift(failedService, concreteMaterial, steelMaterial, _
                targetN, eps0, kx, ky)
        Case 4
            BuildAttemptForExtensionState = TryBuildScaledFailedState(failedService, concreteMaterial, steelMaterial, eps0, kx, ky)
        Case 5
            BuildAttemptForExtensionState = TryBuildAxialZeroGuess(concreteMaterial, steelMaterial, targetN, eps0, kx, ky)
    End Select
End Function

' Использует последнюю плоскость несошедшегося solver-а как подсказку.
' Если первая попытка уже уперлась в физический предел, эта форма обычно
' правильно отражает изгибную картину; нужно только пересобрать линейную оценку
' по тем элементам, которые в этой форме попали бы в technical extension.
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
        Optional ByVal requireExtension As Boolean = True) As Boolean
    Dim currentEps0 As Double
    Dim currentKx As Double
    Dim currentKy As Double
    Dim nextEps0 As Double
    Dim nextKx As Double
    Dim nextKy As Double
    Dim iter As Long
    Dim solvedAtLeastOnce As Boolean

    currentEps0 = trialEps0
    currentKx = trialKx
    currentKy = trialKy

    ' Уточняем старт не через частные условия по форме сечения, а через
    ' active-set для кусочно-линейных диаграмм. После каждого линейного решения
    ' смотрим, на каких участках диаграммы фактически оказались элементы, и
    ' пересобираем систему. Если набор участков стабилизировался, полученная
    ' плоскость уже является хорошим стартом для обычного CSectionSolver.
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
            Exit For
        End If

        currentEps0 = nextEps0
        currentKx = nextKx
        currentKy = nextKy
    Next iter

    If solvedAtLeastOnce Then
        If requireExtension Then
            TryBuildFromTrialPlane = CandidateUsesExtension(section, concreteMaterial, steelMaterial, eps0, kx, ky)
        Else
            TryBuildFromTrialPlane = Abs(eps0) + Abs(kx) + Abs(ky) > STATE_NEAR_ZERO
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

    ' Арматуру также учитываем полностью: упругие стержни дают касательную
    ' жесткость, плато дает постоянное усилие, extension - малую добавочную
    ' жесткость за физическим пределом. Важно повторять контракт CSectionSolver:
    ' стержень добавляется как "сталь минус вытесненный бетон", потому что
    ' бетонная сетка уже содержит площадь в зоне расположения арматуры.
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

' Строит первую линейно-упругую оценку около нулевого состояния.
' Сжатый малый sample нужен, чтобы бетонная ветвь с Ignore не стартовала из
' полностью нулевой жесткости. Следующий active-set шаг уже сам выключит те
' бетонные элементы, которые по найденной пробной плоскости ушли в растяжение.
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

' Сдвигает уже найденную форму деформаций к ближайшей extension-области без
' изменения кривизн. Это простой запасной вариант: если первая попытка хорошо
' определила направление изгиба, но остановилась на плато, достаточно подвинуть
' epsilon0, чтобы solver начал вторую попытку за eps_ult.
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

    eps0 = failedService.Epsilon0 + shift
    kx = failedService.KappaX
    ky = failedService.KappaY
    TryBuildUniformExtensionShift = True
End Function

' Последний мягкий fallback: сохраняем форму несошедшегося состояния и только
' масштабируем ее до ближайшего physical limit. Это менее точная оценка, чем
' линейная система по активным extension-ветвям, но она часто лучше нуля.
Private Function TryBuildScaledFailedState(ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    TryBuildScaledFailedState = BuildExtensionTrialPlane(failedService, concreteMaterial, steelMaterial, eps0, kx, ky)
End Function

' Если первая попытка не дала вообще никакой формы, используем знак N только
' как аварийное начальное приближение. Полное равновесие по-прежнему будет
' искать CSectionSolver на следующем шаге.
Private Function TryBuildAxialZeroGuess(ByVal concreteMaterial As CMaterialDiagram, _
        ByVal steelMaterial As CMaterialDiagram, ByVal targetN As Double, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If targetN > 0# And steelMaterial.HasTensionExtension Then
        eps0 = steelMaterial.PhysicalTensionStrain * EXTENSION_START_FACTOR
        TryBuildAxialZeroGuess = True
    ElseIf targetN < 0# And concreteMaterial.HasCompressionExtension Then
        eps0 = concreteMaterial.PhysicalCompressionStrain * EXTENSION_START_FACTOR
        TryBuildAxialZeroGuess = True
    End If
End Function

' Создает пробную плоскость из несошедшегося solver-а: либо берет уже найденное
' состояние за физическим пределом, либо масштабирует его до небольшого выхода
' за ближайший предел. Метод не решает равновесие, а только подготавливает
' форму деформаций для выбора активных extension-элементов.
Private Function BuildExtensionTrialPlane(ByVal failedService As CSectionSolver, _
        ByVal concreteMaterial As CMaterialDiagram, ByVal steelMaterial As CMaterialDiagram, _
        ByRef eps0 As Double, ByRef kx As Double, ByRef ky As Double) As Boolean
    If failedService Is Nothing Then Exit Function

    Dim maxAbsState As Double
    maxAbsState = MaxDouble(Abs(failedService.Epsilon0), _
        MaxDouble(Abs(failedService.KappaX), Abs(failedService.KappaY)))
    If maxAbsState <= STATE_NEAR_ZERO Then Exit Function

    If FailedStateAlreadyInExtension(failedService, concreteMaterial, steelMaterial) Then
        eps0 = failedService.Epsilon0
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

    eps0 = failedService.Epsilon0 * scaleFactor
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
    intercept = stress - tangent * sampleStrain
    ea = tangent * area
    force0 = intercept * area

    If tangent > 0# Then
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

' Решает линейное равновесие, полученное из активных участков диаграмм.
' Если активных элементов нет или матрица вырождена, builder просто отдает
' управление следующему fallback-способу.
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
            targetN - n0, targetMx - mx0, targetMy - my0) Then
        Exit Function
    End If

    eps0 = system.X1
    kx = system.X2
    ky = system.X3
    SolveLinearizedSystem = (Abs(eps0) + Abs(kx) + Abs(ky) > STATE_NEAR_ZERO)
End Function

' Подталкивает глобальную растянутую оценку так, чтобы хотя бы крайние стержни
' не остались на физическом плато. Так solver сразу получает рабочий наклон
' для осевого растяжения, но не заменяет последующий solve.
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

' Проверяет, что стартовая плоскость действительно попала в technical extension
' хотя бы одним элементом. Иначе повторный solve начнется в той же физической
' области плато и не получит новой касательной жесткости.
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

' Возвращает True, если несошедшийся solver уже вышел за физический eps_ult.
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
    PlaneStrain = eps0 + kx * y + ky * x
End Function

' Проверяет, изменилась ли плоскость после очередного active-set шага.
' Если почти не изменилась, дальнейшие пересборки дадут тот же набор участков
' диаграмм, поэтому старт можно отдавать обычному solver-у.
Private Function ExtensionPlaneClose(ByVal oldEps0 As Double, ByVal oldKx As Double, ByVal oldKy As Double, _
        ByVal newEps0 As Double, ByVal newKx As Double, ByVal newKy As Double) As Boolean
    ExtensionPlaneClose = Abs(newEps0 - oldEps0) <= ACTIVE_SET_EPSILON_TOLERANCE And _
        Abs(newKx - oldKx) <= ACTIVE_SET_KAPPA_TOLERANCE And _
        Abs(newKy - oldKy) <= ACTIVE_SET_KAPPA_TOLERANCE
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
