VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подготовка данных и расчет ширины нормальной трещины
' ==========================================================================
' По готовому FormationResult и текущему CrackedState выбирает tension zone,
' арматуру, Abt/As/ds/ls, sigma_s/sigma_s,crc и коэффициенты. Численное
' трещинообразование принадлежит CCrackFormationCalculator; здесь нет Search,
' solver retries, Pre/Post solves или собственного рабочего formation-контекста.
' Чистые численные методы формулы a_crc отделены от подготовки зоны и статусов.

Private Const PHI1_LONG_TERM As Double = 1.4 ' Коэффициент длительного действия нагрузки для a_crc1 по СП 63.
Private Const PHI2_DEFAULT As Double = 0.5 ' Типовой default для обычной арматуры периодического профиля.
Private Const PHI3_TENSION As Double = 1.2 ' Коэффициент для центрального/сквозного растяжения.
Private Const PHI3_BENDING As Double = 1# ' Коэффициент для изгиба и внецентренного сжатия.
Private Const NEAR_ZERO_STRAIN As Double = 0.000000000001 ' Численный ноль для деформаций и кривизн.
Private Const NEAR_ZERO_FORCE As Double = 0.000001 ' Численный ноль для отличия центрального растяжения от нулевой нагрузки.
Private Const NEAR_ZERO_STRESS As Double = 0.000000001 ' Численный ноль для справочных напряжений, МПа.
Private Const COVER_MODE_NEAREST_CONTOUR As String = "NearestContour" ' a_s до ближайшего выхода бетонного контура по нормали.
Private Const COVER_MODE_GLOBAL_EXTREME As String = "GlobalExtreme" ' a_s до дальней растянутой опорной линии всего сечения.
Private Const PSI_MODE_USER As String = "User" ' psi_s задает пользователь через SLS.Crack.PsiS.
Private Const PSI_MODE_AUTO As String = "Auto" ' psi_s уточняется только если первая проверка с psi_s = 1 не прошла.
Private Const PSI_MODE_ALWAYS_CALC As String = "AlwaysCalc" ' psi_s всегда считается по sigma_s,crc, если такая точка доступна.
Private Const SIGMA_S_CRC_AVERAGE_ALL_SELECTED As String = "AllSelected" ' Для sigma_s,crc учитывать все выбранные стержни со знаком напряжения.
Private Const SIGMA_S_CRC_AVERAGE_TENSION_ONLY As String = "TensionOnly" ' Сжатые напряжения в sigma_s,crc принимать равными нулю.

Private mConverged As Boolean ' True, если расчет трещины завершился без ошибки входных данных или несходимости.
Private mStopReason As String ' Короткое пояснение результата: расчет выполнен, трещина не образована или solver не сошелся.
Private mDiagnosticLog As String ' Текстовая диагностика выбора зоны, арматуры, ls и psi_s.
Private mDiagnosticsEnabled As Boolean ' Отключает подробный журнал, когда txt-отчет выполнения не нужен.
Private mWidthInternalStatus As EResultInternalStatus ' Машинный статус формульного расчета ширины нормальной трещины.
Private mWidthResultCode As EResultCode ' Машинный код причины результата формульной ширины трещины.
Private mFormulaCalculated As Boolean ' Собственная формула уже выполнялась; готовая Formation сама по себе этого не означает.
Private mAllowableCrackWidth As Double ' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
Private mCrackWidth As Double ' Расчетная ширина раскрытия a_crc, мм.
Private mUtilization As Double ' Коэффициент использования по трещине: a_crc / a_crc,ult.
Private mCrackFormed As Boolean ' True, если по выбранной ветке расчета трещина считается образованной.
Private mPsiMode As String ' User, Auto или AlwaysCalc: как принимать psi_s в формуле раскрытия.
Private mSigmaSCrcAveragingMode As String ' Как усреднять sigma_s,crc по ранее выбранной группе стержней.
Private mPhi3Mode As String ' User или Auto: брать phi3 из Config или выбирать по характеру продольной силы.
Private mTensionZoneMode As String ' Effective или FullTension: способ задания бетонной зоны Abt.
Private mCoverMode As String ' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
Private mPhi1Input As Double ' Пользовательское значение phi1 для формулы раскрытия.
Private mPhi2Input As Double ' Пользовательское значение phi2 для формулы раскрытия.
Private mPhi3Input As Double ' Пользовательское значение phi3, если Phi3Mode = User.
Private mPsiSInput As Double ' Пользовательское значение psi_s, если PsiMode = User.
Private mSteelEs As Double ' Es для формулы раскрытия трещины, МПа.
Private mCrackedStateSpec As CMaterialModelSpec ' Спецификация модели раскрытой трещины текущего LC.
Private mTensionRebarCount As Long ' Сколько стержней попало в расчетную растянутую арматуру As.
Private mMaxSteelStrain As Double ' Максимальная растягивающая деформация среди выбранных стержней As.
Private mMaxSteelStress As Double ' Напряжение в стержне с максимальной растягивающей деформацией, МПа.
Private mSigmaS As Double ' Средневзвешенное напряжение выбранной растянутой арматуры при текущей нагрузке, МПа.
Private mSigmaSCrc As Double ' То же напряжение в состоянии lambda_crc * LC после образования трещины, МПа.
Private mAs As Double ' Площадь выбранной растянутой арматуры, мм2.
Private mAbt As Double ' Площадь бетона, попавшая в расчетную растянутую зону, мм2.
Private mDsEquivalent As Double ' Эквивалентный диаметр выбранных стержней для разных d, мм.
Private mLsRaw As Double ' Нормативная длина ls до ограничений 10ds/100/40ds/400, мм.
Private mLs As Double ' Итоговое расстояние между трещинами после ограничений, мм.
Private mPsiS As Double ' Коэффициент неравномерности деформаций арматуры между трещинами.
Private mPhi1 As Double ' Коэффициент длительности действия нагрузки.
Private mPhi2 As Double ' Пользовательский коэффициент, учитывающий поверхность арматуры.
Private mPhi3 As Double ' Коэффициент характера нагружения.
Private mTensionRebarIds As String ' Имена стержней, вошедших в расчет sigma_s и As, через запятую.
Private mSectionDepthH As Double ' Габарит сечения по нормали к нейтральной линии по общему бетонному контуру КЭ.
Private mCoverA As Double ' Расстояние a_s по выбранному режиму CoverDistanceMode, мм.
Private mTensionDepth As Double ' Фактическая глубина растянутой бетонной зоны до нейтральной линии, мм.
Private mEffectiveZoneDepth As Double ' Глубина зоны Abt, реально использованная после ограничений, мм.
Private mCentralTensionBranch As Boolean ' True, если сработала отдельная ветка центрального растяжения.
Private mFormation As CCrackFormationResult ' Единственный источник стадии образования трещины.


' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Значения соответствуют текущей принятой постановке: длительное раскрытие,
' эффективная зона СП и пользовательские коэффициенты с нормативными defaults.
Private Sub Class_Initialize()
    mAllowableCrackWidth = 0.3
    mPsiMode = PSI_MODE_ALWAYS_CALC
    mSigmaSCrcAveragingMode = SIGMA_S_CRC_AVERAGE_ALL_SELECTED
    mPhi3Mode = "Auto"
    mTensionZoneMode = "Effective"
    mCoverMode = COVER_MODE_GLOBAL_EXTREME
    mPhi1Input = PHI1_LONG_TERM
    mPhi2Input = PHI2_DEFAULT
    mPhi3Input = PHI3_BENDING
    mPsiSInput = 1#
    mSteelEs = 200000#
    mDiagnosticsEnabled = True
End Sub


' Сбрасывает данные подготовки формулы предыдущего сочетания.
Private Sub ClearResult()
    mConverged = False
    mStopReason = vbNullString
    mDiagnosticLog = vbNullString
    mWidthInternalStatus = rsUnset
    mWidthResultCode = rcNone
    mFormulaCalculated = False
    mCrackWidth = 0#
    mUtilization = 0#
    mCrackFormed = False
    mTensionRebarCount = 0
    mMaxSteelStrain = 0#
    mMaxSteelStress = 0#
    mSigmaS = 0#
    mSigmaSCrc = 0#
    mAs = 0#
    mAbt = 0#
    mDsEquivalent = 0#
    mLsRaw = 0#
    mLs = 0#
    mPsiS = mPsiSInput
    mPhi1 = mPhi1Input
    mPhi2 = mPhi2Input
    mPhi3 = 0#
    mTensionRebarIds = vbNullString
    Set mCrackedStateSpec = Nothing
    mSectionDepthH = 0#
    mCoverA = 0#
    mTensionDepth = 0#
    mEffectiveZoneDepth = 0#
    mCentralTensionBranch = False
    Set mFormation = Nothing
End Sub


' Принимает готовый результат formation без поиска НДС. Этот вход позволяет
' orchestrator-у хранить ширину как зависимую, еще не выполненную ветвь.
Public Sub InitializeForFormation(ByVal formation As CCrackFormationResult)
    ClearResult
    Set mFormation = formation
    If formation Is Nothing Then
        SetWidthResult rsInternalError, rcInternalError, "Не передан результат образования трещины."
        Exit Sub
    End If
    mConverged = (formation.ResultMeta.InternalStatus = rsSuccess Or _
        formation.ResultMeta.InternalStatus = rsSuccessWithWarning)
    mCrackFormed = formation.CrackFormed
    mCentralTensionBranch = formation.CentralTensionBranch
    If Not mConverged Then
        SetWidthResult rsBlockedByDependency, rcBlockedByDependency, _
            "Ширина нормальной трещины не рассчитана: нет корректного результата образования трещины."
    Else
        SetWidthResult rsNotApplicable, rcCrackNotFormed, formation.StopReason
    End If
End Sub

' Готовит зону, арматуру и коэффициенты по готовым состояниям. Сам поиск
' образования трещины и Pre/Post НДС здесь не выполняется; чистая формула
' получает только подготовленные числа и не знает о State/status.
Public Sub Calculate(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal formation As CCrackFormationResult, ByVal loadStateForClassification As CSectionLoadState)
    InitializeForFormation formation
    If formation Is Nothing Then Exit Sub
    If Not mConverged Or Not mCrackFormed Then Exit Sub
    If currentSolver Is Nothing Then
        mConverged = False
        SetWidthResult rsBlockedByDependency, rcBlockedByDependency, "Для ширины трещины не передано текущее НДС сочетания."
        Exit Sub
    End If
    If Not currentSolver.Converged Or Not currentSolver.WithinPhysicalRange Then
        mConverged = False
        SetWidthResult rsBlockedByDependency, rcBlockedByDependency, "Ширина трещины не рассчитана: текущее НДС не найдено или находится за физическими пределами материалов."
        Exit Sub
    End If
    ValidateInputs currentSolver, section, materialProvider
    If loadStateForClassification Is Nothing Then _
        Err.Raise vbObjectError + 3715, "CCrackWidthCalculator", "Состояние нагрузки для коэффициента phi3 не передано."
    If crackedStateSpec Is Nothing Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "MaterialModel.CrackedState не передан."
    If Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "MaterialModel.CrackedState не заполнен."
    Set mCrackedStateSpec = crackedStateSpec.Clone
    mSteelEs = materialProvider.SteelElasticModulusFromSpec(mCrackedStateSpec)
    NormalizeSettings
    mConverged = False
    mCrackFormed = False
    mStopReason = vbNullString
    mWidthInternalStatus = rsUnset
    mWidthResultCode = rcNone
    If formation.CentralTensionBranch Then
        CalculateCentralTension currentSolver, section, materialProvider, 0#, 0#, 0#, loadStateForClassification
    Else
        CalculateGeneralBending currentSolver, section, materialProvider, 0#, 0#, 0#, loadStateForClassification
    End If
End Sub

' Дает доступ к единственному готовому formation-узлу, не копируя его поля.
Public Property Get FormationResult() As CCrackFormationResult
    Set FormationResult = mFormation
End Property


' ------------------------------
' Чтение настроек Config
' ------------------------------
' Загружает настройки подготовки формулы ширины. Поиск НДС и параметры
' трещинообразования настраиваются у соответствующих владельцев; допустимая
' ширина здесь переводится в миллиметры внутренней системы единиц.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub

    mAllowableCrackWidth = settings.GetDouble("SLS.Crack.Allowable", mAllowableCrackWidth)
    mPsiMode = settings.GetString("SLS.Crack.PsiMode", mPsiMode)
    mSigmaSCrcAveragingMode = settings.GetString("SLS.Crack.SigmaSCrcAveragingMode", mSigmaSCrcAveragingMode)
    mTensionZoneMode = settings.GetString("SLS.Crack.TensionZoneMode", mTensionZoneMode)
    If settings.HasKey("SLS.Crack.CoverDistanceMode") Then
        mCoverMode = settings.GetString("SLS.Crack.CoverDistanceMode", mCoverMode)
    Else
        mCoverMode = settings.GetString("SLS.Crack.CoverMode", mCoverMode)
    End If
    mPhi1Input = settings.GetDouble("SLS.Crack.Phi1", mPhi1Input)
    mPhi2Input = settings.GetDouble("SLS.Crack.Phi2", mPhi2Input)
    mPhi3Mode = settings.GetString("SLS.Crack.Phi3Mode", mPhi3Mode)
    mPhi3Input = settings.GetDouble("SLS.Crack.Phi3", mPhi3Input)
    mPsiSInput = settings.GetDouble("SLS.Crack.PsiS", mPsiSInput)


    If Not units Is Nothing Then
        mAllowableCrackWidth = units.InputLengthToInternal(mAllowableCrackWidth)
    End If
End Sub



' Выполняет отдельную ветку СП для истинного центрального растяжения.
' Здесь нет нейтральной линии, поэтому Abt принимается равной всей площади
' бетона, а As - всей продольной арматуре, которая растянута в текущем НДС.
Private Sub CalculateCentralTension(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mCentralTensionBranch = True
    mPhi3 = ResolvePhi3(loadStateForClassification)

    If Not IsTensionLoad(loadStateForClassification) Then
        FinishNoCrack "Центральная ветка: продольная сила не является растягивающей."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectAllTensionRebars currentSolver, section, selected
    If mTensionRebarCount = 0 Then
        FinishNoCrack "В центральном растяжении не найдено растянутой арматуры."
        Exit Sub
    End If

    mAbt = TotalConcreteArea(section)
    mAs = SelectedRebarArea(section, selected)
    CompleteCrackFormula currentSolver, section, materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected
    If mStopReason <> vbNullString And Not mCrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(section, materialProvider, currentSolver, selected, _
            targetN, targetMx, targetMy, "Расчет выполнен для центрального растяжения; Abt принята по всей площади бетона.") Then Exit Sub
End Sub



' Выполняет общий расчет для N + Mx + My при существующей нейтральной линии.
' Все геометрические величины строятся в координате s вдоль нормали к
' нейтральной линии, направленной в сторону растяжения.
Private Sub CalculateGeneralBending(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal targetN As Double, _
        ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState)
    mPhi3 = ResolvePhi3(loadStateForClassification)

    Dim nx As Double
    Dim ny As Double
    Dim normK As Double
    normK = Sqr(currentSolver.KappaX * currentSolver.KappaX + currentSolver.KappaY * currentSolver.KappaY)
    nx = currentSolver.KappaY / normK
    ny = currentSolver.KappaX / normK

    Dim st As Double
    Dim sc As Double
    ProjectionBoundsByConcreteContour section, nx, ny, st, sc
    mSectionDepthH = st - sc
    If mSectionDepthH <= NEAR_ZERO_STRAIN Then
        SetWidthInvalidInput "InputError: не удалось определить габарит сечения по нормали к нейтральной линии."
        Exit Sub
    End If

    Dim neutralS As Double
    neutralS = -currentSolver.Epsilon0 / normK
    mTensionDepth = st - MaxDouble(neutralS, sc)
    If mTensionDepth <= NEAR_ZERO_STRAIN Then
        FinishNoCrack "В текущем состоянии нет растянутой бетонной зоны."
        Exit Sub
    End If

    Dim sBarMax As Double
    Dim coverRebarIndex As Long
    If Not FindMostTensionRebarProjection(currentSolver, section, nx, ny, sBarMax, coverRebarIndex) Then
        FinishNoCrack "В расчетной растянутой зоне не найдено растянутой арматуры."
        Exit Sub
    End If
    mCoverA = CoverDistanceForRebar(section, coverRebarIndex, nx, ny, st, sBarMax)
    If mCoverA < 0# Then mCoverA = 0#

    Dim zoneLowerS As Double
    zoneLowerS = TensionZoneLowerBound(st, sc, neutralS)
    mEffectiveZoneDepth = st - zoneLowerS
    If mEffectiveZoneDepth <= NEAR_ZERO_STRAIN Then
        FinishNoCrack "Эффективная растянутая зона имеет нулевую глубину."
        Exit Sub
    End If

    mAbt = ConcreteAreaInZone(currentSolver, section, nx, ny, zoneLowerS, _
        StrComp(mTensionZoneMode, "FullTension", vbTextCompare) = 0)
    If mAbt <= 0# Then
        FinishNoCrack "В расчетную растянутую зону не попали бетонные элементы."
        Exit Sub
    End If

    Dim selected() As Boolean
    ReDim selected(1 To section.RebarCount)
    SelectRebarsInZone currentSolver, section, nx, ny, zoneLowerS, selected
    If mTensionRebarCount = 0 Then
        FinishNoCrack "В расчетную растянутую зону не попали растянутые стержни."
        Exit Sub
    End If

    CompleteCrackFormula currentSolver, section, materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected
    If mStopReason <> vbNullString And Not mCrackFormed Then Exit Sub
    If Not FinalizeWithPsiMode(section, materialProvider, currentSolver, selected, _
            targetN, targetMx, targetMy, "Расчет выполнен для произвольного сечения по общей НДМ-ветке.") Then Exit Sub
End Sub



' Завершает подготовку нормативной формулы раскрытия трещин после выбора
' растянутой зоны и арматуры. Метод собирает As, Abt, d_s,eq, sigma_s и
' расстояние между трещинами l_s; сам коэффициент psi_s уточняется ниже.
Private Sub CompleteCrackFormula(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean)
    mPhi1 = mPhi1Input
    mPhi2 = mPhi2Input
    If mPhi3 <= 0# Then mPhi3 = mPhi3Input
    mPsiS = mPsiSInput

    mAs = SelectedRebarArea(section, selected)
    mTensionRebarIds = SelectedRebarIds(section, selected)
    mDsEquivalent = EquivalentDiameter(section, selected)
    mSigmaS = AverageSelectedSteelStress(solver.Epsilon0, solver.KappaX, solver.KappaY, section, steelMaterial, selected)
    If mAs <= 0# Or mAbt <= 0# Or mDsEquivalent <= 0# Or mSigmaS <= 0# Then
        FinishNoCrack "Недостаточно растянутой арматуры или бетона для расчета раскрытия трещины."
        Exit Sub
    End If

    mLsRaw = 0.5 * (mAbt / mAs) * mDsEquivalent
    mLs = LimitedCrackSpacing(mLsRaw, mDsEquivalent)
End Sub



' Завершает расчет раскрытия с учетом режима psi_s.
' User заканчивает расчет после первой формулы с пользовательским psi_s.
' Auto сначала делает проверку с psi_s = 1 и уточняет коэффициент только при
' непрохождении. AlwaysCalc сразу пытается применить уточненный psi_s. Точка
' lambda_crc и состояния PreCrackState/PostCrackState уже найдены обязательной
' предварительной проверкой образования трещины, поэтому повторный поиск точки
' образования здесь не выполняется.
Private Function FinalizeWithPsiMode(ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal currentSolver As CSectionStateResult, ByRef selected() As Boolean, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal finalReason As String) As Boolean
    If StrComp(mPsiMode, PSI_MODE_USER, vbTextCompare) = 0 Then
        mPsiS = mPsiSInput
        CaptureSigmaSCrcIfAvailable section, materialProvider, selected
        FinalizeWidth finalReason & " Psi_s принят из настройки SLS.Crack.PsiS."
        FinalizeWithPsiMode = mConverged
        Exit Function
    End If

    If StrComp(mPsiMode, PSI_MODE_ALWAYS_CALC, vbTextCompare) = 0 Then
        If ResolveAutoPsiIfPossible(section, materialProvider, selected, _
                "AlwaysCalc psi_s") Then
            FinalizeWidth finalReason & AutomaticPsiComment()
        ElseIf mFormation.FallbackPsi1 Then
            mPsiS = 1#
            FinalizeWidth finalReason & " Psi_s оставлен равным 1, потому что выбранный путь не дает отдельной точки sigma_s,crc."
        Else
            Exit Function
        End If
        FinalizeWithPsiMode = mConverged
        Exit Function
    End If

    mPsiS = 1#
    CaptureSigmaSCrcIfAvailable section, materialProvider, selected
    FinalizeWidth finalReason & " Auto psi_s: первая проверка выполнена с psi_s=1."
    If Not mConverged Or Not mCrackFormed Then Exit Function

    If mAllowableCrackWidth > 0# And mCrackWidth <= mAllowableCrackWidth Then
        AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 проходит; sigma_s,crc для уточнения psi_s не потребовался."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    If mFormation.FallbackPsi1 Then
        AppendDiagnostic "Auto psi_s: выбранный путь образования трещины не дает отдельной точки sigma_s,crc; psi_s оставлен равным 1."
        FinalizeWithPsiMode = True
        Exit Function
    End If

    AppendDiagnostic "Auto psi_s: первая проверка с psi_s=1 не проходит; используется ранее найденный PostCrackState."
    mConverged = False

    If Not ResolveAutoPsiIfPossible(section, materialProvider, selected, _
            "Auto psi_s") Then Exit Function
    FinalizeWidth finalReason & AutomaticPsiComment()
    FinalizeWithPsiMode = mConverged
End Function



' Объясняет фактически примененный коэффициент после подготовки PostCrackState.
' Неположительное среднее напряжение не является несходимостью и не уменьшает
' раскрытие; текст результата должен явно объяснять принятие psi_s = 1.
Private Function AutomaticPsiComment() As String
    If mSigmaSCrc <= NEAR_ZERO_STRESS Then
        AutomaticPsiComment = " Среднее напряжение sigma_s,crc по выбранным стержням неположительное или практически нулевое; " & _
            "psi_s принят равным 1, уменьшение раскрытия не применяется."
    Else
        AutomaticPsiComment = " Psi_s рассчитан по состоянию сразу после образования трещины."
    End If
End Function

' Готовит и применяет автоматическое значение psi_s по sigma_s,crc.
' Метод не запускает solve и не достраивает PostCrackState: на этом этапе
' используются только данные, уже подготовленные проверкой образования трещины.
' Если sigma_s,crc получить нельзя, psi_s остается равным 1 с диагностикой.
Private Function ResolveAutoPsiIfPossible(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByRef selected() As Boolean, _
        ByVal contextText As String) As Boolean
    If mFormation.FallbackPsi1 Then
        AppendDiagnostic contextText & ": выбранный путь не дает отдельной точки sigma_s,crc."
        Exit Function
    End If

    If mFormation.LambdaCrc < 0# Then
        SetWidthInvalidInput "InputError: lambda_crc предварительной проверки образования трещины не определена."
        Exit Function
    End If

    If mSigmaSCrc <= NEAR_ZERO_STRESS Then _
        CaptureSigmaSCrcIfAvailable section, materialProvider, selected

    If mSigmaSCrc <= NEAR_ZERO_STRESS Then
        AppendDiagnostic contextText & ": среднее sigma_s,crc неположительное или практически нулевое; psi_s оставлен равным 1."
        mPsiS = 1#
        ResolveAutoPsiIfPossible = True
        Exit Function
    End If

    mPsiS = AutoPsi(mSigmaS, mSigmaSCrc)
    ResolveAutoPsiIfPossible = True
End Function



' Заполняет справочное sigma_s,crc без дополнительных solve-вызовов.
' Если состояние PostCrackState уже найдено предварительной проверкой
' образования трещины, берем плоскость деформаций из сохраненного State.
' Живой solver не нужен. Для fallback-путей отдельной точки sigma_s,crc нет.
Private Sub CaptureSigmaSCrcIfAvailable(ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByRef selected() As Boolean)
    If mSigmaSCrc > NEAR_ZERO_STRESS Then Exit Sub
    If mFormation.FallbackPsi1 Then Exit Sub
    Dim postState As CSectionStateResult
    Set postState = mFormation.PostCrackState
    If postState Is Nothing Then Exit Sub
    If Not postState.Converged Then Exit Sub
    If Not postState.WithinPhysicalRange Then Exit Sub

    mSigmaSCrc = AverageSelectedSteelStress(postState.Epsilon0, postState.KappaX, postState.KappaY, section, _
        materialProvider.SteelMaterialFromSpec(mCrackedStateSpec), selected, False, IncludeCompressedStressForSigmaSCrc())
End Sub



' Завершает расчет по формуле СП 63 п. 8.2.15 для продолжительного раскрытия.
Private Sub FinalizeWidth(ByVal reason As String)
    If mSigmaS <= 0# Or mLs <= 0# Then
        FinishNoCrack "Напряжение растянутой арматуры или ls равны нулю."
        Exit Sub
    End If

    mFormulaCalculated = True
    mCrackWidth = CrackWidthFromData(mPhi1, mPhi2, mPhi3, mPsiS, mSigmaS, mSteelEs, mLs)
    mUtilization = UtilizationFromData(mCrackWidth, mAllowableCrackWidth)
    mCrackFormed = True
    mConverged = True
    If mAllowableCrackWidth > 0# And mCrackWidth > mAllowableCrackWidth Then
        SetWidthResult rsCheckFailed, rcCrackWidthExceeded, _
            "Раскрытие нормальной трещины " & FormatNumberInvariant(mCrackWidth) & _
            " мм превышает допускаемое " & FormatNumberInvariant(mAllowableCrackWidth) & " мм. " & reason
    Else
        SetWidthResult rsSuccess, rcCheckPassed, reason
    End If
    AppendDiagnostic "a_crc=" & FormatNumberInvariant(mCrackWidth) & _
        "; phi1=" & FormatNumberInvariant(mPhi1) & _
        "; phi2=" & FormatNumberInvariant(mPhi2) & _
        "; phi3=" & FormatNumberInvariant(mPhi3) & _
        "; psi_s=" & FormatNumberInvariant(mPsiS) & _
        "; sigma_s=" & FormatNumberInvariant(mSigmaS) & _
        "; ls=" & FormatNumberInvariant(mLs)
End Sub

' Выполняет только численную формулу a_crc по готовым коэффициентам,
' напряжению, модулю и расстоянию между трещинами. Не получает State,
' не выбирает арматуру и не меняет статус либо рабочий результат calculator-а.
Public Function CrackWidthFromData(ByVal phi1 As Double, ByVal phi2 As Double, _
        ByVal phi3 As Double, ByVal psiS As Double, ByVal sigmaS As Double, _
        ByVal steelEs As Double, ByVal crackSpacing As Double) As Double
    If steelEs <= 0# Then Exit Function
    CrackWidthFromData = phi1 * phi2 * phi3 * psiS * (sigmaS / steelEs) * crackSpacing
End Function

' Возвращает численное отношение готового раскрытия к допускаемому.
' При нулевом пределе отношение не определено; решение о применимости
' проверки принимается до вызова этой чистой операции.
Public Function UtilizationFromData(ByVal crackWidthValue As Double, ByVal allowableCrackWidth As Double) As Double
    If allowableCrackWidth <= 0# Then Exit Function
    UtilizationFromData = crackWidthValue / allowableCrackWidth
End Function

' Публикует машинный итог только формульного этапа. Получение Formation/State
' не объявляет формулу выполненной; отсутствующий итог является внутренней
' ошибкой, а не повторной численной несходимостью или успешным fallback.
Public Property Get ResultMeta() As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If mWidthInternalStatus = rsUnset Then
        meta.SetResult rsInternalError, rcInternalError, rkCrackWidth, "Результат расчета ширины трещины не заполнен.", _
            vbNullString, True, False
    Else
        meta.SetResult mWidthInternalStatus, mWidthResultCode, rkCrackWidth, mStopReason, _
            vbNullString, True, mFormulaCalculated
    End If
    Set ResultMeta = meta
End Property



' Записывает машинный результат формульной части ширины раскрытия.
Private Sub SetWidthResult(ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal reason As String)
    mWidthInternalStatus = internalStatus
    mWidthResultCode = resultCode
    mStopReason = UserStopReason(reason)
End Sub



' Готовит короткий пользовательский комментарий из внутренней причины.
' Машинная классификация уже сохранена в InternalStatus/ResultCode, поэтому
' технические префиксы NumericalFailure/InputError на лист выводить не нужно.
Private Function UserStopReason(ByVal reason As String) As String
    UserStopReason = Trim$(reason)
    If Len(UserStopReason) = 0 Then Exit Function

    UserStopReason = Replace$(UserStopReason, "NumericalFailure:", vbNullString, 1, 1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "InputError:", vbNullString, 1, 1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "lambda_crc", "lambda трещинообразования", 1, -1, vbTextCompare)
    UserStopReason = Replace$(UserStopReason, "eps_bt,ult", "предельная растягивающая деформация бетона", 1, -1, vbTextCompare)
    UserStopReason = Trim$(UserStopReason)
    Do While InStr(1, UserStopReason, "..", vbBinaryCompare) > 0
        UserStopReason = Replace$(UserStopReason, "..", ".")
    Loop
End Function



' Помечает ошибку входных данных, найденную уже при подготовке данных формулы.
Private Sub SetWidthInvalidInput(ByVal reason As String)
    mConverged = False
    mCrackFormed = False
    SetWidthResult rsInvalidInput, rcInvalidInput, reason
End Sub



' Помечает расчет как корректно завершенный без раскрытой трещины.
' Это не численная ошибка: заданное состояние либо не растягивает арматуру,
' либо еще не достигло момента образования трещины.
Private Sub FinishNoCrack(ByVal reason As String)
    mConverged = True
    mCrackFormed = False
    mCrackWidth = 0#
    mUtilization = 0#
    mStopReason = reason
    If mWidthInternalStatus = rsUnset Then _
        SetWidthResult rsNotApplicable, rcCrackNotFormed, reason
    AppendDiagnostic reason
End Sub



' Выбирает все растянутые стержни для центрального растяжения.
Private Sub SelectAllTensionRebars(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            selected(i) = True
            mTensionRebarCount = mTensionRebarCount + 1
        End If
    Next i
End Sub



' Находит крайнюю растянутую арматуру по нормали к нейтральной линии.
Private Function FindMostTensionRebarProjection(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef sBarMax As Double, ByRef rebarIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            Dim s As Double
            s = Projection(section.RebarX(i), section.RebarY(i), nx, ny)
            If Not FindMostTensionRebarProjection Or s > sBarMax Then sBarMax = s
            If s >= sBarMax - NEAR_ZERO_STRAIN Then rebarIndex = i
            FindMostTensionRebarProjection = True
        End If
    Next i
End Function



' Выбирает смысл расстояния a_s по настройке CoverMode. NearestContour
' использует локальный выход луча от стержня к бетонному контуру;
' GlobalExtreme считает расстояние до опорной линии всего сечения.
' Если локальный контур не определен, используется расстояние GlobalExtreme.
Private Function CoverDistanceForRebar(ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double, ByVal st As Double, ByVal sBarMax As Double) As Double
    If StrComp(mCoverMode, COVER_MODE_GLOBAL_EXTREME, vbTextCompare) = 0 Then
        CoverDistanceForRebar = st - sBarMax
        Exit Function
    End If

    CoverDistanceForRebar = LocalConcreteCoverForRebar(section, rebarIndex, nx, ny)
    If CoverDistanceForRebar < 0# Then CoverDistanceForRebar = st - sBarMax
End Function



' Для a_s нужна не глобальная крайняя проекция всего сечения, а расстояние от
' оси выбранного растянутого стержня до ближайшего бетонного контура по нормали
' к нейтральной линии. Это особенно важно для Г-сечений и импортированных
' повернутых сеток, где глобальная опорная линия может лежать на другом выступе.
Private Function LocalConcreteCoverForRebar(ByVal section As CSectionModel, ByVal rebarIndex As Long, _
        ByVal nx As Double, ByVal ny As Double) As Double
    If rebarIndex <= 0 Then
        LocalConcreteCoverForRebar = -1#
        Exit Function
    End If

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    LocalConcreteCoverForRebar = props.ConcreteCoverFromPointAlongDirection(section, _
        section.RebarX(rebarIndex), section.RebarY(rebarIndex), nx, ny)
End Function



' Определяет нижнюю границу расчетной зоны в координате s.
' FullTension берет всю фактически растянутую часть до нейтральной линии.
' Effective задает нормативную глубину h_bt = min(max(x_t, 2a), 0.5h).
' Если фактическая растянутая зона x_t меньше 2a, полоса может геометрически
' зайти за нейтральную линию. Для Abt это намеренно учитывается: бетонные КЭ
' выбираются по попаданию центра в полосу h_bt, а не по знаку текущей деформации.
' Арматура As остается только растянутой, потому что sigma_s в формуле раскрытия
' относится к работающей растянутой арматуре между трещинами.
Private Function TensionZoneLowerBound(ByVal st As Double, ByVal sc As Double, ByVal neutralS As Double) As Double
    Dim lower As Double
    If StrComp(mTensionZoneMode, "FullTension", vbTextCompare) = 0 Then
        lower = MaxDouble(neutralS, sc)
    Else
        Dim effectiveDepth As Double
        effectiveDepth = MinDouble(MaxDouble(mTensionDepth, 2# * mCoverA), 0.5 * mSectionDepthH)
        If effectiveDepth <= NEAR_ZERO_STRAIN Then effectiveDepth = MinDouble(mTensionDepth, 0.5 * mSectionDepthH)
        lower = st - effectiveDepth
        lower = MaxDouble(lower, sc)
    End If
    TensionZoneLowerBound = lower
End Function



' Суммирует площадь бетонных КЭ, чьи центры попали в расчетную растянутую
' зону. Частичные площади намеренно не считаются: это принятое дискретное
' приближение текущей НДМ-модели.
Private Function ConcreteAreaInZone(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double, _
        ByVal requireTensionStrain As Boolean) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        If Not requireTensionStrain Or ConcreteStrain(solver, section, i) > 0# Then
            If Projection(section.ConcreteX(i), section.ConcreteY(i), nx, ny) >= lowerS - NEAR_ZERO_STRAIN Then
                ConcreteAreaInZone = ConcreteAreaInZone + section.ConcreteArea(i)
            End If
        End If
    Next i
End Function



' Формирует набор стержней As по той же полосе, которая принята для Abt,
' но с обязательным физическим фильтром по растяжению стержня. Поэтому при
' Effective и срабатывании ограничения 2a бетонная полоса может зайти за
' нейтральную линию, однако сжатые стержни в As, sigma_s и d_s,eq не попадают.
Private Sub SelectRebarsInZone(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByVal lowerS As Double, ByRef selected() As Boolean)
    Dim i As Long
    For i = 1 To section.RebarCount
        If RebarStrain(solver, section, i) > 0# Then
            If Projection(section.RebarX(i), section.RebarY(i), nx, ny) >= lowerS - NEAR_ZERO_STRAIN Then
                selected(i) = True
                mTensionRebarCount = mTensionRebarCount + 1
            End If
        End If
    Next i
End Sub



' Внешняя растянутая поверхность берется через общий геометрический калькулятор.
' Для сгенерированной сетки используются реальные размеры бетонных КЭ, а для
' импортированных Region без габаритов сохраняется общий fallback: сначала
' эквивалентный прямоугольник по A/I, затем квадрат той же площади. Та же
' логика используется устойчивостью.
Private Sub ProjectionBoundsByConcreteContour(ByVal section As CSectionModel, _
        ByVal nx As Double, ByVal ny As Double, ByRef st As Double, ByRef sc As Double)
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, nx, ny, False, sc, st
End Sub



' Считает средневзвешенное напряжение выбранной арматуры. В формулу раскрытия
' идет одно глобальное sigma_s, а не локальные значения по отдельным зонам.
Private Function AverageSelectedSteelStress(ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal section As CSectionModel, _
        ByVal steelMaterial As Object, ByRef selected() As Boolean, _
        Optional ByVal trackCurrentMaxima As Boolean = True, _
        Optional ByVal includeCompressedStress As Boolean = False) As Double
    Dim weightedStress As Double
    Dim areaSum As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim strain As Double
            Dim stress As Double
            strain = epsilon0 + kappaX * section.RebarY(i) + kappaY * section.RebarX(i)
            stress = steelMaterial.GetStress(strain)
            If (Not includeCompressedStress) And stress < 0# Then stress = 0#
            weightedStress = weightedStress + section.RebarArea(i) * stress
            areaSum = areaSum + section.RebarArea(i)
            If trackCurrentMaxima Then TrackMaxSteel section, i, strain, stress
        End If
    Next i
    If areaSum > 0# Then AverageSelectedSteelStress = weightedStress / areaSum
End Function



' Возвращает режим усреднения именно для sigma_s,crc. Текущее sigma_s основной
' формулы раскрытия по-прежнему считается только по растягивающей части
' выбранной группы; этот переключатель нужен для состояния PostCrackState,
' где часть той же группы может перейти в сжатие сразу после образования трещины.
Private Function IncludeCompressedStressForSigmaSCrc() As Boolean
    IncludeCompressedStressForSigmaSCrc = (StrComp(mSigmaSCrcAveragingMode, _
        SIGMA_S_CRC_AVERAGE_ALL_SELECTED, vbTextCompare) = 0)
End Function



' Считает площадь выбранной арматуры.
Private Function SelectedRebarArea(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then SelectedRebarArea = SelectedRebarArea + section.RebarArea(i)
    Next i
End Function



' Перечисляет имена выбранных стержней для контрольной таблицы напряжений.
Private Function SelectedRebarIds(ByVal section As CSectionModel, ByRef selected() As Boolean) As String
    Dim result As String
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            If Len(result) > 0 Then result = result & ", "
            result = result & section.RebarID(i)
        End If
    Next i
    SelectedRebarIds = result
End Function



' Для разных диаметров используется эквивалентный диаметр Eurocode:
' ds,eq = sum(n_i*d_i^2) / sum(n_i*d_i). В дискретной модели каждый стержень
' считается отдельным i, поэтому n_i фактически равен количеству совпадающих d.
Private Function EquivalentDiameter(ByVal section As CSectionModel, ByRef selected() As Boolean) As Double
    Dim numerator As Double
    Dim denominator As Double
    Dim i As Long
    For i = 1 To section.RebarCount
        If selected(i) Then
            Dim d As Double
            d = section.RebarDiameter(i)
            If d <= 0# Then d = Sqr(4# * section.RebarArea(i) / GEOM_PI)
            numerator = numerator + d * d
            denominator = denominator + d
        End If
    Next i
    If denominator > 0# Then EquivalentDiameter = numerator / denominator
End Function



' Применяет ограничения СП 63 к расстоянию между трещинами.
Private Function LimitedCrackSpacing(ByVal rawLs As Double, ByVal ds As Double) As Double
    Dim lower As Double
    Dim upper As Double
    lower = MaxDouble(10# * ds, 100#)
    upper = MinDouble(40# * ds, 400#)
    If lower > upper Then
        AppendDiagnostic "Предупреждение: ограничения ls конфликтуют; использована верхняя граница " & FormatNumberInvariant(upper) & " мм."
        LimitedCrackSpacing = upper
    Else
        LimitedCrackSpacing = MinDouble(MaxDouble(rawLs, lower), upper)
    End If
End Function



' Формула psi_s = 1 - 0.8*sigma_s,crc/sigma_s. Ограничение 0...1 является
' численной защитой от отрицательного psi_s при почти равных напряжениях или
' случайных малых несоответствиях solve-ов.
Private Function AutoPsi(ByVal sigmaS As Double, ByVal sigmaSCrc As Double) As Double
    If sigmaS <= 0# Then
        AutoPsi = 1#
    Else
        AutoPsi = 1# - 0.8 * sigmaSCrc / sigmaS
        If AutoPsi < 0# Then AutoPsi = 0#
        If AutoPsi > 1# Then AutoPsi = 1#
    End If
End Function



' Суммарная площадь бетона модели.
Private Function TotalConcreteArea(ByVal section As CSectionModel) As Double
    Dim i As Long
    For i = 1 To section.ConcreteCount
        TotalConcreteArea = TotalConcreteArea + section.ConcreteArea(i)
    Next i
End Function



' Деформация бетонного КЭ.
Private Function ConcreteStrain(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteStrain = solver.Epsilon0 + solver.KappaX * section.ConcreteY(index) + solver.KappaY * section.ConcreteX(index)
End Function



' Деформация стержня арматуры.
Private Function RebarStrain(ByVal solver As CSectionStateResult, ByVal section As CSectionModel, ByVal index As Long) As Double
    RebarStrain = solver.Epsilon0 + solver.KappaX * section.RebarY(index) + solver.KappaY * section.RebarX(index)
End Function



' Проекция точки на нормаль к нейтральной линии.
Private Function Projection(ByVal xCoord As Double, ByVal yCoord As Double, ByVal nx As Double, ByVal ny As Double) As Double
    Projection = nx * xCoord + ny * yCoord
End Function



' Обновляет максимум деформации и напряжения арматуры для диагностики.
Private Sub TrackMaxSteel(ByVal section As CSectionModel, ByVal index As Long, ByVal strain As Double, ByVal stress As Double)
    AppendDiagnostic "bar=" & section.RebarID(index) & _
        "; strain=" & FormatNumberInvariant(strain) & _
        "; stress=" & FormatNumberInvariant(stress)
    If strain > mMaxSteelStrain Then
        mMaxSteelStrain = strain
        mMaxSteelStress = stress
    End If
End Sub



' Возвращает phi3 для текущего LC. В Auto используется принятое правило СП:
' при растягивающей продольной силе принимается 1.2, иначе 1.0. В User
' программа не выводит характер нагрузки из НДС, а берет введенное значение.
Private Function ResolvePhi3(ByVal loadStateForClassification As CSectionLoadState) As Double
    If StrComp(mPhi3Mode, "User", vbTextCompare) = 0 Then
        ResolvePhi3 = mPhi3Input
    ElseIf IsTensionLoad(loadStateForClassification) Then
        ResolvePhi3 = PHI3_TENSION
    Else
        ResolvePhi3 = PHI3_BENDING
    End If
End Function



' Определяет растягивающий характер продольной силы только через единый
' объект нагрузки. Так central tension и phi3 используют ту же интерпретацию
' Load.ReferenceOffset и фильтра малых моментов, что batch/stability/capacity.
Private Function IsTensionLoad(ByVal loadStateForClassification As CSectionLoadState) As Boolean
    IsTensionLoad = loadStateForClassification.IsTension(NEAR_ZERO_FORCE)
End Function



' Нормализует текстовые режимы и проверяет численные параметры до запуска
' внутренних solve, чтобы ошибки Config были понятны пользователю.
Private Sub NormalizeSettings()
    Select Case LCase$(Trim$(mPsiMode))
        Case "user", "manual", "fixed1", "unity"
            mPsiMode = PSI_MODE_USER
        Case "auto", "refined"
            mPsiMode = PSI_MODE_AUTO
        Case "alwayscalc", "alwayscalculate", "always", "calc"
            mPsiMode = PSI_MODE_ALWAYS_CALC
        Case Else
            Err.Raise vbObjectError + 3707, "CCrackWidthCalculator", "SLS.Crack.PsiMode должен быть User, Auto или AlwaysCalc."
    End Select

    Select Case LCase$(Trim$(mSigmaSCrcAveragingMode))
        Case vbNullString, "allselected", "all", "includecompression", "withcompression", "signed"
            mSigmaSCrcAveragingMode = SIGMA_S_CRC_AVERAGE_ALL_SELECTED
        Case "tensiononly", "tension", "positiveonly", "old"
            mSigmaSCrcAveragingMode = SIGMA_S_CRC_AVERAGE_TENSION_ONLY
        Case Else
            Err.Raise vbObjectError + 3736, "CCrackWidthCalculator", _
                "SLS.Crack.SigmaSCrcAveragingMode должен быть AllSelected или TensionOnly."
    End Select

    Select Case LCase$(Trim$(mPhi3Mode))
        Case "user", "manual"
            mPhi3Mode = "User"
        Case "auto"
            mPhi3Mode = "Auto"
        Case Else
            Err.Raise vbObjectError + 3709, "CCrackWidthCalculator", "SLS.Crack.Phi3Mode должен быть User или Auto."
    End Select

    Select Case LCase$(Trim$(mTensionZoneMode))
        Case "effective": mTensionZoneMode = "Effective"
        Case "fulltension": mTensionZoneMode = "FullTension"
        Case Else: Err.Raise vbObjectError + 3708, "CCrackWidthCalculator", "SLS.Crack.TensionZoneMode должен быть Effective или FullTension."
    End Select

    Select Case LCase$(Trim$(mCoverMode))
        Case "nearestcontour", "nearest", "local"
            mCoverMode = COVER_MODE_NEAREST_CONTOUR
        Case "globalextreme", "global", "extreme"
            mCoverMode = COVER_MODE_GLOBAL_EXTREME
        Case Else
            Err.Raise vbObjectError + 3721, "CCrackWidthCalculator", "SLS.Crack.CoverDistanceMode должен быть NearestContour или GlobalExtreme."
    End Select


    If mAllowableCrackWidth <= 0# Then Err.Raise vbObjectError + 3710, "CCrackWidthCalculator", "SLS.Crack.Allowable должен быть положительным."
    If mPhi1Input <= 0# Then Err.Raise vbObjectError + 3717, "CCrackWidthCalculator", "SLS.Crack.Phi1 должен быть положительным."
    If mPhi2Input <= 0# Then Err.Raise vbObjectError + 3718, "CCrackWidthCalculator", "SLS.Crack.Phi2 должен быть положительным."
    If mPhi3Input <= 0# Then Err.Raise vbObjectError + 3719, "CCrackWidthCalculator", "SLS.Crack.Phi3 должен быть положительным."
    If mPsiSInput <= 0# Then Err.Raise vbObjectError + 3720, "CCrackWidthCalculator", "SLS.Crack.PsiS должен быть положительным."
    If mSteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackWidthCalculator", "Steel.Es должен быть положительным для расчета трещин."
End Sub



' Проверяет расчетный контракт класса.
Private Sub ValidateInputs(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider)
    If currentSolver Is Nothing Then Err.Raise vbObjectError + 3700, "CCrackWidthCalculator", "Результат решателя сечения не передан."
    If Not currentSolver.Converged Then Err.Raise vbObjectError + 3701, "CCrackWidthCalculator", "Результат решателя сечения не сошелся."
    If section Is Nothing Then Err.Raise vbObjectError + 3702, "CCrackWidthCalculator", "Модель расчетного сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3703, "CCrackWidthCalculator", "В модели расчетного сечения нет бетонных элементов."
    If section.RebarCount <= 0 Then Err.Raise vbObjectError + 3704, "CCrackWidthCalculator", "В модели расчетного сечения нет арматурных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3705, "CCrackWidthCalculator", "Поставщик материалов не передан."
End Sub



' Добавляет строку в диагностический журнал.
Private Sub AppendDiagnostic(ByVal text As String)
    If Not mDiagnosticsEnabled Then Exit Sub
    mDiagnosticLog = mDiagnosticLog & text & vbCrLf
End Sub



' Форматирует числа диагностического журнала через точку.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    FormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
End Function



' Утилитарный минимум без WorksheetFunction в расчетном ядре трещин.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function



' Утилитарный максимум без WorksheetFunction в расчетном ядре трещин.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function



' --------------------------------------------------------------------------
' Свойства результатов
' --------------------------------------------------------------------------
' Writer-ы и тесты читают эти значения после Calculate. Свойства ничего не
' пересчитывают и не обращаются к solver-у повторно.

Public Property Get Converged() As Boolean
    Converged = mConverged
End Property



Public Property Get StopReason() As String
    StopReason = mStopReason
End Property



Public Property Get CrackWidthInternalStatus() As EResultInternalStatus
    CrackWidthInternalStatus = mWidthInternalStatus
End Property



Public Property Get CrackWidthResultCode() As EResultCode
    CrackWidthResultCode = mWidthResultCode
End Property



Public Property Get CrackFormed() As Boolean
    CrackFormed = mCrackFormed
End Property



Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mTensionRebarCount
End Property



Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property



Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property



Public Property Get CrackWidth() As Double
    CrackWidth = mCrackWidth
End Property



Public Property Get AllowableCrackWidth() As Double
    AllowableCrackWidth = mAllowableCrackWidth
End Property



Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property



Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property



Public Property Get PsiMode() As String
    PsiMode = mPsiMode
End Property



Public Property Get SigmaSCrcAveragingMode() As String
    SigmaSCrcAveragingMode = mSigmaSCrcAveragingMode
End Property



Public Property Get TensionZoneMode() As String
    TensionZoneMode = mTensionZoneMode
End Property



Public Property Get CoverMode() As String
    CoverMode = mCoverMode
End Property



Public Property Get Phi1() As Double
    Phi1 = mPhi1
End Property



Public Property Get Phi2() As Double
    Phi2 = mPhi2
End Property



Public Property Get Phi3() As Double
    Phi3 = mPhi3
End Property



Public Property Get PsiS() As Double
    PsiS = mPsiS
End Property



Public Property Get SigmaS() As Double
    SigmaS = mSigmaS
End Property



Public Property Get SigmaSCrc() As Double
    SigmaSCrc = mSigmaSCrc
End Property



Public Property Get SteelEs() As Double
    SteelEs = mSteelEs
End Property



Public Property Get AsTension() As Double
    AsTension = mAs
End Property



Public Property Get Abt() As Double
    Abt = mAbt
End Property



Public Property Get DsEquivalent() As Double
    DsEquivalent = mDsEquivalent
End Property



Public Property Get CrackSpacingRaw() As Double
    CrackSpacingRaw = mLsRaw
End Property



Public Property Get CrackSpacing() As Double
    CrackSpacing = mLs
End Property



Public Property Get TensionRebarIds() As String
    TensionRebarIds = mTensionRebarIds
End Property



Public Property Get SectionDepthH() As Double
    SectionDepthH = mSectionDepthH
End Property



Public Property Get CoverA() As Double
    CoverA = mCoverA
End Property



Public Property Get TensionDepth() As Double
    TensionDepth = mTensionDepth
End Property



Public Property Get EffectiveZoneDepth() As Double
    EffectiveZoneDepth = mEffectiveZoneDepth
End Property



Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property



Public Property Let AllowableCrackWidth(ByVal value As Double)
    mAllowableCrackWidth = value
End Property



Public Property Let PsiMode(ByVal value As String)
    mPsiMode = value
End Property



Public Property Let SigmaSCrcAveragingMode(ByVal value As String)
    mSigmaSCrcAveragingMode = value
End Property



Public Property Let Phi1(ByVal value As Double)
    mPhi1Input = value
End Property



Public Property Let Phi2(ByVal value As Double)
    mPhi2Input = value
End Property



Public Property Let Phi3Mode(ByVal value As String)
    mPhi3Mode = value
End Property



Public Property Let Phi3(ByVal value As Double)
    mPhi3Input = value
End Property



Public Property Let PsiS(ByVal value As Double)
    mPsiSInput = value
End Property



Public Property Let TensionZoneMode(ByVal value As String)
    mTensionZoneMode = value
End Property



Public Property Let CoverMode(ByVal value As String)
    mCoverMode = value
End Property



Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

' ============================== ДЛЯ ТЕСТОВ ==============================

' Проверяет подписанные входы действующей формулы psi без подмены найденных
' НДС и напряжений. Пользовательский режим User этот расчет не использует.
Public Function Audit02AutoPsiForTests(ByVal sigmaS As Double, ByVal sigmaSCrc As Double) As Double
    Audit02AutoPsiForTests = AutoPsi(sigmaS, sigmaSCrc)
End Function
