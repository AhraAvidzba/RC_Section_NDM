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
' Единая точка входа расчёта ширины нормальной трещины
' ==========================================================================
' Хранит настройки и результат, проверяет общие зависимости и выбирает
' CCrackWidthSP35Calculator или CCrackWidthSP63Calculator. Нормативный алгоритм
' получает состояние ByRef: подготовленные данные не копируются и не читаются из Excel.
' Нагрузку образования трещины определяет отдельный CCrackFormationCalculator.

Private mWork As TCrackWidthWork ' Единое состояние настроек, подготовки и результата.
Private mSP35 As CCrackWidthSP35Calculator ' Нормативный алгоритм СП 35 без собственного изменяемого состояния.
Private mSP63 As CCrackWidthSP63Calculator ' Нормативный алгоритм СП 63 без собственного изменяемого состояния.


' ------------------------------
' Настройки по умолчанию
' ------------------------------
' Значения соответствуют текущей принятой постановке: длительное раскрытие,
' эффективная зона СП и пользовательские коэффициенты с нормативными defaults.
Private Sub Class_Initialize()
    Set mSP35 = New CCrackWidthSP35Calculator
    Set mSP63 = New CCrackWidthSP63Calculator
    mWork.StandardCode = "SP63"
    mWork.SP35GroupGapTolerance = 10#: mWork.SP35RowTolerance = 50#
    mWork.SP35NeighborRatioLimit = 0.2
    mWork.SP35RadiusDiameterMode = "Max": mWork.SP35InteractionRadiusMode = "6d"
    mWork.SP35RebarProfile = "Periodic"
    mWork.AllowableCrackWidth = 0.3
    mWork.PsiMode = CRACK_PSI_MODE_ALWAYS_CALC
    mWork.SigmaSCrcAveragingMode = CRACK_SIGMA_S_CRC_AVERAGE_ALL_SELECTED
    mWork.Phi3Mode = "Auto"
    mWork.TensionZoneMode = "Effective"
    mWork.CoverMode = CRACK_COVER_MODE_GLOBAL_EXTREME
    mWork.Phi1Input = CRACK_PHI1_LONG_TERM
    mWork.Phi2Input = CRACK_PHI2_DEFAULT
    mWork.Phi3Input = CRACK_PHI3_BENDING
    mWork.PsiSInput = 1#
    mWork.SteelEs = 200000#
    mWork.DiagnosticsEnabled = True
End Sub


' Сбрасывает данные подготовки формулы предыдущего сочетания.
Private Sub ClearResult()
    mWork.Converged = False
    mWork.StopReason = vbNullString
    mWork.DiagnosticLog = vbNullString
    mWork.WidthInternalStatus = rsUnset
    mWork.WidthResultCode = rcNone
    mWork.FormulaCalculated = False
    mWork.CrackWidth = 0#
    mWork.Utilization = 0#
    mWork.CrackFormed = False
    mWork.TensionRebarCount = 0
    mWork.MaxSteelStrain = 0#
    mWork.MaxSteelStress = 0#
    mWork.SigmaS = 0#
    mWork.SigmaSCrc = 0#
    mWork.RebarArea = 0#
    mWork.Abt = 0#
    mWork.DsEquivalent = 0#
    mWork.LsRaw = 0#
    mWork.Ls = 0#
    mWork.SpacingLimitWarning = vbNullString: mWork.HasSpacingLimitWarning = False
    mWork.PsiS = mWork.PsiSInput
    mWork.Phi1 = mWork.Phi1Input
    mWork.Phi2 = mWork.Phi2Input
    mWork.Phi3 = 0#
    mWork.TensionRebarIds = vbNullString
    Set mWork.CrackedStateSpec = Nothing
    mWork.SectionDepthH = 0#
    mWork.CoverA = 0#
    mWork.TensionDepth = 0#
    mWork.EffectiveZoneDepth = 0#
    mWork.CentralTensionBranch = False
    Set mWork.Formation = Nothing
    mWork.FallbackPsi1 = False
    mWork.PsiFallbackReason = vbNullString
    mWork.PostHasSelectedTension = False
    Set mWork.GeometryQuery = Nothing
    Set mWork.InteractionRegion = Nothing
    Set mWork.SP35Data = Nothing
    mWork.SP35CriticalCandidate = 0: mWork.SP35Radius = 0#: mWork.SP35PsiCm = 0#
End Sub


' Принимает готовую проверку образования трещины без нового решения НДС.
' Если пригодная пороговая точка отсутствует, проверка ширины остаётся доступной
' по текущему НДС. Правила коэффициентов определяет выбранный нормативный калькулятор.
' Подтверждённое отсутствие трещины NotCracked
' исключает расчёт ширины независимо от режима коэффициентов.
Public Sub InitializeForFormation(ByVal formation As CCrackFormationResult)
    ClearResult
    Set mWork.Formation = formation
    If formation Is Nothing Then
        CrackSetWidthResult mWork, rsInternalError, rcInternalError, "Не передан результат образования трещины."
        Exit Sub
    End If
    ValidateSettings
    If formation.ConfirmedNotCracked Then
        mWork.Converged = True
        mWork.CentralTensionBranch = formation.CentralTensionBranch
        CrackSetWidthResult mWork, rsNotApplicable, rcCrackNotFormed, formation.ResultMeta.ResultComment
        Exit Sub
    End If
    mWork.Converged = True
    mWork.CrackFormed = formation.CrackFormed Or Not formation.HasUsableFormationPoint
    mWork.CentralTensionBranch = formation.CentralTensionBranch
    CrackSetWidthResult mWork, rsNotApplicable, rcCrackNotFormed, formation.StopReason
End Sub


' Готовит область, арматуру и коэффициенты по сохранённым состояниям.
' Нагрузку образования трещины и НДС до и после неё заново не ищет.
' Формула получает только числа; геометрические операции использует
' переданные на один вызов или создаёт самостоятельно при отдельном вызове класса.
Public Sub Calculate(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider, ByVal crackedStateSpec As CMaterialModelSpec, _
        ByVal formation As CCrackFormationResult, ByVal loadStateForClassification As CSectionLoadState, _
        Optional ByVal geometryQuery As CSectionGeometryQuery = Nothing)
    InitializeForFormation formation
    If formation Is Nothing Then Exit Sub
    If Not mWork.Converged Or Not mWork.CrackFormed Then Exit Sub
    If currentSolver Is Nothing Then
        mWork.Converged = False
        CrackSetWidthResult mWork, rsBlockedByDependency, rcBlockedByDependency, "Для ширины трещины не передано текущее НДС сочетания."
        Exit Sub
    End If
    If Not currentSolver.Converged Or Not currentSolver.WithinPhysicalRange Then
        mWork.Converged = False
        CrackSetWidthResult mWork, rsBlockedByDependency, rcBlockedByDependency, "Ширина трещины не рассчитана: текущее НДС не найдено или находится за физическими пределами материалов."
        Exit Sub
    End If
    ValidateInputs currentSolver, section, materialProvider
    If crackedStateSpec Is Nothing Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "Не передана модель материалов для НДС с трещинами (MaterialModel.CrackedState)."
    If Not crackedStateSpec.IsComplete Then _
        Err.Raise vbObjectError + 3714, "CCrackWidthCalculator", "Модель материалов для НДС с трещинами заполнена не полностью (MaterialModel.CrackedState)."
    Set mWork.CrackedStateSpec = crackedStateSpec.Clone
    mWork.SteelEs = materialProvider.SteelElasticModulusFromSpec(mWork.CrackedStateSpec)
    ValidateSettings
    mWork.Converged = False
    mWork.CrackFormed = False
    mWork.StopReason = vbNullString
    mWork.WidthInternalStatus = rsUnset
    mWork.WidthResultCode = rcNone
    Set mWork.GeometryQuery = geometryQuery
    If mWork.GeometryQuery Is Nothing Then Set mWork.GeometryQuery = New CSectionGeometryQuery
    mWork.GeometryQuery.EnsureSection section
    If mWork.StandardCode = "SP35" Then
        mSP35.CalculateSP35 mWork, currentSolver, section, materialProvider.SteelMaterialFromSpec(mWork.CrackedStateSpec)
    Else
        mSP63.Calculate mWork, currentSolver, section, materialProvider, loadStateForClassification
    End If
End Sub


' Возвращает готовый результат образования трещины без отдельного копирования его полей.
Public Property Get FormationResult() As CCrackFormationResult
    Set FormationResult = mWork.Formation
End Property


' ------------------------------
' Чтение настроек Config
' ------------------------------
' Загружает настройки подготовки формулы ширины. Поиск НДС и параметры
' трещинообразования настраиваются у соответствующих владельцев; допустимая
' ширина здесь переводится в миллиметры внутренней системы единиц. Проверяются
' обязательные настройки выбранного СП; неверный ввод не заменяется начальными значениями.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub

    mWork.StandardCode = settings.GetRequiredChoice("SLS.Crack.Code", Array("SP63", "SP35"))
    mWork.AllowableCrackWidth = settings.GetRequiredDouble("SLS.Crack.Allowable")
    If mWork.StandardCode = "SP35" Then
        mWork.SP35GroupGapTolerance = settings.GetRequiredDouble("SLS.Crack.SP35.GroupGapTolerance")
        mWork.SP35RowTolerance = settings.GetRequiredDouble("SLS.Crack.SP35.RowTolerance")
        mWork.SP35NeighborRatioLimit = settings.GetRequiredDouble("SLS.Crack.SP35.NeighborRatioLimit")
        mWork.SP35RadiusDiameterMode = settings.GetRequiredChoice("SLS.Crack.SP35.RadiusDiameterMode", Array("Max", "Min", "Average"))
        mWork.SP35InteractionRadiusMode = settings.GetRequiredChoice("SLS.Crack.SP35.InteractionRadiusMode", Array("3d", "5d", "6d"))
        mWork.SP35RebarProfile = settings.GetRequiredChoice("SLS.Crack.SP35.RebarProfile", Array("Periodic", "Smooth"))
    Else
        mWork.PsiMode = settings.GetRequiredString("SLS.Crack.PsiMode")
        mWork.SigmaSCrcAveragingMode = settings.GetRequiredString("SLS.Crack.SigmaSCrcAveragingMode")
        mWork.TensionZoneMode = settings.GetRequiredString("SLS.Crack.TensionZoneMode")
        mWork.CoverMode = settings.GetRequiredString("SLS.Crack.CoverDistanceMode")
        mWork.Phi1Input = settings.GetRequiredDouble("SLS.Crack.Phi1")
        mWork.Phi2Input = settings.GetRequiredDouble("SLS.Crack.Phi2")
        mWork.Phi3Mode = settings.GetRequiredString("SLS.Crack.Phi3Mode")
        mWork.Phi3Input = settings.GetRequiredDouble("SLS.Crack.Phi3")
        mWork.PsiSInput = settings.GetRequiredDouble("SLS.Crack.PsiS")
    End If

    If Not units Is Nothing Then
        mWork.AllowableCrackWidth = units.InputLengthToInternal(mWork.AllowableCrackWidth)
        If mWork.StandardCode = "SP35" Then
            mWork.SP35GroupGapTolerance = units.InputLengthToInternal(mWork.SP35GroupGapTolerance)
            mWork.SP35RowTolerance = units.InputLengthToInternal(mWork.SP35RowTolerance)
        End If
    End If
    ValidateSettings settings
End Sub


' Вычисляет a_crc только по переданным коэффициентам, напряжению,
' модулю и расстоянию между трещинами. Состав арматуры, НДС
' и статус проверки этим методом не определяются.
Public Function CrackWidthFromData(ByVal phi1 As Double, ByVal phi2 As Double, _
        ByVal phi3 As Double, ByVal psiS As Double, ByVal sigmaS As Double, _
        ByVal steelEs As Double, ByVal crackSpacing As Double) As Double
    CrackWidthFromData = mSP63.CrackWidthFromData(phi1, phi2, phi3, psiS, sigmaS, steelEs, crackSpacing)
End Function


' Возвращает численное отношение готового раскрытия к допускаемому.
' При нулевом пределе отношение не определено; решение о применимости
' проверки принимается до вызова этой чистой операции.
Public Function UtilizationFromData(ByVal crackWidthValue As Double, ByVal allowableCrackWidth As Double) As Double
    UtilizationFromData = CrackUtilizationFromData(crackWidthValue, allowableCrackWidth)
End Function


' ------------------------------ ЧИСТЫЕ ФОРМУЛЫ СП 35 ------------------------------
' [SP35] П. 7.110, пособие (4.310): вычисляет радиус армирования, мм.
' Область взаимодействия и состав групп подготовлены заранее;
' формула получает числа и не выполняет геометрический поиск или решение НДС.
Public Function SP35ReinforcementRadiusFromData(ByVal interactionArea As Double, ByVal betaDiameterSum As Double) As Double
    SP35ReinforcementRadiusFromData = mSP35.SP35ReinforcementRadiusFromData(interactionArea, betaDiameterSum)
End Function


' [SP35] П. 7.109, пособие (4.308)-(4.309): psi имеет размерность см.
' Нормативные коэффициенты используют Rr в см; входной радиус хранится в мм.
' Строка профиля уже проверена слоем настроек, неизвестный вариант не заменяется.
Public Function SP35PsiFromData(ByVal reinforcementRadiusMm As Double, ByVal rebarProfile As String) As Double
    SP35PsiFromData = mSP35.SP35PsiFromData(reinforcementRadiusMm, rebarProfile)
End Function


' [SP35] П. 7.105, пособие (4.306): a_cr=(sigma_s/Es)*psi.
' Напряжение и модуль задаются в МПа, psi — в см; результат возвращается в мм.
' До вызова выбираются наибольшее напряжение проверяемой группы и соответствующий модуль.
Public Function SP35CrackWidthFromData(ByVal sigmaS As Double, ByVal steelEs As Double, ByVal psiCm As Double) As Double
    SP35CrackWidthFromData = mSP35.SP35CrackWidthFromData(sigmaS, steelEs, psiCm)
End Function


' Возвращает статус именно выполнения формулы ширины. Получение НДС
' или проверка образования трещины не считается выполненной формулой.
' Отсутствие обязательного итога отмечается как внутренняя ошибка программы.
Public Property Get ResultMeta() As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    If mWork.WidthInternalStatus = rsUnset Then
        meta.SetResult rsInternalError, rcInternalError, rkCrackWidth, "Результат расчета ширины трещины не заполнен.", _
            vbNullString, True, False
    Else
        meta.SetResult mWork.WidthInternalStatus, mWork.WidthResultCode, rkCrackWidth, mWork.StopReason, _
            vbNullString, True, mWork.FormulaCalculated
    End If
    Set ResultMeta = meta
End Property


' Проверяет режимы и коэффициенты до решения НДС и до выхода «трещины нет».
' При чтении Config ошибка содержит фактический адрес ячейки.
' Равновесие не решается и готовый результат этой проверкой не меняется.
Public Sub ValidateSettings(Optional ByVal settings As CSystemSettingsReader = Nothing)
    Select Case UCase$(Trim$(mWork.StandardCode))
        Case "SP63", "SP35": mWork.StandardCode = UCase$(Trim$(mWork.StandardCode))
        Case Else: RaiseInvalidSetting 3786, "SLS.Crack.Code", "Неизвестная методика ширины трещин.", "Выберите SP63 или SP35.", settings
    End Select
    If mWork.AllowableCrackWidth <= 0# Then RaiseInvalidSetting 3710, "SLS.Crack.Allowable", "Допустимая ширина должна быть положительной.", "Введите ширину больше нуля в единицах INPUT Length.", settings
    If mWork.SteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackWidthCalculator", "Steel.Es должен быть положительным для расчета трещин."
    If mWork.StandardCode = "SP35" Then
        ValidateSP35Settings settings
        Exit Sub
    End If
    Select Case LCase$(Trim$(mWork.PsiMode))
        Case "user", "manual", "fixed1", "unity"
            mWork.PsiMode = CRACK_PSI_MODE_USER
        Case "auto", "refined"
            mWork.PsiMode = CRACK_PSI_MODE_AUTO
        Case "alwayscalc", "alwayscalculate", "always", "calc"
            mWork.PsiMode = CRACK_PSI_MODE_ALWAYS_CALC
        Case Else
            RaiseInvalidSetting 3707, "SLS.Crack.PsiMode", "Неизвестный режим коэффициента psi_s.", "Выберите User, Auto или AlwaysCalc.", settings
    End Select

    Select Case LCase$(Trim$(mWork.SigmaSCrcAveragingMode))
        Case vbNullString, "allselected", "all", "includecompression", "withcompression", "signed"
            mWork.SigmaSCrcAveragingMode = CRACK_SIGMA_S_CRC_AVERAGE_ALL_SELECTED
        Case "tensiononly", "tension", "positiveonly", "old"
            mWork.SigmaSCrcAveragingMode = CRACK_SIGMA_S_CRC_AVERAGE_TENSION_ONLY
        Case Else
            RaiseInvalidSetting 3736, "SLS.Crack.SigmaSCrcAveragingMode", _
                "Неизвестный способ усреднения напряжения сразу после образования трещины.", "Выберите AllSelected или TensionOnly.", settings
    End Select

    Select Case LCase$(Trim$(mWork.Phi3Mode))
        Case "user", "manual"
            mWork.Phi3Mode = "User"
        Case "auto"
            mWork.Phi3Mode = "Auto"
        Case Else
            RaiseInvalidSetting 3709, "SLS.Crack.Phi3Mode", "Неизвестный режим коэффициента phi3.", "Выберите User или Auto.", settings
    End Select

    Select Case LCase$(Trim$(mWork.TensionZoneMode))
        Case "effective": mWork.TensionZoneMode = "Effective"
        Case "fulltension": mWork.TensionZoneMode = "FullTension"
        Case Else: RaiseInvalidSetting 3708, "SLS.Crack.TensionZoneMode", "Неизвестный способ выбора растянутой зоны.", "Выберите Effective или FullTension.", settings
    End Select

    Select Case LCase$(Trim$(mWork.CoverMode))
        Case "nearestcontour", "nearest", "local"
            mWork.CoverMode = CRACK_COVER_MODE_NEAREST_CONTOUR
        Case "globalextreme", "global", "extreme"
            mWork.CoverMode = CRACK_COVER_MODE_GLOBAL_EXTREME
        Case Else
            RaiseInvalidSetting 3721, "SLS.Crack.CoverDistanceMode", "Неизвестный способ определения расстояния до границы бетона.", "Выберите NearestContour или GlobalExtreme.", settings
    End Select


    If mWork.AllowableCrackWidth <= 0# Then RaiseInvalidSetting 3710, "SLS.Crack.Allowable", "Допустимая ширина должна быть положительной.", "Введите ширину больше нуля в единицах INPUT Length.", settings
    If mWork.Phi1Input <= 0# Then RaiseInvalidSetting 3717, "SLS.Crack.Phi1", "Коэффициент должен быть положительным.", "Введите коэффициент больше нуля.", settings
    If mWork.Phi2Input <= 0# Then RaiseInvalidSetting 3718, "SLS.Crack.Phi2", "Коэффициент должен быть положительным.", "Введите коэффициент больше нуля.", settings
    If mWork.Phi3Input <= 0# Then RaiseInvalidSetting 3719, "SLS.Crack.Phi3", "Коэффициент должен быть положительным.", "Введите коэффициент больше нуля.", settings
    If mWork.PsiSInput <= 0# Then RaiseInvalidSetting 3720, "SLS.Crack.PsiS", "Коэффициент должен быть положительным.", "Введите коэффициент больше нуля.", settings
    If mWork.SteelEs <= 0# Then Err.Raise vbObjectError + 3713, "CCrackWidthCalculator", "Steel.Es должен быть положительным для расчета трещин."
End Sub


' Проверяет применяемые настройки СП 35 и при ошибке указывает адрес Config.
' Выбранный справочный диаметр не заменяет реальные диаметры
' в сумме beta*nd; допуски длин уже переданы в мм.
Private Sub ValidateSP35Settings(ByVal settings As CSystemSettingsReader)
    If mWork.SP35GroupGapTolerance < 0# Then RaiseInvalidSetting 3787, "SLS.Crack.SP35.GroupGapTolerance", "Чистый зазор не может быть отрицательным.", "Введите неотрицательную длину в единицах INPUT Length.", settings
    If mWork.SP35RowTolerance < 0# Then RaiseInvalidSetting 3788, "SLS.Crack.SP35.RowTolerance", "Допуск ряда не может быть отрицательным.", "Введите неотрицательную длину в единицах INPUT Length.", settings
    If mWork.SP35NeighborRatioLimit < 0# Or mWork.SP35NeighborRatioLimit > 1# Then _
        RaiseInvalidSetting 3792, "SLS.Crack.SP35.NeighborRatioLimit", "Порог отношения t/L должен быть от 0 до 1.", "Введите число от 0 до 1; начальное значение 0.2.", settings
    Select Case LCase$(Trim$(mWork.SP35RadiusDiameterMode))
        Case "max": mWork.SP35RadiusDiameterMode = "Max"
        Case "min": mWork.SP35RadiusDiameterMode = "Min"
        Case "average": mWork.SP35RadiusDiameterMode = "Average"
        Case Else: RaiseInvalidSetting 3789, "SLS.Crack.SP35.RadiusDiameterMode", "Неизвестный способ выбора диаметра радиуса.", "Выберите Max, Min или Average.", settings
    End Select
    Select Case LCase$(Trim$(mWork.SP35InteractionRadiusMode))
        Case "3d", "5d", "6d": mWork.SP35InteractionRadiusMode = LCase$(Trim$(mWork.SP35InteractionRadiusMode))
        Case Else: RaiseInvalidSetting 3790, "SLS.Crack.SP35.InteractionRadiusMode", "Неизвестный радиус взаимодействия.", "Выберите 3d, 5d или 6d.", settings
    End Select
    Select Case LCase$(Trim$(mWork.SP35RebarProfile))
        Case "periodic": mWork.SP35RebarProfile = "Periodic"
        Case "smooth": mWork.SP35RebarProfile = "Smooth"
        Case Else: RaiseInvalidSetting 3791, "SLS.Crack.SP35.RebarProfile", "Неизвестный профиль арматуры.", "Выберите Periodic или Smooth.", settings
    End Select
End Sub


' Сохраняет код ошибки инженерной настройки. CSystemSettingsReader добавляет адрес
' ввода, но допустимость значения проверяет расчёт. При программном вызове
' без него сохраняются тот же ключ параметра и понятное пояснение.
Private Sub RaiseInvalidSetting(ByVal errorCode As Long, ByVal key As String, _
        ByVal reason As String, ByVal correction As String, ByVal settings As CSystemSettingsReader)
    Dim message As String
    If settings Is Nothing Then
        message = "На листе Config в строке " & key & ": " & reason & " " & correction
    Else
        message = settings.InputErrorMessage(key, reason, correction)
    End If
    Err.Raise vbObjectError + errorCode, "CCrackWidthCalculator", message
End Sub


' Проверяет, переданы ли обязательные модель и материалы для расчёта ширины трещины.
Private Sub ValidateInputs(ByVal currentSolver As CSectionStateResult, ByVal section As CSectionModel, _
        ByVal materialProvider As CMaterialModelProvider)
    If currentSolver Is Nothing Then Err.Raise vbObjectError + 3700, "CCrackWidthCalculator", "Результат решателя сечения не передан."
    If Not currentSolver.Converged Then Err.Raise vbObjectError + 3701, "CCrackWidthCalculator", "Результат решателя сечения не сошелся."
    If section Is Nothing Then Err.Raise vbObjectError + 3702, "CCrackWidthCalculator", "Модель расчетного сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 3703, "CCrackWidthCalculator", "В модели расчетного сечения нет бетонных элементов."
    If section.RebarCount <= 0 And mWork.StandardCode <> "SP35" Then Err.Raise vbObjectError + 3704, "CCrackWidthCalculator", "В модели расчетного сечения нет арматурных элементов."
    If materialProvider Is Nothing Then Err.Raise vbObjectError + 3705, "CCrackWidthCalculator", "Поставщик материалов не передан."
End Sub


' --------------------------------------------------------------------------
' Свойства результатов
' --------------------------------------------------------------------------
' Возвращает сохранённые после Calculate значения для вывода и тестов
' без повторного расчёта и решения равновесия.

Public Property Get Converged() As Boolean
    Converged = mWork.Converged
End Property


' Инженерная причина результата или предупреждение по готовым входным состояниям.
Public Property Get StopReason() As String
    StopReason = mWork.StopReason
End Property


' Машинный статус формульного расчета ширины нормальной трещины.
Public Property Get CrackWidthInternalStatus() As EResultInternalStatus
    CrackWidthInternalStatus = mWork.WidthInternalStatus
End Property


' Машинный код причины результата формульной ширины трещины.
Public Property Get CrackWidthResultCode() As EResultCode
    CrackWidthResultCode = mWork.WidthResultCode
End Property


' True, если по выбранной ветке расчета трещина считается образованной.
Public Property Get CrackFormed() As Boolean
    CrackFormed = mWork.CrackFormed
End Property


' Сколько стержней попало в расчетную растянутую арматуру As.
Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mWork.TensionRebarCount
End Property


' Максимальная растягивающая деформация среди выбранных стержней As.
Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mWork.MaxSteelStrain
End Property


' Напряжение в стержне с максимальной растягивающей деформацией, МПа.
Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mWork.MaxSteelStress
End Property


' Расчетная ширина раскрытия a_crc, мм.
Public Property Get CrackWidth() As Double
    CrackWidth = mWork.CrackWidth
End Property


' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
Public Property Get AllowableCrackWidth() As Double
    AllowableCrackWidth = mWork.AllowableCrackWidth
End Property


' Коэффициент использования по трещине: a_crc / a_crc,ult.
Public Property Get Utilization() As Double
    Utilization = mWork.Utilization
End Property


' Текстовая диагностика выбора зоны, арматуры, ls и psi_s.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mWork.DiagnosticLog
End Property


' User, Auto или AlwaysCalc: как принимать psi_s в формуле раскрытия.
Public Property Get PsiMode() As String
    PsiMode = mWork.PsiMode
End Property


' Как усреднять sigma_s,crc по ранее выбранной группе стержней.
Public Property Get SigmaSCrcAveragingMode() As String
    SigmaSCrcAveragingMode = mWork.SigmaSCrcAveragingMode
End Property


' Effective или FullTension: способ задания бетонной зоны Abt.
Public Property Get TensionZoneMode() As String
    TensionZoneMode = mWork.TensionZoneMode
End Property


' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
Public Property Get CoverMode() As String
    CoverMode = mWork.CoverMode
End Property


' Коэффициент длительности действия нагрузки.
Public Property Get Phi1() As Double
    Phi1 = mWork.Phi1
End Property


' Пользовательский коэффициент, учитывающий поверхность арматуры.
Public Property Get Phi2() As Double
    Phi2 = mWork.Phi2
End Property


' Коэффициент характера нагружения.
Public Property Get Phi3() As Double
    Phi3 = mWork.Phi3
End Property


' Коэффициент неравномерности деформаций арматуры между трещинами.
Public Property Get PsiS() As Double
    PsiS = mWork.PsiS
End Property


' Средневзвешенное напряжение выбранной растянутой арматуры при текущей нагрузке, МПа.
Public Property Get SigmaS() As Double
    SigmaS = mWork.SigmaS
End Property


' Среднее той же группы в AfterMcrcState выбранного пути с учетом режима усреднения, МПа.
Public Property Get SigmaSCrc() As Double
    SigmaSCrc = mWork.SigmaSCrc
End Property


' Возвращает независимую копию готовой расчётной области без повторного построения.
' Вызывающий код не может изменить область, сохранённую в результате проверки.
Public Property Get InteractionRegion() As CConcreteRegion
    If Not mWork.InteractionRegion Is Nothing Then Set InteractionRegion = mWork.InteractionRegion.Clone
End Property


' Возвращает защищённые от изменений данные СП 35 для вывода всех рядов
' и расчётных величин выбранной области взаимодействия.
Public Property Get SP35Data() As CSP35CrackData
    Set SP35Data = mWork.SP35Data
End Property


' Возвращает номер области с наибольшей шириной трещины; наибольшая площадь не является
' критерием выбора.
Public Property Get SP35CriticalCandidate() As Long
    SP35CriticalCandidate = mWork.SP35CriticalCandidate
End Property


' Rr в мм; размерный psi в см, не psi_s СП 63.
Public Property Get SP35ReinforcementRadius() As Double
    SP35ReinforcementRadius = mWork.SP35Radius
End Property


' Рассчитанный коэффициент psi_cm выбранной области СП35.
Public Property Get SP35PsiCm() As Double
    SP35PsiCm = mWork.SP35PsiCm
End Property


' Новые настройки применяются при следующем Calculate, не меняя готовые результаты.
' При программном вызове численные длины передаются во внутренних мм.
Public Property Get StandardCode() As String
    StandardCode = mWork.StandardCode
End Property


' SP63 или SP35: только выбор формулы ширины, не модели НДС.
Public Property Let StandardCode(ByVal value As String)
    mWork.StandardCode = value
End Property


' Чистый зазор объединения реальных стержней, мм.
Public Property Let SP35GroupGapTolerance(ByVal value As Double)
    mWork.SP35GroupGapTolerance = value
End Property


' Максимальная ширина ряда по нормали к Н.О., мм.
Public Property Let SP35RowTolerance(ByVal value As Double)
    mWork.SP35RowTolerance = value
End Property


' Безразмерный порог t/L, не допуск группировки рядов.
Public Property Let SP35NeighborRatioLimit(ByVal value As Double)
    mWork.SP35NeighborRatioLimit = value
End Property


' Max/Min/Average только для геометрических радиусов.
Public Property Let SP35RadiusDiameterMode(ByVal value As String)
    mWork.SP35RadiusDiameterMode = value
End Property


' 3d/5d/6d: нормативный множитель диаметра.
Public Property Let SP35InteractionRadiusMode(ByVal value As String)
    mWork.SP35InteractionRadiusMode = value
End Property


' Periodic/Smooth выбирает формулу psi СП 35.
Public Property Let SP35RebarProfile(ByVal value As String)
    mWork.SP35RebarProfile = value
End Property


' Es для формулы раскрытия трещины, МПа.
Public Property Get SteelEs() As Double
    SteelEs = mWork.SteelEs
End Property


' Площадь выбранной растянутой арматуры, мм2.
Public Property Get AsTension() As Double
    AsTension = mWork.RebarArea
End Property


' Площадь бетона, попавшая в расчетную растянутую зону, мм2.
Public Property Get Abt() As Double
    Abt = mWork.Abt
End Property


' Эквивалентный диаметр выбранных стержней для разных d, мм.
Public Property Get DsEquivalent() As Double
    DsEquivalent = mWork.DsEquivalent
End Property


' Нормативная длина ls до ограничений 10ds/100/40ds/400, мм.
Public Property Get CrackSpacingRaw() As Double
    CrackSpacingRaw = mWork.LsRaw
End Property


' Итоговое расстояние между трещинами после ограничений, мм.
Public Property Get CrackSpacing() As Double
    CrackSpacing = mWork.Ls
End Property


' Имена стержней, вошедших в расчет sigma_s и As, через запятую.
Public Property Get TensionRebarIds() As String
    TensionRebarIds = mWork.TensionRebarIds
End Property


' Габарит сечения по нормали к нейтральной линии по общему бетонному контуру КЭ.
Public Property Get SectionDepthH() As Double
    SectionDepthH = mWork.SectionDepthH
End Property


' Расстояние a_s по выбранному режиму CoverDistanceMode, мм.
Public Property Get CoverA() As Double
    CoverA = mWork.CoverA
End Property


' Фактическая глубина растянутой бетонной зоны до нейтральной линии, мм.
Public Property Get TensionDepth() As Double
    TensionDepth = mWork.TensionDepth
End Property


' Глубина зоны Abt, реально использованная после ограничений, мм.
Public Property Get EffectiveZoneDepth() As Double
    EffectiveZoneDepth = mWork.EffectiveZoneDepth
End Property


' True, если сработала отдельная ветка центрального растяжения.
Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mWork.CentralTensionBranch
End Property


' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
Public Property Let AllowableCrackWidth(ByVal value As Double)
    mWork.AllowableCrackWidth = value
End Property


' User, Auto или AlwaysCalc: как принимать psi_s в формуле раскрытия.
Public Property Let PsiMode(ByVal value As String)
    mWork.PsiMode = value
End Property


' Как усреднять sigma_s,crc по ранее выбранной группе стержней.
Public Property Let SigmaSCrcAveragingMode(ByVal value As String)
    mWork.SigmaSCrcAveragingMode = value
End Property


' Пользовательское значение phi1 для формулы раскрытия.
Public Property Let Phi1(ByVal value As Double)
    mWork.Phi1Input = value
End Property


' Пользовательское значение phi2 для формулы раскрытия.
Public Property Let Phi2(ByVal value As Double)
    mWork.Phi2Input = value
End Property


' User или Auto: брать phi3 из Config или выбирать по характеру продольной силы.
Public Property Let Phi3Mode(ByVal value As String)
    mWork.Phi3Mode = value
End Property


' Пользовательское значение phi3, если Phi3Mode = User.
Public Property Let Phi3(ByVal value As Double)
    mWork.Phi3Input = value
End Property


' Пользовательское значение psi_s, если PsiMode = User.
Public Property Let PsiS(ByVal value As Double)
    mWork.PsiSInput = value
End Property


' Effective или FullTension: способ задания бетонной зоны Abt.
Public Property Let TensionZoneMode(ByVal value As String)
    mWork.TensionZoneMode = value
End Property


' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
Public Property Let CoverMode(ByVal value As String)
    mWork.CoverMode = value
End Property


' Отключает подробный журнал, когда txt-отчет выполнения не нужен.
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
    mWork.DiagnosticsEnabled = value
End Property


' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет подписанные входы действующей формулы psi без подмены найденных
' НДС и напряжений. Пользовательский режим User этот расчет не использует.
Public Function Audit02AutoPsiForTests(ByVal sigmaS As Double, ByVal sigmaSCrc As Double) As Double
    Audit02AutoPsiForTests = mSP63.AutoPsi(sigmaS, sigmaSCrc)
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет только завершение готовой формулы с другим пояснением события;
' НДС, коэффициенты, напряжения, расстояние ls и машинный признак не подменяются.
Friend Sub TestRefinalizeWithSpacingText(ByVal text As String)
    mSP63.TestRefinalizeWithSpacingText mWork, text
End Sub
