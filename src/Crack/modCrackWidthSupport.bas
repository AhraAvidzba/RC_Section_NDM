Attribute VB_Name = "modCrackWidthSupport"
Option Explicit

' ==========================================================================
' Общее состояние подготовки раскрытия трещин и единая диагностика
' ==========================================================================

' Внутренний контракт нормативных алгоритмов: мм, мм2, МПа.
' Фасад владеет состоянием; алгоритмы получают его ByRef только на время вызова.
' Структура не читает Excel и не выполняет поиск равновесия.
Public Type TCrackWidthWork
    Converged As Boolean ' Подготовка данных и формула завершены без ошибки; равновесие этот класс не решает.
    StopReason As String ' Инженерная причина результата или предупреждение по готовым входным состояниям.
    DiagnosticLog As String ' Текстовая диагностика выбора зоны, арматуры, ls и psi_s.
    DiagnosticsEnabled As Boolean ' Отключает подробный журнал, когда txt-отчет выполнения не нужен.
    WidthInternalStatus As EResultInternalStatus ' Машинный статус формульного расчета ширины нормальной трещины.
    WidthResultCode As EResultCode ' Машинный код причины результата формульной ширины трещины.
    FormulaCalculated As Boolean ' Признак фактического выполнения формулы ширины; полученный результат образования трещины ещё не означает расчёт ширины.
    AllowableCrackWidth As Double ' Пользовательский предел a_crc,ult во внутренних единицах длины, мм.
    CrackWidth As Double ' Расчетная ширина раскрытия a_crc, мм.
    Utilization As Double ' Коэффициент использования по трещине: a_crc / a_crc,ult.
    CrackFormed As Boolean ' True, если по выбранной ветке расчета трещина считается образованной.
    PsiMode As String ' User, Auto или AlwaysCalc: как принимать psi_s в формуле раскрытия.
    SigmaSCrcAveragingMode As String ' Как усреднять sigma_s,crc по ранее выбранной группе стержней.
    Phi3Mode As String ' User или Auto: брать phi3 из Config или выбирать по характеру продольной силы.
    TensionZoneMode As String ' Effective или FullTension: способ задания бетонной зоны Abt.
    CoverMode As String ' NearestContour или GlobalExtreme: способ измерения расстояния a_s.
    Phi1Input As Double ' Пользовательское значение phi1 для формулы раскрытия.
    Phi2Input As Double ' Пользовательское значение phi2 для формулы раскрытия.
    Phi3Input As Double ' Пользовательское значение phi3, если Phi3Mode = User.
    PsiSInput As Double ' Пользовательское значение psi_s, если PsiMode = User.
    SteelEs As Double ' Es для формулы раскрытия трещины, МПа.
    CrackedStateSpec As CMaterialModelSpec ' Спецификация модели раскрытой трещины текущего LC.
    TensionRebarCount As Long ' Сколько стержней попало в расчетную растянутую арматуру As.
    MaxSteelStrain As Double ' Максимальная растягивающая деформация среди выбранных стержней As.
    MaxSteelStress As Double ' Напряжение в стержне с максимальной растягивающей деформацией, МПа.
    SigmaS As Double ' Средневзвешенное напряжение выбранной растянутой арматуры при текущей нагрузке, МПа.
    SigmaSCrc As Double ' Среднее той же группы в AfterMcrcState выбранного пути с учетом режима усреднения, МПа.
    RebarArea As Double ' Площадь выбранной растянутой арматуры, мм2.
    Abt As Double ' Площадь бетона, попавшая в расчетную растянутую зону, мм2.
    DsEquivalent As Double ' Эквивалентный диаметр выбранных стержней для разных d, мм.
    LsRaw As Double ' Нормативная длина ls до ограничений 10ds/100/40ds/400, мм.
    Ls As Double ' Итоговое расстояние между трещинами после ограничений, мм.
    SpacingLimitWarning As String ' Объяснение несовместимых границ ls; публикуется только при выполненной формуле.
    HasSpacingLimitWarning As Boolean ' Фактический конфликт нижнего/верхнего ограничения, независимо от текста.
    PsiS As Double ' Коэффициент неравномерности деформаций арматуры между трещинами.
    Phi1 As Double ' Коэффициент длительности действия нагрузки.
    Phi2 As Double ' Пользовательский коэффициент, учитывающий поверхность арматуры.
    Phi3 As Double ' Коэффициент характера нагружения.
    TensionRebarIds As String ' Имена стержней, вошедших в расчет sigma_s и As, через запятую.
    SectionDepthH As Double ' Габарит сечения по нормали к нейтральной линии по общему бетонному контуру КЭ.
    CoverA As Double ' Расстояние a_s по выбранному режиму CoverDistanceMode, мм.
    TensionDepth As Double ' Фактическая глубина растянутой бетонной зоны до нейтральной линии, мм.
    EffectiveZoneDepth As Double ' Глубина зоны Abt, реально использованная после ограничений, мм.
    CentralTensionBranch As Boolean ' True, если сработала отдельная ветка центрального растяжения.
    Formation As CCrackFormationResult ' Единственный источник стадии образования трещины.
    FallbackPsi1 As Boolean ' Нет пригодной точки образования/состояния Post; ширина считается консервативно по текущему НДС.
    PsiFallbackReason As String ' Предупреждение только о коэффициенте, действительно использованном в формуле.
    PostHasSelectedTension As Boolean ' В Post есть положительное напряжение хотя бы одного выбранного текущим НДС стержня.
    GeometryQuery As CSectionGeometryQuery ' Один общий geometry session на подготовку ширины.
    InteractionRegion As CConcreteRegion ' Та же конечная область, площадь которой принята в формуле Abt.
    StandardCode As String ' SP63 или SP35: только выбор формулы ширины, не модели НДС.
    SP35GroupGapTolerance As Double ' Чистый зазор объединения реальных стержней, мм.
    SP35RowTolerance As Double ' Максимальная ширина ряда по нормали к Н.О., мм.
    SP35NeighborRatioLimit As Double ' Безразмерный порог t/L, не допуск группировки рядов.
    SP35RadiusDiameterMode As String ' Max/Min/Average только для геометрических радиусов.
    SP35InteractionRadiusMode As String ' 3d/5d/6d: нормативный множитель диаметра.
    SP35RebarProfile As String ' Periodic/Smooth выбирает формулу psi СП 35.
    SP35Data As CSP35CrackData ' Готовые данные всех проверяемых областей взаимодействия СП 35 с запретом изменений.
    SP35CriticalCandidate As Long ' Номер проверяемой области с наибольшей шириной трещины; наибольшая площадь не является критерием выбора.
    SP35Radius As Double ' Радиус армирования Rr, мм, и размерный параметр psi СП 35, см.
    SP35PsiCm As Double ' Радиус армирования Rr, мм, и размерный параметр psi СП 35, см.
End Type

Public Const CRACK_PHI1_LONG_TERM As Double = 1.4 ' Коэффициент длительного действия нагрузки для a_crc1 по СП 63.
Public Const CRACK_PHI2_DEFAULT As Double = 0.5 ' Типовой default для обычной арматуры периодического профиля.
Public Const CRACK_PHI3_TENSION As Double = 1.2 ' Коэффициент для центрального/сквозного растяжения.
Public Const CRACK_PHI3_BENDING As Double = 1# ' Коэффициент для изгиба и внецентренного сжатия.
Public Const CRACK_NEAR_ZERO_STRAIN As Double = 0.000000000001 ' Численный ноль для деформаций и кривизн.
Public Const CRACK_NEAR_ZERO_FORCE As Double = 0.000001 ' Численный ноль для отличия центрального растяжения от нулевой нагрузки.
Public Const CRACK_NEAR_ZERO_STRESS As Double = 0.000000001 ' Численный ноль для справочных напряжений, МПа.
Public Const CRACK_COVER_MODE_NEAREST_CONTOUR As String = "NearestContour" ' a_s до ближайшего выхода бетонного контура по нормали.
Public Const CRACK_COVER_MODE_GLOBAL_EXTREME As String = "GlobalExtreme" ' a_s до дальней растянутой опорной линии всего сечения.
Public Const CRACK_PSI_MODE_USER As String = "User" ' psi_s задает пользователь через SLS.Crack.PsiS.
Public Const CRACK_PSI_MODE_AUTO As String = "Auto" ' psi_s уточняется только если первая проверка с psi_s = 1 не прошла.
Public Const CRACK_PSI_MODE_ALWAYS_CALC As String = "AlwaysCalc" ' psi_s всегда считается по sigma_s,crc, если такая точка доступна.
Public Const CRACK_SIGMA_S_CRC_AVERAGE_ALL_SELECTED As String = "AllSelected" ' Для sigma_s,crc учитывать все выбранные стержни со знаком напряжения.
Public Const CRACK_SIGMA_S_CRC_AVERAGE_TENSION_ONLY As String = "TensionOnly" ' Сжатые напряжения в sigma_s,crc принимать равными нулю.

' Вычисляет деформацию в центре бетонной сетки для выбора почти равномерной
' ветки; центр уже кэширован подготовкой НДС, новый обход элементов не нужен.
Public Function CrackStrainAtConcreteCenter(ByVal state As CSectionStateResult, ByVal section As CSectionModel) As Double
    Dim centerX As Double, centerY As Double, concreteArea As Double
    section.ConcreteMeshCentroid centerX, centerY, concreteArea
    CrackStrainAtConcreteCenter = state.Epsilon0 + state.KappaX * centerY + state.KappaY * centerX
End Function

' Возвращает численное отношение готового раскрытия к допускаемому.
' При нулевом пределе отношение не определено; решение о применимости
' проверки принимается до вызова этой чистой операции.
Public Function CrackUtilizationFromData(ByVal crackWidthValue As Double, ByVal allowableCrackWidth As Double) As Double
    If allowableCrackWidth <= 0# Then Exit Function
    CrackUtilizationFromData = crackWidthValue / allowableCrackWidth
End Function

' Записывает машинный результат формульной части ширины раскрытия.
Public Sub CrackSetWidthResult(ByRef work As TCrackWidthWork, ByVal internalStatus As EResultInternalStatus, _
        ByVal resultCode As EResultCode, ByVal reason As String)
    work.WidthInternalStatus = internalStatus
    work.WidthResultCode = resultCode
    work.StopReason = CrackUserStopReason(reason)
End Sub

' Готовит короткий пользовательский комментарий из внутренней причины.
' Машинная классификация уже сохранена в InternalStatus/ResultCode, поэтому
' технические префиксы NumericalFailure/InputError на лист выводить не нужно.
Public Function CrackUserStopReason(ByVal reason As String) As String
    CrackUserStopReason = Trim$(reason)
    If Len(CrackUserStopReason) = 0 Then Exit Function

    CrackUserStopReason = Replace$(CrackUserStopReason, "NumericalFailure:", vbNullString, 1, 1, vbTextCompare)
    CrackUserStopReason = Replace$(CrackUserStopReason, "InputError:", vbNullString, 1, 1, vbTextCompare)
    CrackUserStopReason = Replace$(CrackUserStopReason, "lambda_crc", "lambda трещинообразования", 1, -1, vbTextCompare)
    CrackUserStopReason = Replace$(CrackUserStopReason, "eps_bt,ult", "предельная растягивающая деформация бетона", 1, -1, vbTextCompare)
    CrackUserStopReason = Trim$(CrackUserStopReason)
    Do While InStr(1, CrackUserStopReason, "..", vbBinaryCompare) > 0
        CrackUserStopReason = Replace$(CrackUserStopReason, "..", ".")
    Loop
End Function

' Помечает ошибку входных данных, найденную уже при подготовке данных формулы.
Public Sub CrackSetWidthInvalidInput(ByRef work As TCrackWidthWork, ByVal reason As String)
    work.Converged = False
    work.CrackFormed = False
    CrackSetWidthResult work, rsInvalidInput, rcInvalidInput, reason
End Sub

' Помечает расчет как корректно завершенный без раскрытой трещины.
' Это не численная ошибка: заданное состояние либо не растягивает арматуру,
' либо еще не достигло момента образования трещины.
Public Sub CrackFinishNoCrack(ByRef work As TCrackWidthWork, ByVal reason As String)
    work.Converged = True
    work.CrackFormed = False
    work.CrackWidth = 0#
    work.Utilization = 0#
    work.StopReason = reason
    If work.WidthInternalStatus = rsUnset Then _
        CrackSetWidthResult work, rsNotApplicable, rcCrackNotFormed, reason
    CrackAppendDiagnostic work, reason
End Sub

' Добавляет строку в диагностический журнал.
Public Sub CrackAppendDiagnostic(ByRef work As TCrackWidthWork, ByVal text As String)
    If Not work.DiagnosticsEnabled Then Exit Sub
    work.DiagnosticLog = work.DiagnosticLog & text & vbCrLf
End Sub

' Форматирует числа журнала и инженерного комментария через точку.
' У целого числа убирает пустой десятичный разделитель перед единицей измерения.
Public Function CrackFormatNumberInvariant(ByVal value As Double) As String
    CrackFormatNumberInvariant = Replace$(Format$(value, "0.############"), ",", ".")
    If Right$(CrackFormatNumberInvariant, 1) = "." Then _
        CrackFormatNumberInvariant = Left$(CrackFormatNumberInvariant, Len(CrackFormatNumberInvariant) - 1)
End Function

' Утилитарный минимум без WorksheetFunction в расчетном ядре трещин.
Public Function CrackMinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then CrackMinDouble = a Else CrackMinDouble = b
End Function

' Утилитарный максимум без WorksheetFunction в расчетном ядре трещин.
Public Function CrackMaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then CrackMaxDouble = a Else CrackMaxDouble = b
End Function
