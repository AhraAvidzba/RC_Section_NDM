VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackWidthResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат расчёта ширины раскрытия нормальной трещины
' ==========================================================================
' Хранит a_crc и исходные инженерные величины для ручной проверки расчёта.
' Отсутствие обязательных данных отражается в ResultMeta как недоступность
' или неприменимость проверки. Если нагрузка образования трещины не найдена,
' ширина может быть рассчитана по пригодному текущему НДС; в СП 63 недоступное
' уточнение psi_s заменяется значением 1. В формуле СП 35 этот коэффициент не применяется.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mMeta As CResultMeta ' Типизированный статус именно формульного этапа width.
Private mCrackWidth As Double ' a_crc, мм.
Private mAllowableCrackWidth As Double ' a_crc,ult, мм.
Private mUtilization As Double ' Отношение расчётной ширины трещины к допустимой: a_crc / a_crc,ult.
Private mPhi1 As Double ' Коэффициент длительности нагрузки.
Private mPhi2 As Double ' Коэффициент поверхности арматуры.
Private mPhi3 As Double ' Коэффициент характера работы элемента.
Private mPsiS As Double ' Коэффициент неравномерности деформаций арматуры.
Private mSigmaS As Double ' Среднее напряжение выбранной растянутой арматуры.
Private mSigmaSCrc As Double ' Напряжение той же группы стержней при образовании трещины.
Private mSteelEs As Double ' Модуль упругости арматуры для ручной проверки a_crc.
Private mAsTension As Double ' Площадь выбранной растянутой арматуры.
Private mAbt As Double ' Площадь растянутого бетона принятой зоны взаимодействия.
Private mDsEquivalent As Double ' Эквивалентный диаметр выбранной арматуры.
Private mCrackSpacingRaw As Double ' l_s0 до нормативных ограничений.
Private mCrackSpacing As Double ' Принятое l_s после ограничений.
Private mTensionRebarCount As Long ' Число выбранных стержней.
Private mTensionRebarIds As String ' Имена выбранных стержней.
Private mSectionDepthH As Double ' Полная высота по нормали к нейтральной линии.
Private mCoverA As Double ' Расстояние от растянутой поверхности до крайней арматуры.
Private mTensionDepth As Double ' Фактическая высота растянутой зоны.
Private mEffectiveZoneDepth As Double ' Принятая высота зоны взаимодействия.
Private mCentralTensionBranch As Boolean ' True для ветки центрального растяжения.
Private mStopReason As String ' Пользовательская причина остановки/итога.
Private mDiagnosticLog As String ' Техническая диагностика подготовки данных и формулы.
Private mInteractionRegion As CConcreteRegion ' Аналитический снимок той же области, площадь которой вошла в Abt/Ar.
Private mStandardCode As String ' Методика действительно выполненной формулы ширины.
Private mSP35Data As CSP35CrackData ' Готовые данные областей взаимодействия и реальных стержней всех рядов с запретом изменений.
Private mSP35CriticalCandidate As Long ' Номер области взаимодействия с наибольшей шириной трещины.
Private mSP35Radius As Double, mSP35PsiCm As Double ' Радиус армирования Rr, мм, и размерный параметр psi СП 35, см.

' Создает нейтральный not-applicable результат до заполнения.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkCrackWidth
End Sub

' Сохраняет подготовленные величины расчёта ширины вместе со статусом и пояснением.
' При выводе используются эти данные, без обращения к временным полям расчётного класса.
Public Sub InitializeFromCalculator(ByVal crack As CCrackWidthCalculator, ByVal meta As CResultMeta, _
        Optional ByVal combinationID As String = vbNullString)
    AssertWritable
    ResetData
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackWidth, _
            "Не сформированы статус и причина результата расчёта ширины трещины.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If

    If crack Is Nothing Then Exit Sub
    mCrackWidth = crack.CrackWidth
    mAllowableCrackWidth = crack.AllowableCrackWidth
    mUtilization = crack.Utilization
    mPhi1 = crack.Phi1
    mPhi2 = crack.Phi2
    mPhi3 = crack.Phi3
    mPsiS = crack.PsiS
    mSigmaS = crack.SigmaS
    mSigmaSCrc = crack.SigmaSCrc
    mSteelEs = crack.SteelEs
    mAsTension = crack.AsTension
    mAbt = crack.Abt
    mDsEquivalent = crack.DsEquivalent
    mCrackSpacingRaw = crack.CrackSpacingRaw
    mCrackSpacing = crack.CrackSpacing
    mTensionRebarCount = crack.TensionRebarCount
    mTensionRebarIds = crack.TensionRebarIds
    mSectionDepthH = crack.SectionDepthH
    mCoverA = crack.CoverA
    mTensionDepth = crack.TensionDepth
    mEffectiveZoneDepth = crack.EffectiveZoneDepth
    mCentralTensionBranch = crack.CentralTensionBranch
    mStopReason = crack.StopReason
    mDiagnosticLog = crack.DiagnosticLog
    Set mInteractionRegion = crack.InteractionRegion
    mStandardCode = crack.StandardCode
    Set mSP35Data = crack.SP35Data
    mSP35CriticalCandidate = crack.SP35CriticalCandidate
    mSP35Radius = crack.SP35ReinforcementRadius: mSP35PsiCm = crack.SP35PsiCm
    Dim anchorID As String
    If mSP35CriticalCandidate > 0 Then anchorID = "G" & CStr(mSP35Data.CandidateAnchorGroup(mSP35CriticalCandidate))
    If Not mInteractionRegion Is Nothing Then mInteractionRegion.SetOwner combinationID, mStandardCode, anchorID
End Sub

' Возвращает копию статуса и пояснения формулы ширины или причины её недоступности.
' Чтение не пересчитывает ширину и не запускает поиск НДС.
Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' a_crc, мм.
Public Property Get CrackWidth() As Double
    CrackWidth = mCrackWidth
End Property

' a_crc,ult, мм.
Public Property Get AllowableCrackWidth() As Double
    AllowableCrackWidth = mAllowableCrackWidth
End Property

' Сохранённый коэффициент использования проверяемого предела; смысл предела задаёт результат
' этого класса.
Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

' Коэффициент длительности нагрузки.
Public Property Get Phi1() As Double
    Phi1 = mPhi1
End Property

' Коэффициент поверхности арматуры.
Public Property Get Phi2() As Double
    Phi2 = mPhi2
End Property

' Коэффициент характера работы элемента.
Public Property Get Phi3() As Double
    Phi3 = mPhi3
End Property

' Коэффициент неравномерности деформаций арматуры.
Public Property Get PsiS() As Double
    PsiS = mPsiS
End Property

' Среднее напряжение выбранной растянутой арматуры.
Public Property Get SigmaS() As Double
    SigmaS = mSigmaS
End Property

' Напряжение той же группы стержней при образовании трещины.
Public Property Get SigmaSCrc() As Double
    SigmaSCrc = mSigmaSCrc
End Property

' Модуль упругости арматуры для ручной проверки a_crc.
Public Property Get SteelEs() As Double
    SteelEs = mSteelEs
End Property

' Площадь выбранной растянутой арматуры.
Public Property Get AsTension() As Double
    AsTension = mAsTension
End Property

' Площадь растянутого бетона принятой зоны взаимодействия.
Public Property Get Abt() As Double
    Abt = mAbt
End Property

' Эквивалентный диаметр выбранной арматуры.
Public Property Get DsEquivalent() As Double
    DsEquivalent = mDsEquivalent
End Property

' l_s0 до нормативных ограничений.
Public Property Get CrackSpacingRaw() As Double
    CrackSpacingRaw = mCrackSpacingRaw
End Property

' Принятое l_s после ограничений.
Public Property Get CrackSpacing() As Double
    CrackSpacing = mCrackSpacing
End Property

' Число выбранных стержней.
Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mTensionRebarCount
End Property

' Имена выбранных стержней.
Public Property Get TensionRebarIds() As String
    TensionRebarIds = mTensionRebarIds
End Property

' Полная высота по нормали к нейтральной линии.
Public Property Get SectionDepthH() As Double
    SectionDepthH = mSectionDepthH
End Property

' Расстояние от растянутой поверхности до крайней арматуры.
Public Property Get CoverA() As Double
    CoverA = mCoverA
End Property

' Фактическая высота растянутой зоны.
Public Property Get TensionDepth() As Double
    TensionDepth = mTensionDepth
End Property

' Принятая высота зоны взаимодействия.
Public Property Get EffectiveZoneDepth() As Double
    EffectiveZoneDepth = mEffectiveZoneDepth
End Property

' True для ветки центрального растяжения.
Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property

' Пользовательская причина остановки/итога.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Техническая диагностика подготовки данных и формулы.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Возвращает независимую копию точных замкнутых контуров расчётной области.
' Запись на лист и экспорт используют её напрямую, без восстановления по числу Abt
' и без новой обрезки геометрии.
Public Property Get InteractionRegion() As CConcreteRegion
    If Not mInteractionRegion Is Nothing Then Set InteractionRegion = mInteractionRegion.Clone
End Property

' Возвращают собственный снимок методики и критического расчета без поиска
' НДС или геометрии. SP35Data уже заморожен владельцем подготовки и формулы.
Public Property Get StandardCode() As String
    StandardCode = mStandardCode
End Property
' Возвращает готовые данные областей взаимодействия и реальных стержней всех рядов.
' Изменять эти данные после расчёта нельзя.
Public Property Get SP35Data() As CSP35CrackData
    Set SP35Data = mSP35Data
End Property
' Возвращает номер области взаимодействия с наибольшей шириной трещины.
Public Property Get SP35CriticalCandidate() As Long
    SP35CriticalCandidate = mSP35CriticalCandidate
End Property
' Радиус армирования в мм и размерный psi в см.
Public Property Get SP35ReinforcementRadius() As Double
    SP35ReinforcementRadius = mSP35Radius
End Property
' Рассчитанный коэффициент psi_cm выбранной области СП35.
Public Property Get SP35PsiCm() As Double
    SP35PsiCm = mSP35PsiCm
End Property
' Сохраняет номер области взаимодействия с наибольшей шириной трещины.
Public Property Get CriticalAnchorGroupID() As String
    If mSP35CriticalCandidate > 0 Then CriticalAnchorGroupID = "G" & CStr(mSP35Data.CandidateAnchorGroup(mSP35CriticalCandidate))
End Property

' Возвращает запас только завершенной проверки с ненулевой шириной.
' Неприменимый или заблокированный расчет не получает вымышленный запас.
Public Property Get ReserveFactor() As Double
    If mUtilization <= 0# Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsMetaFinished(mMeta) Then Exit Property
    ReserveFactor = 1# / mUtilization
End Property

' Сбрасывает численный снимок, чтобы повторная инициализация без calculator-а
' не оставляла ширину, коэффициенты и выбранные стержни прошлого сочетания.
Private Sub ResetData()
    mCrackWidth = 0#
    mAllowableCrackWidth = 0#
    mUtilization = 0#
    mPhi1 = 0#
    mPhi2 = 0#
    mPhi3 = 0#
    mPsiS = 0#
    mSigmaS = 0#
    mSigmaSCrc = 0#
    mSteelEs = 0#
    mAsTension = 0#
    mAbt = 0#
    mDsEquivalent = 0#
    mCrackSpacingRaw = 0#
    mCrackSpacing = 0#
    mTensionRebarCount = 0
    mTensionRebarIds = vbNullString
    mSectionDepthH = 0#
    mCoverA = 0#
    mTensionDepth = 0#
    mEffectiveZoneDepth = 0#
    mCentralTensionBranch = False
    mStopReason = vbNullString
    mDiagnosticLog = vbNullString
    Set mInteractionRegion = Nothing
    mStandardCode = vbNullString: Set mSP35Data = Nothing
    mSP35CriticalCandidate = 0: mSP35Radius = 0#: mSP35PsiCm = 0#
End Sub

' Завершает заполнение результата ширины, запрещая изменение чисел, статуса и пояснения.
' Повторный вызов безопасен. Сам этот результат НДС не хранит.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся отдельный объект; сохранённый результат не изменяется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCrackWidthResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
