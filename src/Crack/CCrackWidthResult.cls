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
' Результат расчета ширины раскрытия нормальной трещины
' ==========================================================================
' CCrackWidthResult хранит итог формульного этапа a_crc и все подготовленные
' инженерные величины, по которым пользователь может проверить расчет вручную.
' Если обязательные данные отсутствуют, orchestration-слой формирует ResultMeta
' rsBlockedByDependency или rsNotApplicable, а чистая формула при этом не
' вызывается. Точка Formation не является обязательной зависимостью:
' при ее отсутствии calculator может получить ширину по текущему НДС с psi_s = 1.

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mMeta As CResultMeta ' Типизированный статус именно формульного этапа width.
Private mCrackWidth As Double ' a_crc, мм.
Private mAllowableCrackWidth As Double ' a_crc,ult, мм.
Private mUtilization As Double ' a_crc / a_crc,ult.
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

' Создает нейтральный not-applicable результат до заполнения.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkCrackWidth
End Sub

' Копирует все подготовленные crack-width величины из calculator-а.
' Downstream-слой работает с отдельным result-объектом и ResultMeta, не
' обращаясь к внутренним временным полям calculator-а.
Public Sub InitializeFromCalculator(ByVal crack As CCrackWidthCalculator, ByVal meta As CResultMeta)
    AssertWritable
    ResetData
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackWidth, _
            "Meta результата ширины трещины не сформирована.", vbNullString, True, False
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
End Sub

' Возвращает копию meta формульного расчета ширины или его блокировки.
' Чтение результата не пересчитывает ширину и не запускает поиск НДС.
Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
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
    AsTension = mAsTension
End Property

Public Property Get Abt() As Double
    Abt = mAbt
End Property

Public Property Get DsEquivalent() As Double
    DsEquivalent = mDsEquivalent
End Property

Public Property Get CrackSpacingRaw() As Double
    CrackSpacingRaw = mCrackSpacingRaw
End Property

Public Property Get CrackSpacing() As Double
    CrackSpacing = mCrackSpacing
End Property

Public Property Get TensionRebarCount() As Long
    TensionRebarCount = mTensionRebarCount
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

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Возвращает запас только завершенной проверки с ненулевой шириной.
' Неприменимый или заблокированный расчет не получает вымышленный запас.
Public Property Get ReserveFactor() As Double
    If mUtilization <= 0# Then Exit Property
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsFinished(policy.ExternalStatus(mMeta)) Then Exit Property
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
End Sub

' Завершает жизненный цикл заполнения перед передачей snapshot потребителю.
' Повторный вызов безопасен. Запрет распространяется на численные параметры
' раскрытия и meta; State-объекты этот формульный результат не хранит.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCrackWidthResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
