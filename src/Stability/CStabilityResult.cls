VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStabilityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат проверки устойчивости одного сочетания
' ==========================================================================
' CStabilityResult хранит typed meta и инженерные величины, которые CStabilityCalculator
' уже посчитал для LC. Класс не запускает расчет устойчивости; его задача - быть
' stability-subtree внутри CCombinationResult и единым источником данных для Results.

Private mFrozen As Boolean ' После публикации заполнение/reset запрещены; численные данные не копируются.
Private mMeta As CResultMeta ' Единственная typed meta ветви устойчивости.
Private mCode As String ' Нормативная ветка SP35/SP63.
Private mBranch As String ' Диагностическое имя расчетной ветви.
Private mReferenceX As Double ' X точки, относительно которой выполнен перенос нагрузок.
Private mReferenceY As Double ' Y точки, относительно которой выполнен перенос нагрузок.
Private mDesignN As Double ' Итоговая продольная сила для последующего НДМ.
Private mDesignMx As Double ' Итоговый Mx после учета устойчивости/фильтров.
Private mDesignMy As Double ' Итоговый My после учета устойчивости/фильтров.
Private mSustainedN As Double ' Длительная/постоянная часть N.
Private mMoment1 As Double
Private mMoment2 As Double
Private mSustainedMoment1 As Double
Private mSustainedMoment2 As Double
Private mDesignMoment1 As Double
Private mDesignMoment2 As Double
Private mCriticalForce As Double
Private mReserveFactor As Double
Private mUtilization As Double
Private mNcr1 As Double
Private mNcr2 As Double
Private mNultimate1 As Double
Private mNultimate2 As Double
Private mReserve1 As Double
Private mReserve2 As Double
Private mEta1 As Double
Private mEta2 As Double
Private mPhiL1 As Double
Private mPhiL2 As Double
Private mPhiLTable1 As String
Private mPhiLTable2 As String
Private mPhiM1 As Double
Private mPhiM2 As Double
Private mPhiValue1 As Double
Private mPhiValue2 As Double
Private mPhiP1 As Double
Private mPhiP2 As Double
Private mEccentricity1 As Double
Private mEccentricity2 As Double
Private mAccidentalEcc1 As Double
Private mAccidentalEcc2 As Double
Private mEffectiveLength1 As Double
Private mEffectiveLength2 As Double
Private mSlenderness1 As Double
Private mSlenderness2 As Double
Private mDepth1 As Double
Private mDepth2 As Double
Private mRadius1 As Double
Private mRadius2 As Double
Private mCoreDistance1 As Double
Private mCoreDistance2 As Double
Private mDelta1 As Double
Private mDelta2 As Double
Private mKb1 As Double
Private mKb2 As Double
Private mKs1 As Double
Private mKs2 As Double
Private mStiffnessD1 As Double
Private mStiffnessD2 As Double
Private mPlaneBranch1 As String
Private mPlaneBranch2 As String
Private mPlaneApplicable1 As Boolean
Private mPlaneApplicable2 As Boolean
Private mPlanePassed1 As Boolean
Private mPlanePassed2 As Boolean

' Создает N/A-result для случаев, когда устойчивость не запрошена профилем.
Private Sub Class_Initialize()
    Clear
End Sub

' Полностью сбрасывает meta и численные поля перед повторным заполнением.
Public Sub Clear()
    AssertWritable
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkStability
    mCode = vbNullString
    mBranch = vbNullString
    mReferenceX = 0#
    mReferenceY = 0#
    mDesignN = 0#
    mDesignMx = 0#
    mDesignMy = 0#
    mSustainedN = 0#
    mMoment1 = 0#
    mMoment2 = 0#
    mSustainedMoment1 = 0#
    mSustainedMoment2 = 0#
    mDesignMoment1 = 0#
    mDesignMoment2 = 0#
    mCriticalForce = 0#
    mReserveFactor = 0#
    mUtilization = 0#
    mNcr1 = 0#
    mNcr2 = 0#
    mNultimate1 = 0#
    mNultimate2 = 0#
    mReserve1 = 0#
    mReserve2 = 0#
    mEta1 = 0#
    mEta2 = 0#
    mPhiL1 = 0#
    mPhiL2 = 0#
    mPhiLTable1 = vbNullString
    mPhiLTable2 = vbNullString
    mPhiM1 = 0#
    mPhiM2 = 0#
    mPhiValue1 = 0#
    mPhiValue2 = 0#
    mPhiP1 = 0#
    mPhiP2 = 0#
    mEccentricity1 = 0#
    mEccentricity2 = 0#
    mAccidentalEcc1 = 0#
    mAccidentalEcc2 = 0#
    mEffectiveLength1 = 0#
    mEffectiveLength2 = 0#
    mSlenderness1 = 0#
    mSlenderness2 = 0#
    mDepth1 = 0#
    mDepth2 = 0#
    mRadius1 = 0#
    mRadius2 = 0#
    mCoreDistance1 = 0#
    mCoreDistance2 = 0#
    mDelta1 = 0#
    mDelta2 = 0#
    mKb1 = 0#
    mKb2 = 0#
    mKs1 = 0#
    mKs2 = 0#
    mStiffnessD1 = 0#
    mStiffnessD2 = 0#
    mPlaneBranch1 = vbNullString
    mPlaneBranch2 = vbNullString
    mPlaneApplicable1 = False
    mPlaneApplicable2 = False
    mPlanePassed1 = False
    mPlanePassed2 = False
End Sub

' Сохраняет инженерную meta без отдельного изменяемого display-статуса.
Public Sub SetMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    AssertWritable
    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
        If Len(Trim$(resultComment)) > 0 Then mMeta.ResultComment = StabilityBranchDisplayName(resultComment)
    End If
    mBranch = StabilityBranchDisplayName(resultComment)
End Sub

' Копирует результат CStabilityCalculator и формирует typed meta для summary.
Public Sub InitializeFromCalculator(ByVal meta As CResultMeta, ByVal stability As CStabilityCalculator, _
        ByVal referenceXValue As Double, ByVal referenceYValue As Double)
    AssertWritable
    Clear
    mReferenceX = referenceXValue
    mReferenceY = referenceYValue

    If meta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If

    If stability Is Nothing Then
        Exit Sub
    End If

    mCode = stability.Code
    ' mBranch остается машинным ключом расчетной ветви. Причина отказа
    ' принадлежит ResultComment, а не заменяет ключ для typed-потребителей.
    mBranch = stability.Branch
    mDesignN = stability.DesignN
    mDesignMx = stability.DesignMx
    mDesignMy = stability.DesignMy
    mSustainedN = stability.SustainedN
    mMoment1 = stability.Moment1
    mMoment2 = stability.Moment2
    mSustainedMoment1 = stability.SustainedMoment1
    mSustainedMoment2 = stability.SustainedMoment2
    mDesignMoment1 = stability.DesignMoment1
    mDesignMoment2 = stability.DesignMoment2
    mCriticalForce = stability.CriticalForce
    mReserveFactor = stability.ReserveFactor
    If mReserveFactor > 0# Then mUtilization = 1# / mReserveFactor
    mNcr1 = stability.Ncr1
    mNcr2 = stability.Ncr2
    mNultimate1 = stability.Nultimate1
    mNultimate2 = stability.Nultimate2
    mReserve1 = stability.Reserve1
    mReserve2 = stability.Reserve2
    mEta1 = stability.Eta1
    mEta2 = stability.Eta2
    mPhiL1 = stability.PhiL1
    mPhiL2 = stability.PhiL2
    mPhiLTable1 = stability.PhiLTable1
    mPhiLTable2 = stability.PhiLTable2
    mPhiM1 = stability.PhiM1
    mPhiM2 = stability.PhiM2
    mPhiValue1 = stability.PhiValue1
    mPhiValue2 = stability.PhiValue2
    mPhiP1 = stability.PhiP1
    mPhiP2 = stability.PhiP2
    mEccentricity1 = stability.Eccentricity1
    mEccentricity2 = stability.Eccentricity2
    mAccidentalEcc1 = stability.AccidentalEccentricity1
    mAccidentalEcc2 = stability.AccidentalEccentricity2
    mEffectiveLength1 = stability.EffectiveLength1
    mEffectiveLength2 = stability.EffectiveLength2
    mSlenderness1 = stability.Slenderness1
    mSlenderness2 = stability.Slenderness2
    mDepth1 = stability.Depth1
    mDepth2 = stability.Depth2
    mRadius1 = stability.Radius1
    mRadius2 = stability.Radius2
    mCoreDistance1 = stability.CoreDistance1
    mCoreDistance2 = stability.CoreDistance2
    mDelta1 = stability.Delta1
    mDelta2 = stability.Delta2
    mKb1 = stability.Kb1
    mKb2 = stability.Kb2
    mKs1 = stability.Ks1
    mKs2 = stability.Ks2
    mStiffnessD1 = stability.StiffnessD1
    mStiffnessD2 = stability.StiffnessD2
    mPlaneBranch1 = stability.PlaneBranch1
    mPlaneBranch2 = stability.PlaneBranch2
    mPlaneApplicable1 = stability.PlaneApplicable1
    mPlaneApplicable2 = stability.PlaneApplicable2
    mPlanePassed1 = stability.PlanePassed1
    mPlanePassed2 = stability.PlanePassed2
End Sub

' Переводит внутренние коды ветвей устойчивости в русский текст для Results.
' Сами коды SP35-eta/SP35-table/SP35-mixed остаются во внутренних полях batch.
Private Function StabilityBranchDisplayName(ByVal branchText As String) As String
    Select Case LCase$(Trim$(branchText))
        Case "nocompression"
            StabilityBranchDisplayName = "устойчивость не требуется: нет сжатия"
        Case "sp63"
            StabilityBranchDisplayName = "СП 63: расчет eta и Ncr"
        Case "sp35-eta"
            StabilityBranchDisplayName = "СП 35: ветвь ec > r через eta и Ncr"
        Case "sp35-table"
            StabilityBranchDisplayName = "СП 35: ветвь ec <= r через Nult,stab"
        Case "sp35-mixed"
            StabilityBranchDisplayName = "СП 35: плоскости попали в разные расчетные ветви"
        Case Else
            StabilityBranchDisplayName = branchText
    End Select
End Function

' Обновляет финальный расчетный вектор после общего zero-filter batch-слоя.
Public Sub UpdateDesignLoads(ByVal designNValue As Double, ByVal designMxValue As Double, _
        ByVal designMyValue As Double, ByVal principalAngleRad As Double)
    AssertWritable
    mDesignN = designNValue
    mDesignMx = designMxValue
    mDesignMy = designMyValue
    mDesignMoment1 = designMxValue * Cos(principalAngleRad) + designMyValue * Sin(principalAngleRad)
    mDesignMoment2 = -designMxValue * Sin(principalAngleRad) + designMyValue * Cos(principalAngleRad)
End Sub

' Возвращает копию meta, чтобы вызывающий код не менял status-subtree напрямую.
Public Property Get Meta() As CResultMeta
    If mMeta Is Nothing Then
        Set mMeta = New CResultMeta
        mMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован.", vbNullString, True, False
    End If
    Set Meta = mMeta.Clone
End Property

' Display-статус всегда получается из meta, без независимого mutable поля.
Public Property Get Status() As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    Status = policy.ExternalStatus(mMeta)
End Property

' Выдает минимальный доступный запас главных плоскостей для сводки и Worst.
' Числа уже учитывают примененный критерий устойчивости; здесь не пересчитываются СП.
Public Property Get SummaryReserve() As Double
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If Not policy.IsFinished(Status) Then Exit Property
    If mPlaneApplicable1 And mReserve1 > 0# Then SummaryReserve = mReserve1
    If mPlaneApplicable2 And mReserve2 > 0# Then
        If SummaryReserve <= 0# Or mReserve2 < SummaryReserve Then SummaryReserve = mReserve2
    End If
    If SummaryReserve <= 0# And mReserveFactor > 0# Then SummaryReserve = mReserveFactor
End Property

' Численные значения являются сохраненными данными calculator-а, а не
' изменяемыми входами. Все getter-ы этой группы доступны только для чтения.
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Get Branch() As String
    Branch = mBranch
End Property

Public Property Get ReferenceX() As Double
    ReferenceX = mReferenceX
End Property

Public Property Get ReferenceY() As Double
    ReferenceY = mReferenceY
End Property

Public Property Get DesignN() As Double
    DesignN = mDesignN
End Property

Public Property Get DesignMx() As Double
    DesignMx = mDesignMx
End Property

Public Property Get DesignMy() As Double
    DesignMy = mDesignMy
End Property

Public Property Get SustainedN() As Double
    SustainedN = mSustainedN
End Property

Public Property Get Moment1() As Double
    Moment1 = mMoment1
End Property

Public Property Get Moment2() As Double
    Moment2 = mMoment2
End Property

Public Property Get SustainedMoment1() As Double
    SustainedMoment1 = mSustainedMoment1
End Property

Public Property Get SustainedMoment2() As Double
    SustainedMoment2 = mSustainedMoment2
End Property

Public Property Get DesignMoment1() As Double
    DesignMoment1 = mDesignMoment1
End Property

Public Property Get DesignMoment2() As Double
    DesignMoment2 = mDesignMoment2
End Property

Public Property Get CriticalForce() As Double
    CriticalForce = mCriticalForce
End Property

Public Property Get ReserveFactor() As Double
    ReserveFactor = mReserveFactor
End Property

Public Property Get Utilization() As Double
    Utilization = mUtilization
End Property

Public Property Get Ncr1() As Double
    Ncr1 = mNcr1
End Property

Public Property Get Ncr2() As Double
    Ncr2 = mNcr2
End Property

Public Property Get Nultimate1() As Double
    Nultimate1 = mNultimate1
End Property

Public Property Get Nultimate2() As Double
    Nultimate2 = mNultimate2
End Property

Public Property Get Reserve1() As Double
    Reserve1 = mReserve1
End Property

Public Property Get Reserve2() As Double
    Reserve2 = mReserve2
End Property

Public Property Get Eta1() As Double
    Eta1 = mEta1
End Property

Public Property Get Eta2() As Double
    Eta2 = mEta2
End Property

Public Property Get PhiL1() As Double
    PhiL1 = mPhiL1
End Property

Public Property Get PhiL2() As Double
    PhiL2 = mPhiL2
End Property

Public Property Get PhiLTable1() As String
    PhiLTable1 = mPhiLTable1
End Property

Public Property Get PhiLTable2() As String
    PhiLTable2 = mPhiLTable2
End Property

Public Property Get PhiM1() As Double
    PhiM1 = mPhiM1
End Property

Public Property Get PhiM2() As Double
    PhiM2 = mPhiM2
End Property

Public Property Get PhiValue1() As Double
    PhiValue1 = mPhiValue1
End Property

Public Property Get PhiValue2() As Double
    PhiValue2 = mPhiValue2
End Property

Public Property Get PhiP1() As Double
    PhiP1 = mPhiP1
End Property

Public Property Get PhiP2() As Double
    PhiP2 = mPhiP2
End Property

Public Property Get Eccentricity1() As Double
    Eccentricity1 = mEccentricity1
End Property

Public Property Get Eccentricity2() As Double
    Eccentricity2 = mEccentricity2
End Property

Public Property Get AccidentalEcc1() As Double
    AccidentalEcc1 = mAccidentalEcc1
End Property

Public Property Get AccidentalEcc2() As Double
    AccidentalEcc2 = mAccidentalEcc2
End Property

Public Property Get EffectiveLength1() As Double
    EffectiveLength1 = mEffectiveLength1
End Property

Public Property Get EffectiveLength2() As Double
    EffectiveLength2 = mEffectiveLength2
End Property

Public Property Get Slenderness1() As Double
    Slenderness1 = mSlenderness1
End Property

Public Property Get Slenderness2() As Double
    Slenderness2 = mSlenderness2
End Property

Public Property Get Depth1() As Double
    Depth1 = mDepth1
End Property

Public Property Get Depth2() As Double
    Depth2 = mDepth2
End Property

Public Property Get Radius1() As Double
    Radius1 = mRadius1
End Property

Public Property Get Radius2() As Double
    Radius2 = mRadius2
End Property

Public Property Get CoreDistance1() As Double
    CoreDistance1 = mCoreDistance1
End Property

Public Property Get CoreDistance2() As Double
    CoreDistance2 = mCoreDistance2
End Property

Public Property Get Delta1() As Double
    Delta1 = mDelta1
End Property

Public Property Get Delta2() As Double
    Delta2 = mDelta2
End Property

Public Property Get Kb1() As Double
    Kb1 = mKb1
End Property

Public Property Get Kb2() As Double
    Kb2 = mKb2
End Property

Public Property Get Ks1() As Double
    Ks1 = mKs1
End Property

Public Property Get Ks2() As Double
    Ks2 = mKs2
End Property

Public Property Get StiffnessD1() As Double
    StiffnessD1 = mStiffnessD1
End Property

Public Property Get StiffnessD2() As Double
    StiffnessD2 = mStiffnessD2
End Property

Public Property Get PlaneBranch1() As String
    PlaneBranch1 = mPlaneBranch1
End Property

Public Property Get PlaneBranch2() As String
    PlaneBranch2 = mPlaneBranch2
End Property

Public Property Get PlaneApplicable1() As Boolean
    PlaneApplicable1 = mPlaneApplicable1
End Property

Public Property Get PlaneApplicable2() As Boolean
    PlaneApplicable2 = mPlaneApplicable2
End Property

Public Property Get PlanePassed1() As Boolean
    PlanePassed1 = mPlanePassed1
End Property

Public Property Get PlanePassed2() As Boolean
    PlanePassed2 = mPlanePassed2
End Property


' Завершает заполнение результата устойчивости перед публикацией.
' Повторный вызов безопасен; сохраненные числа и meta защищаются от
' последующих reset/update, без копирования данных и повторного расчета.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Явно отклоняет изменение уже выданного результата вместо тихой порчи снимка.
' Для повторного расчета workflow создает новый объект, сохраняя прежние ссылки.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CStabilityResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
