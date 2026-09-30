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

Public Status As String ' Пользовательский статус ветки устойчивости.
Public ResultMeta As CResultMeta ' Типизированный статус устойчивости.
Public Code As String ' Нормативная ветка SP35/SP63.
Public Branch As String ' Диагностическое имя расчетной ветви.
Public ReferenceX As Double ' X точки, относительно которой выполнен перенос нагрузок.
Public ReferenceY As Double ' Y точки, относительно которой выполнен перенос нагрузок.
Public DesignN As Double ' Итоговая продольная сила для последующего НДМ.
Public DesignMx As Double ' Итоговый Mx после учета устойчивости/фильтров.
Public DesignMy As Double ' Итоговый My после учета устойчивости/фильтров.
Public SustainedN As Double ' Длительная/постоянная часть N.
Public Moment1 As Double
Public Moment2 As Double
Public SustainedMoment1 As Double
Public SustainedMoment2 As Double
Public DesignMoment1 As Double
Public DesignMoment2 As Double
Public CriticalForce As Double
Public ReserveFactor As Double
Public Utilization As Double
Public Ncr1 As Double
Public Ncr2 As Double
Public Nultimate1 As Double
Public Nultimate2 As Double
Public Reserve1 As Double
Public Reserve2 As Double
Public Eta1 As Double
Public Eta2 As Double
Public PhiL1 As Double
Public PhiL2 As Double
Public PhiLTable1 As String
Public PhiLTable2 As String
Public PhiM1 As Double
Public PhiM2 As Double
Public PhiValue1 As Double
Public PhiValue2 As Double
Public PhiP1 As Double
Public PhiP2 As Double
Public Eccentricity1 As Double
Public Eccentricity2 As Double
Public AccidentalEcc1 As Double
Public AccidentalEcc2 As Double
Public EffectiveLength1 As Double
Public EffectiveLength2 As Double
Public Slenderness1 As Double
Public Slenderness2 As Double
Public Depth1 As Double
Public Depth2 As Double
Public Radius1 As Double
Public Radius2 As Double
Public CoreDistance1 As Double
Public CoreDistance2 As Double
Public Delta1 As Double
Public Delta2 As Double
Public Kb1 As Double
Public Kb2 As Double
Public Ks1 As Double
Public Ks2 As Double
Public StiffnessD1 As Double
Public StiffnessD2 As Double
Public PlaneBranch1 As String
Public PlaneBranch2 As String
Public PlaneApplicable1 As Boolean
Public PlaneApplicable2 As Boolean
Public PlanePassed1 As Boolean
Public PlanePassed2 As Boolean

' Создает N/A-result для случаев, когда устойчивость не запрошена профилем.
Private Sub Class_Initialize()
    Clear
End Sub

' Очищает stability branch и возвращает его в N/A.
Public Sub Clear()
    Status = "N/A"
    Set ResultMeta = New CResultMeta
    ResultMeta.SetNotApplicable rkStability
    Code = vbNullString
    Branch = vbNullString
    ReferenceX = 0#
    ReferenceY = 0#
    DesignN = 0#
    DesignMx = 0#
    DesignMy = 0#
    SustainedN = 0#
    Moment1 = 0#
    Moment2 = 0#
    SustainedMoment1 = 0#
    SustainedMoment2 = 0#
    DesignMoment1 = 0#
    DesignMoment2 = 0#
    CriticalForce = 0#
    ReserveFactor = 0#
    Utilization = 0#
End Sub

' Синхронизирует typed meta со строкой display-статуса для Results.
Public Sub SetMeta(ByVal meta As CResultMeta, Optional ByVal resultComment As String = vbNullString)
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    If meta Is Nothing Then
        Set ResultMeta = New CResultMeta
        ResultMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован."
    Else
        Set ResultMeta = meta.Clone
        If Len(Trim$(resultComment)) > 0 Then ResultMeta.ResultComment = StabilityBranchDisplayName(resultComment)
    End If
    Status = policy.ExternalStatus(ResultMeta)
    Branch = StabilityBranchDisplayName(resultComment)
End Sub

' Копирует результат CStabilityCalculator и формирует typed meta для summary.
Public Sub InitializeFromCalculator(ByVal meta As CResultMeta, ByVal stability As CStabilityCalculator, _
        ByVal referenceXValue As Double, ByVal referenceYValue As Double)
    ReferenceX = referenceXValue
    ReferenceY = referenceYValue

    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    If meta Is Nothing Then
        Set ResultMeta = New CResultMeta
        ResultMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован."
    Else
        Set ResultMeta = meta.Clone
    End If
    Status = policy.ExternalStatus(ResultMeta)

    If stability Is Nothing Then
        Exit Sub
    End If

    Code = stability.Code
    Branch = StabilityBranchDisplayName(stability.Branch)
    If Len(Trim$(stability.StopReason)) > 0 Then Branch = stability.StopReason
    DesignN = stability.DesignN
    DesignMx = stability.DesignMx
    DesignMy = stability.DesignMy
    SustainedN = stability.SustainedN
    Moment1 = stability.Moment1
    Moment2 = stability.Moment2
    SustainedMoment1 = stability.SustainedMoment1
    SustainedMoment2 = stability.SustainedMoment2
    DesignMoment1 = stability.DesignMoment1
    DesignMoment2 = stability.DesignMoment2
    CriticalForce = stability.CriticalForce
    ReserveFactor = stability.ReserveFactor
    If ReserveFactor > 0# Then Utilization = 1# / ReserveFactor
    Ncr1 = stability.Ncr1
    Ncr2 = stability.Ncr2
    Nultimate1 = stability.Nultimate1
    Nultimate2 = stability.Nultimate2
    Reserve1 = stability.Reserve1
    Reserve2 = stability.Reserve2
    Eta1 = stability.Eta1
    Eta2 = stability.Eta2
    PhiL1 = stability.PhiL1
    PhiL2 = stability.PhiL2
    PhiLTable1 = stability.PhiLTable1
    PhiLTable2 = stability.PhiLTable2
    PhiM1 = stability.PhiM1
    PhiM2 = stability.PhiM2
    PhiValue1 = stability.PhiValue1
    PhiValue2 = stability.PhiValue2
    PhiP1 = stability.PhiP1
    PhiP2 = stability.PhiP2
    Eccentricity1 = stability.Eccentricity1
    Eccentricity2 = stability.Eccentricity2
    AccidentalEcc1 = stability.AccidentalEccentricity1
    AccidentalEcc2 = stability.AccidentalEccentricity2
    EffectiveLength1 = stability.EffectiveLength1
    EffectiveLength2 = stability.EffectiveLength2
    Slenderness1 = stability.Slenderness1
    Slenderness2 = stability.Slenderness2
    Depth1 = stability.Depth1
    Depth2 = stability.Depth2
    Radius1 = stability.Radius1
    Radius2 = stability.Radius2
    CoreDistance1 = stability.CoreDistance1
    CoreDistance2 = stability.CoreDistance2
    Delta1 = stability.Delta1
    Delta2 = stability.Delta2
    Kb1 = stability.Kb1
    Kb2 = stability.Kb2
    Ks1 = stability.Ks1
    Ks2 = stability.Ks2
    StiffnessD1 = stability.StiffnessD1
    StiffnessD2 = stability.StiffnessD2
    PlaneBranch1 = stability.PlaneBranch1
    PlaneBranch2 = stability.PlaneBranch2
    PlaneApplicable1 = stability.PlaneApplicable1
    PlaneApplicable2 = stability.PlaneApplicable2
    PlanePassed1 = stability.PlanePassed1
    PlanePassed2 = stability.PlanePassed2
    If Len(Trim$(Branch)) > 0 And Len(Trim$(ResultMeta.ResultComment)) = 0 Then _
        ResultMeta.ResultComment = Branch
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
    DesignN = designNValue
    DesignMx = designMxValue
    DesignMy = designMyValue
    DesignMoment1 = designMxValue * Cos(principalAngleRad) + designMyValue * Sin(principalAngleRad)
    DesignMoment2 = -designMxValue * Sin(principalAngleRad) + designMyValue * Cos(principalAngleRad)
End Sub

' Возвращает копию meta, чтобы вызывающий код не менял status-subtree напрямую.
Public Property Get Meta() As CResultMeta
    If ResultMeta Is Nothing Then
        Set ResultMeta = New CResultMeta
        ResultMeta.SetResult rsInternalError, rcInternalError, rkStability, _
            "Результат устойчивости не сформирован."
    End If
    Set Meta = ResultMeta.Clone
End Property
