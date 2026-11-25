VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPostCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Отдельная приближённая проверка замкнутой фермы на T после готового НДМ
' ==========================================================================
' Не вызывает решатель, не меняет нагрузки, диаграммы, состояния и результаты
' НДМ. Тяга от T распределяется по периметру на выбранные продольные стержни.
' sigmaNM + deltaSigmaT служит проверкой растяжения с учётом знака НДМ,
' а не новым НДС. Самостоятельные отказы НДМ остаются в его блоке Results.
' Бетонные диагонали проверяются отдельно от нормальных напряжений НДМ.

Private mSection As CSectionModel, mOuter As CTorsionHoopGeometry
Private mP As TTorsionPostParameters
Private mCoefficients() As Double
Private mPlacementNames() As String
Private mAllocatedLengths() As Double ' Доля периметра каждого физического стержня, мм; готовится один раз независимо от T.
Private mCotangent As Double, mHoopArea As Double
Private mConcrete As CTorsionConcreteCheck
Private mOuterArea As Double
Private mFailureSettingKeys As String
Private mConcreteFailureStatus As String, mConcreteFailureReason As String, mConcreteFailureKeys As String ' Отказ бетонной проверки не меняет подготовку продольной арматуры и хомута.

' Готовит геометрию один раз для всех сочетаний T. Входы уже приведены к
' внутренним единицам и проверены адаптером; геометрическая применимость здесь.
Friend Sub Prepare(ByVal section As CSectionModel, ByRef parameters As TTorsionPostParameters)
    If Not mSection Is Nothing Then ModelError "Геометрия этапа T уже подготовлена."
    If section Is Nothing Then ModelError "Не передана сохранённая геометрия сечения."
    Dim hasOuter As Boolean, contourIndex As Long
    For contourIndex = 1 To section.Contours.Count
        If StrComp(section.Contours.LoopRole(contourIndex), "Outer", vbTextCompare) = 0 Then hasOuter = True: Exit For
    Next contourIndex
    If Not hasOuter Then ModelError "Проверка T (кручение) не выполнена: в сохранённой геометрии нет достоверного наружного контура бетона. " & _
        "По одной бетонной сетке ось хомута не восстанавливается. Задайте замкнутый наружный контур при импорте и снова сформируйте сечение. Результаты NM сохраняются независимо."
    If section.RebarCount < 3 Then ModelError "Для проверки T нужны как минимум три наружных продольных стержня."
    Set mSection = section
    mP.StateName = parameters.StateName: mP.AxisOffset = parameters.AxisOffset
    mP.HoopDiameter = parameters.HoopDiameter: mP.HoopSpacing = parameters.HoopSpacing
    mP.SteelResistance = parameters.SteelResistance: mP.HoopResistance = parameters.HoopResistance
    mP.BarBand = parameters.BarBand: mP.StrutAngle = parameters.StrutAngle
    mP.ConcreteMethod = parameters.ConcreteMethod: mP.WallThickness = parameters.WallThickness
    mP.ConcreteInputInvalid = parameters.ConcreteInputInvalid
    mP.ConcreteResistance = parameters.ConcreteResistance
    mP.CharacteristicConcreteStrength = parameters.CharacteristicConcreteStrength
    mP.ConcretePartialFactor = parameters.ConcretePartialFactor
    mP.LongitudinalAxisOffset = parameters.LongitudinalAxisOffset
    mP.GorskyCoefficient = parameters.GorskyCoefficient
    mFailureSettingKeys = "Torsion.HoopAxisOffset,Torsion.HoopDiameter"
    Set mOuter = New CTorsionHoopGeometry
    mOuter.InitializeAutomatic section, mP.AxisOffset, mP.HoopDiameter
    Dim query As New CSectionGeometryQuery, domain As CConcreteRegion, holes As Long, i As Long
    Set domain = query.RegionFromContours(section.Contours)
    For i = 1 To domain.LoopCount
        If domain.LoopIsOpening(i) Then holes = holes + 1
    Next i
    If holes > 1 Then ModelError "Однокамерная проверка T не применяется к нескольким отверстиям."
    Dim demand As New CTorsionLongitudinalDemand
    mFailureSettingKeys = "Torsion.BarBand,Torsion.HoopAxisOffset,Torsion.HoopDiameter"
    demand.Initialize section, mOuter, mP.BarBand
    demand.CopyStressPerUnitForce mCoefficients
    ReDim mAllocatedLengths(1 To section.RebarCount), mPlacementNames(1 To section.RebarCount)
    For i = 1 To section.RebarCount
        mPlacementNames(i) = section.RebarComment(i)
        mAllocatedLengths(i) = mCoefficients(i) * section.RebarArea(i) * mOuter.Length
        If mCoefficients(i) > 0# Then mOuterArea = mOuterArea + section.RebarArea(i)
    Next i
    mCotangent = 1# / Tan(mP.StrutAngle * 4# * Atn(1#) / 180#)
    mHoopArea = 4# * Atn(1#) * mP.HoopDiameter ^ 2 / 4#
    PrepareConcreteCheck domain
End Sub

' Считает T независимо от физических пределов готового состояния. При
' отсутствии сохранённых напряжений только сумма T+NM получает N/A. Явный ноль NM не требует решателя или снимка материалов.
Friend Function Calculate(ByVal identifier As String, ByVal torque As Double, ByVal state As CSectionStateResult, _
        Optional ByVal zeroNormalState As Boolean = False) As CTorsionPostResult
    If mSection Is Nothing Then ModelError "Геометрия отдельного этапа T не подготовлена."
    Dim data As TTorsionPostData
    Dim barIDs() As String, placements() As String, forces() As Double, lengths() As Double
    torque = Abs(torque)
    If torque = 0# Then ModelError "Нулевая строка T не требует расчёта."
    Dim q As Double, force As Double, hoopStress As Double, concreteStress As Double
    q = torque / (2# * mOuter.EnclosedArea)
    force = q * mOuter.Length * mCotangent ' Вся тяга от T распределяется только на наружную продольную арматуру.
    hoopStress = q * mP.HoopSpacing / (mHoopArea * mCotangent)
    If Not mConcrete Is Nothing Then concreteStress = mConcrete.Stress(torque)
    Dim available As Boolean
    If Not state Is Nothing Then available = state.Converged And state.HasMaterialSnapshot
    Dim snapshot As Variant
    If available Then snapshot = state.PrepareElementSnapshot(mSection)
    If zeroNormalState Then available = True
    Dim count As Long, i As Long, outputRow As Long
    For i = 1 To mSection.RebarCount
        If mCoefficients(i) > 0# Then count = count + 1
    Next i
    ReDim barIDs(1 To count), placements(1 To count), forces(1 To count), lengths(1 To count)
    Dim delta As Double, normal As Double, combined As Double, maxDelta As Double, maxCombined As Double
    Dim criticalT As Long, criticalCombined As Long, criticalNM As Double, criticalDelta As Double, criticalForce As Double
    Dim policy As New CResultStatusPolicy
    For i = 1 To mSection.RebarCount
        If mCoefficients(i) > 0# Then
            outputRow = outputRow + 1
            delta = force * mCoefficients(i)
            barIDs(outputRow) = mSection.RebarID(i)
            placements(outputRow) = mPlacementNames(i)
            forces(outputRow) = delta * mSection.RebarArea(i): lengths(outputRow) = mAllocatedLengths(i)
            If criticalT = 0 Or delta > maxDelta Then criticalT = i: maxDelta = delta
            If available Then
                If zeroNormalState Then normal = 0# Else normal = CDbl(snapshot(mSection.ConcreteCount + i, 4))
                combined = normal + delta ' Положительное напряжение — растяжение; сжатие НДМ сохраняет знак минус.
                If criticalCombined = 0 Or combined > maxCombined Then
                    criticalCombined = i: maxCombined = combined
                    criticalNM = normal: criticalDelta = delta: criticalForce = combined * mSection.RebarArea(i)
                End If
            End If
        End If
    Next i
    data.HasTResults = True
    data.CombinationID = identifier: data.StateName = mP.StateName: data.Torque = torque
    data.EnclosedArea = mOuter.EnclosedArea: data.HoopPerimeter = mOuter.Length: data.Cotangent = mCotangent
    data.ShearFlow = q: data.LongitudinalForce = force: data.LongitudinalArea = mOuterArea
    data.CriticalTPlacement = mPlacementNames(criticalT)
    data.RequiredLongitudinalArea = force / mP.SteelResistance: data.CriticalTRebarID = mSection.RebarID(criticalT)
    data.MaximumTStress = maxDelta: data.LongitudinalResistance = mP.SteelResistance
    data.LongitudinalReserve = mP.SteelResistance / maxDelta: data.LongitudinalStatus = TorsionPostCheckStatus(data.LongitudinalReserve)
    data.HoopArea = mHoopArea: data.HoopSpacing = mP.HoopSpacing
    data.HoopStress = hoopStress: data.HoopResistance = mP.HoopResistance
    data.HoopReserve = mP.HoopResistance / hoopStress: data.HoopStatus = TorsionPostCheckStatus(data.HoopReserve)
    data.ConcreteMethod = mP.ConcreteMethod
    If Len(mConcreteFailureStatus) > 0 Then
        data.ConcreteStatus = mConcreteFailureStatus
    ElseIf mP.ConcreteMethod = "None" Then
        data.ConcreteStatus = policy.NotApplicable
    Else
        data.HasConcreteCapacity = True
        data.ConcreteThickness = mConcrete.Thickness
        If mP.ConcreteMethod = "Strut" Then
            data.HasConcreteStress = True
            data.ConcreteStress = concreteStress: data.ConcreteResistance = mConcrete.Resistance
            data.EffectiveConcreteArea = mConcrete.EffectiveArea: data.ConcreteDesignStrength = mConcrete.DesignStrength: data.ConcreteEfficiency = mConcrete.Efficiency
        End If
        data.ConcreteTorqueLimit = mConcrete.Limit: data.ConcreteReserve = mConcrete.Limit / torque
        data.ConcreteStatus = TorsionPostCheckStatus(data.ConcreteReserve)
    End If
    If available Then
        data.HasCombinedResults = True
        data.CriticalNMTPlacement = mPlacementNames(criticalCombined)
        data.CriticalNMTRebarID = mSection.RebarID(criticalCombined): data.NormalStress = criticalNM
        data.TStressAtCriticalNMT = criticalDelta: data.CombinedForce = criticalForce: data.CombinedStress = maxCombined
        data.CombinedResistance = mP.SteelResistance
        If maxCombined > 0# Then
            data.HasCombinedReserve = True
            data.CombinedReserve = mP.SteelResistance / maxCombined: data.CombinedStatus = TorsionPostCheckStatus(data.CombinedReserve)
        Else
            data.CombinedStatus = policy.OK ' При сжатии растягивающий предел не достигается; числовой запас не ограничен.
        End If
    Else
        data.CombinedStatus = policy.NotApplicable
        If state Is Nothing Then
            data.ResultComment = "Сохранённое состояние " & mP.StateName & " не найдено; T рассчитан, сумма NMT недоступна. Включите получение этого состояния в расчётном профиле."
        Else
            data.ResultComment = "Сохранённые напряжения выбранного состояния недоступны для сложения с добавкой T."
        End If
    End If
    Dim stateFailureReason As String: stateFailureReason = data.ResultComment
    If Len(mConcreteFailureReason) > 0 Then
        If Len(data.ResultComment) > 0 Then data.ResultComment = data.ResultComment & vbCrLf
        data.ResultComment = data.ResultComment & mConcreteFailureReason
    End If
    data.BarBand = mP.BarBand: data.OverallStatus = policy.OK
    data.OverallStatus = OverallStatus(data.HoopStatus, data.ConcreteStatus, data.CombinedStatus, policy)
    data.HasResultComment = (Len(data.ResultComment) > 0)
    Dim result As New CTorsionPostResult: result.Initialize data, count, barIDs, placements, forces, lengths, stateFailureReason
    Set Calculate = result
End Function

' Собирает статус NMT по независимым проверкам, без знания табличных позиций.
Private Function OverallStatus(ByVal hoopStatus As String, ByVal concreteStatus As String, _
        ByVal combinedStatus As String, ByVal policy As CResultStatusPolicy) As String
    OverallStatus = policy.OK
    Dim status As Variant
    For Each status In Array(hoopStatus, concreteStatus, combinedStatus)
        If CStr(status) = policy.CalcErr Then OverallStatus = policy.CalcErr: Exit Function
    Next status
    For Each status In Array(hoopStatus, concreteStatus, combinedStatus)
        If CStr(status) = policy.InputErr Then OverallStatus = policy.InputErr: Exit Function
    Next status
    For Each status In Array(hoopStatus, concreteStatus, combinedStatus)
        If CStr(status) = policy.Fail Then OverallStatus = policy.Fail: Exit Function
    Next status
    If combinedStatus = policy.NotApplicable Then OverallStatus = policy.NotApplicable
End Function

' Изолирует только подготовку бетонного предела. Общая геометрия, тяга и
' хомут к этому моменту готовы; неизвестный внутренний сбой сохраняет CalcErr.
Private Sub PrepareConcreteCheck(ByVal domain As CConcreteRegion)
    Dim policy As New CResultStatusPolicy
    If mP.ConcreteInputInvalid Then
        mConcreteFailureStatus = policy.InputErr
        mConcreteFailureReason = "Проверка бетона недоступна: исправьте её исходные данные."
        Exit Sub
    End If
    On Error GoTo Failed
    Set mConcrete = New CTorsionConcreteCheck
    mConcrete.Prepare domain, mP
    Exit Sub
Failed:
    mConcreteFailureReason = Err.Description: mConcreteFailureKeys = mConcrete.FailureSettingKeys
    If Err.Number = vbObjectError + 5670 Then mConcreteFailureStatus = policy.InputErr Else mConcreteFailureStatus = policy.CalcErr
    Set mConcrete = Nothing ' Неподтверждённый предел не попадает в вывод как действительное число.
End Sub

' Возвращает причину независимого отказа бетонной проверки без адресов Excel.
Public Property Get ConcreteFailureReason() As String
    ConcreteFailureReason = mConcreteFailureReason
End Property

' Возвращает ключи параметров только отказавшей бетонной проверки.
Public Property Get ConcreteFailureSettingKeys() As String
    ConcreteFailureSettingKeys = mConcreteFailureKeys
End Property

' Возвращает ключи входов текущего геометрического этапа для общего адаптера
' адресов; никаких ссылок на Excel или текстовых разборов причины в ядре нет.
Public Property Get FailureSettingKeys() As String
    FailureSettingKeys = mFailureSettingKeys
End Property

' Возвращает методическую или геометрическую причину без обращения к Excel.
Private Sub ModelError(ByVal reason As String)
    Err.Raise vbObjectError + 5670, TypeName(Me), reason
End Sub
