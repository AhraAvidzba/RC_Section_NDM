VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPostResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемый типизированный результат проверки T и суммы NM+T
' ==========================================================================
' Хранит смысловые величины во внутренних единицах и массивы физических стержней.
' Признаки наличия отличают неприменимую величину от вычисленного нуля.
' Матрицы Excel и снимка собирает CTorsionSummaryWriter; номера полей здесь неизвестны.

Private mData As TTorsionPostData
Private mBarCount As Long
Private mBarIDs() As String, mBarPlacements() As String
Private mBarForces() As Double, mBarLengths() As Double
Private mStateFailureReason As String
Private mFrozen As Boolean ' Заполнение разрешено только до передачи результата потребителю.

' Принимает готовые значения и копирует типизированные массивы без объектов на каждый стержень.
Friend Sub Initialize(ByRef values As TTorsionPostData, ByVal barCount As Long, _
        ByRef barIDs() As String, ByRef placements() As String, _
        ByRef forces() As Double, ByRef lengths() As Double, _
        Optional ByVal stateFailureReason As String = "")
    AssertWritable
    If barCount < 0 Then ResultError "Количество стержней результата не может быть отрицательным."
    If barCount > 0 Then
        If LBound(barIDs) <> 1 Or UBound(barIDs) <> barCount Then ResultError "Размер массива ID стержней не соответствует результату."
        If LBound(placements) <> 1 Or UBound(placements) <> barCount Then ResultError "Размер массива граней не соответствует результату."
        If LBound(forces) <> 1 Or UBound(forces) <> barCount Then ResultError "Размер массива усилий не соответствует результату."
        If LBound(lengths) <> 1 Or UBound(lengths) <> barCount Then ResultError "Размер массива длин не соответствует результату."
    End If
    CopyData values, mData
    mBarCount = barCount
    If barCount > 0 Then
        ReDim mBarIDs(1 To barCount), mBarPlacements(1 To barCount)
        ReDim mBarForces(1 To barCount), mBarLengths(1 To barCount)
        Dim i As Long
        For i = 1 To barCount
            mBarIDs(i) = barIDs(i): mBarPlacements(i) = placements(i)
            mBarForces(i) = forces(i): mBarLengths(i) = lengths(i)
        Next i
    End If
    mStateFailureReason = stateFailureReason
    mFrozen = True
End Sub

' Создаёт ошибку T без вымышленных численных результатов и без изменения метаданных НДМ.
Friend Sub InitializeFailure(ByVal identifier As String, ByVal reason As String, Optional ByVal internalError As Boolean = False)
    AssertWritable
    Dim meta As New CResultMeta, policy As New CResultStatusPolicy
    If internalError Then
        meta.SetResult rsInternalError, rcInternalError, rkGeneric, reason
    Else
        meta.SetResult rsInvalidInput, rcInvalidInput, rkGeneric, reason
    End If
    mData.CombinationID = identifier
    mData.LongitudinalStatus = policy.ExternalStatus(meta)
    mData.HoopStatus = mData.LongitudinalStatus: mData.ConcreteStatus = mData.LongitudinalStatus
    mData.CombinedStatus = mData.LongitudinalStatus: mData.OverallStatus = mData.LongitudinalStatus
    mData.ResultComment = reason: mData.HasResultComment = True
    mFrozen = True
End Sub

' Создаёт адресную копию с другим пояснением, сохраняя исходные величины и стержни.
Friend Function WithResultComment(ByVal text As String) As CTorsionPostResult
    Dim values As TTorsionPostData
    CopyData mData, values
    values.ResultComment = text: values.HasResultComment = True
    Dim result As New CTorsionPostResult
    result.Initialize values, mBarCount, mBarIDs, mBarPlacements, mBarForces, mBarLengths, mStateFailureReason
    Set WithResultComment = result
End Function

' Копирует значения по именам полей; не преобразует их в Variant или табличные позиции.
Private Sub CopyData(ByRef source As TTorsionPostData, ByRef destination As TTorsionPostData)
    destination.HasTResults = source.HasTResults
    destination.HasConcreteCapacity = source.HasConcreteCapacity
    destination.HasConcreteStress = source.HasConcreteStress
    destination.HasCombinedResults = source.HasCombinedResults
    destination.HasCombinedReserve = source.HasCombinedReserve
    destination.HasResultComment = source.HasResultComment
    destination.CombinationID = source.CombinationID
    destination.StateName = source.StateName
    destination.Torque = source.Torque
    destination.EnclosedArea = source.EnclosedArea
    destination.HoopPerimeter = source.HoopPerimeter
    destination.Cotangent = source.Cotangent
    destination.ShearFlow = source.ShearFlow
    destination.LongitudinalForce = source.LongitudinalForce
    destination.LongitudinalArea = source.LongitudinalArea
    destination.RequiredLongitudinalArea = source.RequiredLongitudinalArea
    destination.CriticalTRebarID = source.CriticalTRebarID
    destination.MaximumTStress = source.MaximumTStress
    destination.LongitudinalResistance = source.LongitudinalResistance
    destination.LongitudinalReserve = source.LongitudinalReserve
    destination.LongitudinalStatus = source.LongitudinalStatus
    destination.HoopArea = source.HoopArea
    destination.HoopSpacing = source.HoopSpacing
    destination.HoopStress = source.HoopStress
    destination.HoopResistance = source.HoopResistance
    destination.HoopReserve = source.HoopReserve
    destination.HoopStatus = source.HoopStatus
    destination.ConcreteMethod = source.ConcreteMethod
    destination.ConcreteThickness = source.ConcreteThickness
    destination.ConcreteStress = source.ConcreteStress
    destination.ConcreteResistance = source.ConcreteResistance
    destination.ConcreteTorqueLimit = source.ConcreteTorqueLimit
    destination.ConcreteReserve = source.ConcreteReserve
    destination.ConcreteStatus = source.ConcreteStatus
    destination.CriticalNMTRebarID = source.CriticalNMTRebarID
    destination.NormalStress = source.NormalStress
    destination.TStressAtCriticalNMT = source.TStressAtCriticalNMT
    destination.CombinedForce = source.CombinedForce
    destination.CombinedStress = source.CombinedStress
    destination.CombinedResistance = source.CombinedResistance
    destination.CombinedReserve = source.CombinedReserve
    destination.CombinedStatus = source.CombinedStatus
    destination.BarBand = source.BarBand
    destination.OverallStatus = source.OverallStatus
    destination.ResultComment = source.ResultComment
    destination.EffectiveConcreteArea = source.EffectiveConcreteArea
    destination.ConcreteDesignStrength = source.ConcreteDesignStrength
    destination.ConcreteEfficiency = source.ConcreteEfficiency
    destination.CriticalTPlacement = source.CriticalTPlacement
    destination.CriticalNMTPlacement = source.CriticalNMTPlacement
End Sub

' Получены расчётные данные T.
Public Property Get HasTResults() As Boolean
    HasTResults = mData.HasTResults
End Property

' Получен предел бетона.
Public Property Get HasConcreteCapacity() As Boolean
    HasConcreteCapacity = mData.HasConcreteCapacity
End Property

' Получены напряжение и параметры бетонного раскоса.
Public Property Get HasConcreteStress() As Boolean
    HasConcreteStress = mData.HasConcreteStress
End Property

' Доступна сумма напряжений NM и T.
Public Property Get HasCombinedResults() As Boolean
    HasCombinedResults = mData.HasCombinedResults
End Property

' Определён конечный растягивающий запас NMT.
Public Property Get HasCombinedReserve() As Boolean
    HasCombinedReserve = mData.HasCombinedReserve
End Property

' Комментарий задан явно, включая пустую строку.
Public Property Get HasResultComment() As Boolean
    HasResultComment = mData.HasResultComment
End Property

' Возвращает: iD сочетания.
Public Property Get CombinationID() As String
    CombinationID = mData.CombinationID
End Property

' Возвращает: имя состояния НДМ.
Public Property Get StateName() As String
    StateName = mData.StateName
End Property

' Возвращает: модуль крутящего момента, Н*мм.
Public Property Get Torque() As Double
    Torque = mData.Torque
End Property

' Возвращает: площадь внутри оси хомута, мм2.
Public Property Get EnclosedArea() As Double
    EnclosedArea = mData.EnclosedArea
End Property

' Возвращает: периметр оси хомута, мм.
Public Property Get HoopPerimeter() As Double
    HoopPerimeter = mData.HoopPerimeter
End Property

' Возвращает: котангенс угла раскоса.
Public Property Get Cotangent() As Double
    Cotangent = mData.Cotangent
End Property

' Возвращает: касательный поток, Н/мм.
Public Property Get ShearFlow() As Double
    ShearFlow = mData.ShearFlow
End Property

' Возвращает: суммарная продольная тяга от T, Н.
Public Property Get LongitudinalForce() As Double
    LongitudinalForce = mData.LongitudinalForce
End Property

' Возвращает: площадь выбранных продольных стержней, мм2.
Public Property Get LongitudinalArea() As Double
    LongitudinalArea = mData.LongitudinalArea
End Property

' Возвращает: требуемая площадь продольной арматуры от T, мм2.
Public Property Get RequiredLongitudinalArea() As Double
    RequiredLongitudinalArea = mData.RequiredLongitudinalArea
End Property

' Возвращает: определяющий стержень от T.
Public Property Get CriticalTRebarID() As String
    CriticalTRebarID = mData.CriticalTRebarID
End Property

' Возвращает: наибольшее напряжение от T, МПа.
Public Property Get MaximumTStress() As Double
    MaximumTStress = mData.MaximumTStress
End Property

' Возвращает: сопротивление продольной арматуры, МПа.
Public Property Get LongitudinalResistance() As Double
    LongitudinalResistance = mData.LongitudinalResistance
End Property

' Возвращает: запас продольной арматуры от T.
Public Property Get LongitudinalReserve() As Double
    LongitudinalReserve = mData.LongitudinalReserve
End Property

' Возвращает: статус продольной арматуры от T.
Public Property Get LongitudinalStatus() As String
    LongitudinalStatus = mData.LongitudinalStatus
End Property

' Возвращает: площадь сечения хомута, мм2.
Public Property Get HoopArea() As Double
    HoopArea = mData.HoopArea
End Property

' Возвращает: шаг хомута, мм.
Public Property Get HoopSpacing() As Double
    HoopSpacing = mData.HoopSpacing
End Property

' Возвращает: напряжение в хомуте, МПа.
Public Property Get HoopStress() As Double
    HoopStress = mData.HoopStress
End Property

' Возвращает: сопротивление хомута, МПа.
Public Property Get HoopResistance() As Double
    HoopResistance = mData.HoopResistance
End Property

' Возвращает: запас хомута.
Public Property Get HoopReserve() As Double
    HoopReserve = mData.HoopReserve
End Property

' Возвращает: статус хомута.
Public Property Get HoopStatus() As String
    HoopStatus = mData.HoopStatus
End Property

' Возвращает: метод проверки бетона.
Public Property Get ConcreteMethod() As String
    ConcreteMethod = mData.ConcreteMethod
End Property

' Возвращает: эффективная толщина бетона, мм; при HasConcreteCapacity.
Public Property Get ConcreteThickness() As Double
    ConcreteThickness = mData.ConcreteThickness
End Property

' Возвращает: напряжение бетонного раскоса, МПа; при HasConcreteStress.
Public Property Get ConcreteStress() As Double
    ConcreteStress = mData.ConcreteStress
End Property

' Возвращает: допускаемое напряжение раскоса, МПа; при HasConcreteStress.
Public Property Get ConcreteResistance() As Double
    ConcreteResistance = mData.ConcreteResistance
End Property

' Возвращает: предельный крутящий момент бетона, Н*мм; при HasConcreteCapacity.
Public Property Get ConcreteTorqueLimit() As Double
    ConcreteTorqueLimit = mData.ConcreteTorqueLimit
End Property

' Возвращает: запас бетона; при HasConcreteCapacity.
Public Property Get ConcreteReserve() As Double
    ConcreteReserve = mData.ConcreteReserve
End Property

' Возвращает: статус бетона.
Public Property Get ConcreteStatus() As String
    ConcreteStatus = mData.ConcreteStatus
End Property

' Возвращает: определяющий стержень суммы NMT; при HasCombinedResults.
Public Property Get CriticalNMTRebarID() As String
    CriticalNMTRebarID = mData.CriticalNMTRebarID
End Property

' Возвращает: напряжение NM в определяющем стержне, МПа; при HasCombinedResults.
Public Property Get NormalStress() As Double
    NormalStress = mData.NormalStress
End Property

' Возвращает: добавка T в определяющем стержне, МПа; при HasCombinedResults.
Public Property Get TStressAtCriticalNMT() As Double
    TStressAtCriticalNMT = mData.TStressAtCriticalNMT
End Property

' Возвращает: суммарное усилие NMT в определяющем стержне, Н; при HasCombinedResults.
Public Property Get CombinedForce() As Double
    CombinedForce = mData.CombinedForce
End Property

' Возвращает: суммарное напряжение NMT, МПа; при HasCombinedResults.
Public Property Get CombinedStress() As Double
    CombinedStress = mData.CombinedStress
End Property

' Возвращает: сопротивление для суммы NMT, МПа; при HasCombinedResults.
Public Property Get CombinedResistance() As Double
    CombinedResistance = mData.CombinedResistance
End Property

' Возвращает: конечный запас NMT; при HasCombinedReserve.
Public Property Get CombinedReserve() As Double
    CombinedReserve = mData.CombinedReserve
End Property

' Возвращает: статус суммы NMT.
Public Property Get CombinedStatus() As String
    CombinedStatus = mData.CombinedStatus
End Property

' Возвращает: ширина полосы отбора стержней, мм.
Public Property Get BarBand() As Double
    BarBand = mData.BarBand
End Property

' Возвращает: общий статус постпроверки.
Public Property Get OverallStatus() As String
    OverallStatus = mData.OverallStatus
End Property

' Возвращает: пояснение результата; при HasResultComment.
Public Property Get ResultComment() As String
    ResultComment = mData.ResultComment
End Property

' Возвращает: расчётная площадь бетона, мм2; при HasConcreteStress.
Public Property Get EffectiveConcreteArea() As Double
    EffectiveConcreteArea = mData.EffectiveConcreteArea
End Property

' Возвращает: расчётная прочность бетона f_cd, МПа; при HasConcreteStress.
Public Property Get ConcreteDesignStrength() As Double
    ConcreteDesignStrength = mData.ConcreteDesignStrength
End Property

' Возвращает: коэффициент эффективности бетона; при HasConcreteStress.
Public Property Get ConcreteEfficiency() As Double
    ConcreteEfficiency = mData.ConcreteEfficiency
End Property

' Возвращает: смысловая грань определяющего стержня от T.
Public Property Get CriticalTPlacement() As String
    CriticalTPlacement = mData.CriticalTPlacement
End Property

' Возвращает: смысловая грань определяющего стержня NMT; при HasCombinedResults.
Public Property Get CriticalNMTPlacement() As String
    CriticalNMTPlacement = mData.CriticalNMTPlacement
End Property

' Число физических стержней, участвующих в распределении тяги.
Public Property Get BarCount() As Long
    BarCount = mBarCount
End Property

' Возвращает ID физического стержня; индекс следует порядку физических стержней результата.
Public Property Get BarID(ByVal index As Long) As String
    ValidateBarIndex index
    BarID = mBarIDs(index)
End Property

' Возвращает смысловую грань стержня; индекс следует порядку физических стержней результата.
Public Property Get BarPlacement(ByVal index As Long) As String
    ValidateBarIndex index
    BarPlacement = mBarPlacements(index)
End Property

' Возвращает добавочное усилие от T, Н; индекс следует порядку физических стержней результата.
Public Property Get BarForce(ByVal index As Long) As Double
    ValidateBarIndex index
    BarForce = mBarForces(index)
End Property

' Возвращает распределённую длину оси хомута, мм; индекс следует порядку физических стержней результата.
Public Property Get BarLength(ByVal index As Long) As Double
    ValidateBarIndex index
    BarLength = mBarLengths(index)
End Property

' Причина недоступности НДС отделена от ошибок бетонной проверки.
Public Property Get StateFailureReason() As String
    StateFailureReason = mStateFailureReason
End Property

' Защищает уже переданный результат от повторного заполнения.
Private Sub AssertWritable()
    If mFrozen Then ResultError "Готовый результат кручения нельзя изменять."
End Sub

' Проверяет индекс до доступа к внутренним массивам.
Private Sub ValidateBarIndex(ByVal index As Long)
    If index < 1 Or index > mBarCount Then ResultError "Индекс стержня выходит за границы результата кручения."
End Sub

' Ошибка контракта объекта является внутренней, а не ошибкой пользовательского ввода.
Private Sub ResultError(ByVal reason As String)
    Err.Raise TORSION_RESULT_CONTRACT_ERROR, TypeName(Me), reason
End Sub
