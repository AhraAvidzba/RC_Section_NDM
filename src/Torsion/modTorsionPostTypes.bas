Attribute VB_Name = "modTorsionPostTypes"
Option Explicit

' ==========================================================================
' Данные отдельной приближённой проверки арматуры и хомута на кручение
' ==========================================================================
' Не являются нагрузками решателя НДМ. Размерности: мм, Н, Н*мм, МПа.
' Проверочная сумма напряжений не задаёт новое равновесное НДС.

Public Const TORSION_POST_COLUMNS As Long = 44
Public Const TORSION_POST_DISPLAY_COLUMNS As Long = 48 ' Самостоятельный блок с расчётной площадью, f_cd и коэффициентом nu проверки бетона.
Public Const TORSION_BAR_COLUMNS As Long = 5 ' ID сочетания, стержень, смысловая грань, усилие от T и распределённая длина.
Public Const TORSION_BATCH_STATUS_COLUMN As Long = 16
Public Const TORSION_BATCH_RESERVE_COLUMN As Long = 30
Public Const TORSION_POST_DATA_KEY As String = "TorsionPost.Data"
Public Const TORSION_POST_HEADERS_KEY As String = "TorsionPost.Headers"
Public Const TORSION_BAR_DATA_KEY As String = "TorsionPost.Bars"
Public Const TORSION_RESULT_CONTRACT_ERROR As Long = vbObjectError + 5671 ' Нарушение внутреннего контракта результата, не ошибка ввода.

' Номера полей сериализации принадлежат слою вывода, не расчётному ядру.
Public Enum ETorsionSummaryColumn
    tscCombinationID = 1
    tscStateName = 2
    tscTorque = 3
    tscEnclosedArea = 4
    tscHoopPerimeter = 5
    tscCotangent = 6
    tscShearFlow = 7
    tscLongitudinalForce = 8
    tscLongitudinalArea = 9
    tscRequiredLongitudinalArea = 10
    tscCriticalTRebarID = 11
    tscMaximumTStress = 12
    tscLongitudinalResistance = 13
    tscLongitudinalReserve = 14
    tscLongitudinalStatus = 15
    tscHoopArea = 16
    tscHoopSpacing = 17
    tscHoopStress = 18
    tscHoopResistance = 19
    tscHoopReserve = 20
    tscHoopStatus = 21
    tscConcreteMethod = 22
    tscConcreteThickness = 23
    tscConcreteStress = 24
    tscConcreteResistance = 25
    tscConcreteTorqueLimit = 26
    tscConcreteReserve = 27
    tscConcreteStatus = 28
    tscCriticalNMTRebarID = 29
    tscNormalStress = 30
    tscTStressAtCriticalNMT = 31
    tscCombinedForce = 32
    tscCombinedStress = 33
    tscCombinedResistance = 34
    tscCombinedReserve = 35
    tscCombinedStatus = 36
    tscBarBand = 37
    tscOverallStatus = 38
    tscResultComment = 39
    tscEffectiveConcreteArea = 40
    tscConcreteDesignStrength = 41
    tscConcreteEfficiency = 42
    tscCriticalTPlacement = 43
    tscCriticalNMTPlacement = 44
End Enum

Public Enum ETorsionBarColumn
    tbcCombinationID = 1
    tbcRebarID = 2
    tbcPlacement = 3
    tbcForce = 4
    tbcAllocatedLength = 5
End Enum

' Типизированные значения одного результата. Не содержит адресов Excel и номеров столбцов.
' Признаки наличия различают вычисленный ноль и неприменимую величину.
Public Type TTorsionPostData
    HasTResults As Boolean ' Получены расчётные данные T.
    HasConcreteCapacity As Boolean ' Получен предел бетона.
    HasConcreteStress As Boolean ' Получены напряжение и параметры бетонного раскоса.
    HasCombinedResults As Boolean ' Доступна сумма напряжений NM и T.
    HasCombinedReserve As Boolean ' Определён конечный растягивающий запас NMT.
    HasResultComment As Boolean ' Комментарий задан явно, включая пустую строку.
    CombinationID As String ' ID сочетания.
    StateName As String ' Имя состояния НДМ.
    Torque As Double ' Модуль крутящего момента, Н*мм.
    EnclosedArea As Double ' Площадь внутри оси хомута, мм2.
    HoopPerimeter As Double ' Периметр оси хомута, мм.
    Cotangent As Double ' Котангенс угла раскоса.
    ShearFlow As Double ' Касательный поток, Н/мм.
    LongitudinalForce As Double ' Суммарная продольная тяга от T, Н.
    LongitudinalArea As Double ' Площадь выбранных продольных стержней, мм2.
    RequiredLongitudinalArea As Double ' Требуемая площадь продольной арматуры от T, мм2.
    CriticalTRebarID As String ' Определяющий стержень от T.
    MaximumTStress As Double ' Наибольшее напряжение от T, МПа.
    LongitudinalResistance As Double ' Сопротивление продольной арматуры, МПа.
    LongitudinalReserve As Double ' Запас продольной арматуры от T.
    LongitudinalStatus As String ' Статус продольной арматуры от T.
    HoopArea As Double ' Площадь сечения хомута, мм2.
    HoopSpacing As Double ' Шаг хомута, мм.
    HoopStress As Double ' Напряжение в хомуте, МПа.
    HoopResistance As Double ' Сопротивление хомута, МПа.
    HoopReserve As Double ' Запас хомута.
    HoopStatus As String ' Статус хомута.
    ConcreteMethod As String ' Метод проверки бетона.
    ConcreteThickness As Double ' Эффективная толщина бетона, мм; при HasConcreteCapacity.
    ConcreteStress As Double ' Напряжение бетонного раскоса, МПа; при HasConcreteStress.
    ConcreteResistance As Double ' Допускаемое напряжение раскоса, МПа; при HasConcreteStress.
    ConcreteTorqueLimit As Double ' Предельный крутящий момент бетона, Н*мм; при HasConcreteCapacity.
    ConcreteReserve As Double ' Запас бетона; при HasConcreteCapacity.
    ConcreteStatus As String ' Статус бетона.
    CriticalNMTRebarID As String ' Определяющий стержень суммы NMT; при HasCombinedResults.
    NormalStress As Double ' Напряжение NM в определяющем стержне, МПа; при HasCombinedResults.
    TStressAtCriticalNMT As Double ' Добавка T в определяющем стержне, МПа; при HasCombinedResults.
    CombinedForce As Double ' Суммарное усилие NMT в определяющем стержне, Н; при HasCombinedResults.
    CombinedStress As Double ' Суммарное напряжение NMT, МПа; при HasCombinedResults.
    CombinedResistance As Double ' Сопротивление для суммы NMT, МПа; при HasCombinedResults.
    CombinedReserve As Double ' Конечный запас NMT; при HasCombinedReserve.
    CombinedStatus As String ' Статус суммы NMT.
    BarBand As Double ' Ширина полосы отбора стержней, мм.
    OverallStatus As String ' Общий статус постпроверки.
    ResultComment As String ' Пояснение результата; при HasResultComment.
    EffectiveConcreteArea As Double ' Расчётная площадь бетона, мм2; при HasConcreteStress.
    ConcreteDesignStrength As Double ' Расчётная прочность бетона f_cd, МПа; при HasConcreteStress.
    ConcreteEfficiency As Double ' Коэффициент эффективности бетона; при HasConcreteStress.
    CriticalTPlacement As String ' Смысловая грань определяющего стержня от T.
    CriticalNMTPlacement As String ' Смысловая грань определяющего стержня NMT; при HasCombinedResults.
End Type

Public Type TTorsionPostParameters
    StateName As String
    AxisOffset As Double
    HoopDiameter As Double
    HoopSpacing As Double
    SteelResistance As Double
    HoopResistance As Double
    BarBand As Double
    StrutAngle As Double
    ConcreteMethod As String
    ConcreteInputInvalid As Boolean ' Ошибка только исходных данных проверки бетона; арматура и хомут рассчитываются независимо.
    WallThickness As Double
    ConcreteResistance As Double
    CharacteristicConcreteStrength As Double ' f_ck: характеристическая цилиндрическая прочность, МПа; не класс B и не R_b.
    ConcretePartialFactor As Double ' gamma_c: коэффициент надёжности для вычисления f_cd.
    LongitudinalAxisOffset As Double ' c_l: расстояние от бетонной поверхности до центра продольной арматуры, мм.
    GorskyCoefficient As Double
End Type

' Возвращает внешний статус применимой проверки по общей политике проекта.
Public Function TorsionPostCheckStatus(ByVal reserve As Double) As String
    Dim meta As New CResultMeta, policy As New CResultStatusPolicy
    If reserve >= 1# Then
        meta.SetResult rsSuccess, rcCheckPassed, rkGeneric
    Else
        meta.SetResult rsCheckFailed, rcCheckFailed, rkGeneric
    End If
    TorsionPostCheckStatus = policy.ExternalStatus(meta)
End Function
