VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarRuleOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Дополнительные параметры расстановки арматуры
' ==========================================================================
' Хранит отступы, способ размещения концов ряда, зависимости и оформление.
' CSectionDefinitionBuilder копирует параметры при объявлении правила,
' поэтому последующее изменение этого объекта не меняет готовое описание сечения.

Public StartSetback As Double ' Отступ начала участка размещения от начала каждой линии, мм >= 0; для полного ClosedUniform равен 0.
Public EndSetback As Double ' Отступ конца участка размещения от конца линии, мм >= 0; при проекции ограничивает участок без перераспределения стержней.
Public EndMode As ERebarEndMode ' Какие концы ряда включать: оба, ни одного, только начало или конец; ClosedUniform задаёт равномерный замкнутый ряд.
Public DistributionScope As ERebarDistributionScope ' Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно к каждой независимой линии группы.
Public RemainderPlacement As ERemainderPlacement ' К какому концу или середине ряда отнести остаток длины при постоянном шаге.
Public TangentOffset As Double ' Смещение зависимого стержня вдоль касательной к исходной линии, мм; знак задаёт направление.
Public NormalOffset As Double ' Смещение зависимого стержня поперёк исходной линии, мм; положительный знак направлен влево по её обходу.
Public ValidateRowOffset As Boolean ' Активный пользовательский отступ требует проверки всех дополнительных стержней своего семейства.
Public EverySecond As Boolean ' Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal задаёт отдельный отсчёт успешных проекций.
Public IncomingTangentAtJoint As Boolean ' В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся касательная следующего сегмента.
Public DirectionX As Double ' Компонента X заданного направления проекции; используется, если SourceNormalSign = 0.
Public DirectionY As Double ' Компонента Y заданного направления проекции; вектор DirectionX/DirectionY должен быть ненулевым.
Public SourceNormalSign As Double ' Направление проекции: +1 по левой нормали исходной линии, -1 в обратную сторону, 0 по вектору DirectionX/DirectionY.
Public SuccessfulProjectionOrdinal As Boolean ' True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд по всей группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на каждой линии.
Public NewPublicationRoot As Boolean ' Для проекции задаёт новую последовательность добавления по успешно найденным позициям вместо порядка исходных стержней.
Public SkipMissingProjection As Boolean ' Разрешает пропуск проекции без пересечения с записью PROJECTION_MISSING; иначе отсутствие пересечения вызывает ошибку.
Public Publication As ERebarPublication ' Порядок добавления: сначала весь ряд или каждый основной стержень вместе с зависимыми; расстояния между стержнями не меняются.
Public PositionOrder As ERebarPositionOrder ' Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. При KeepPublishedOwner также определяет, чей диаметр и материал сохраняются в совпавшей позиции.
Public Coincidence As ERebarCoincidence ' Как обработать совпавшие центры: отклонить, объединить одинаковые стержни или сохранить ранее добавленный стержень.
Public PlacementCheck As ERebarPlacementCheck ' Что проверять относительно бетона и отверстий: только центр стержня или всё его круглое поперечное сечение.
Public OverlapPolicy As ERebarOverlapPolicy ' При пересечении стержней по площади: записать сообщение или отклонить расстановку. Касание допускается; нормативный зазор здесь не проверяется.
Public SourcePrefix As String ' Начало имени исходной позиции; если не задано, имя строится из ID правила или имени исходного стержня.
Public SourceSuffix As String ' Окончание имени исходной позиции, в том числе для различения зависимых рядов.
Public AnnotationGroup As String ' Группа для общей подписи количества и диаметров арматуры на схеме; задаётся отдельно от расчётных групп трещин СП 35.
Public Comment As String ' Пояснение правила для человека; не определяет способ расчёта или статус результата.
Public SourceReference As String ' Реальные адреса исходных параметров правила или программная ссылка.

' Задаёт начальные параметры шага, смещений, проекции, совпадений и подписей арматуры.
Private Sub Class_Initialize()
    EndMode = IncludeBoth
    DistributionScope = PerPath
    RemainderPlacement = RemainderCenter
    Publication = PublishByRule
    PositionOrder = AlongPath
    Coincidence = RejectCoincidence
    PlacementCheck = CheckCenter
    OverlapPolicy = ReportOverlap
End Sub

' Задаёт смещения зависимого ряда для касания стержней: учитывает половины диаметров
' основного и добавляемого стержней и полные диаметры промежуточных рядов.
' interveningDiameterSum также принимает добавки к расстоянию, подготовленные CRebarContactBuilder.
' Меняет NormalOffset, TangentOffset и EverySecond; координаты рассчитывает CRebarRuleExecutor.
' Оба диаметра должны быть положительными; отключённый ряд пропускается до вызова.
Public Sub ConfigureDependentContact(ByVal baseDiameter As Double, ByVal diameter As Double, _
        ByVal axis As ERebarOffsetAxis, ByVal directionSign As Double, ByVal everySecond As Boolean, _
        Optional ByVal interveningDiameterSum As Double = 0#)
    If baseDiameter <= 0# Or diameter <= 0# Or interveningDiameterSum < 0# Then _
        SectionRecipeError "CRebarRuleOptions", "Для контактного ряда нужны положительные диаметры и неотрицательная сумма промежуточных диаметров.", SourceReference
    If axis <> RowNormal And axis <> RowTangent Then SectionRecipeError "CRebarRuleOptions", "Для связанного ряда выберите смещение по нормали к линии (RowNormal) или вдоль неё (RowTangent).", SourceReference
    If directionSign <> -1# And directionSign <> 1# Then SectionRecipeError "CRebarRuleOptions", "Знак направления должен быть -1 или 1.", SourceReference
    Dim distance As Double
    distance = baseDiameter / 2# + interveningDiameterSum + diameter / 2#
    NormalOffset = 0#: TangentOffset = 0#
    If axis = RowNormal Then NormalOffset = directionSign * distance Else TangentOffset = directionSign * distance
    Me.EverySecond = everySecond
End Sub

' Возвращает независимую копию всех настроек; изменение одной копии не влияет на другую.
Public Function Clone() As CRebarRuleOptions
    Dim result As CRebarRuleOptions
    Set result = New CRebarRuleOptions
    result.StartSetback = StartSetback
    result.EndSetback = EndSetback
    result.EndMode = EndMode
    result.DistributionScope = DistributionScope
    result.RemainderPlacement = RemainderPlacement
    result.TangentOffset = TangentOffset
    result.NormalOffset = NormalOffset
    result.ValidateRowOffset = ValidateRowOffset
    result.EverySecond = EverySecond
    result.IncomingTangentAtJoint = IncomingTangentAtJoint
    result.DirectionX = DirectionX
    result.DirectionY = DirectionY
    result.SourceNormalSign = SourceNormalSign
    result.SuccessfulProjectionOrdinal = SuccessfulProjectionOrdinal
    result.NewPublicationRoot = NewPublicationRoot
    result.SkipMissingProjection = SkipMissingProjection
    result.Publication = Publication
    result.PositionOrder = PositionOrder
    result.Coincidence = Coincidence
    result.PlacementCheck = PlacementCheck
    result.OverlapPolicy = OverlapPolicy
    result.SourcePrefix = SourcePrefix
    result.SourceSuffix = SourceSuffix
    result.AnnotationGroup = AnnotationGroup
    result.Comment = Comment
    result.SourceReference = SourceReference
    Set Clone = result
End Function

' Копирует параметры в правило расстановки, чтобы дальнейшие изменения объекта
' не изменяли уже объявленное армирование.
Friend Sub CopyTo(ByRef rule As TRebarRule)
    rule.StartSetback = StartSetback: rule.EndSetback = EndSetback
    rule.DistributionScope = DistributionScope
    rule.EndMode = EndMode: rule.RemainderPlacement = RemainderPlacement
    rule.TangentOffset = TangentOffset: rule.NormalOffset = NormalOffset
    rule.ValidateRowOffset = ValidateRowOffset
    rule.EverySecond = EverySecond
    rule.IncomingTangentAtJoint = IncomingTangentAtJoint
    rule.DirectionX = DirectionX: rule.DirectionY = DirectionY
    rule.SourceNormalSign = SourceNormalSign
    rule.SuccessfulProjectionOrdinal = SuccessfulProjectionOrdinal: rule.NewPublicationRoot = NewPublicationRoot
    rule.SkipMissingProjection = SkipMissingProjection
    rule.Publication = Publication: rule.PlacementCheck = PlacementCheck
    rule.PositionOrder = PositionOrder: rule.Coincidence = Coincidence
    rule.OverlapPolicy = OverlapPolicy
    rule.SourcePrefix = SourcePrefix: rule.SourceSuffix = SourceSuffix
    rule.AnnotationGroup = AnnotationGroup: rule.Comment = Comment
    rule.SourceReference = SourceReference
End Sub
