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

Public StartSetback As Double ' Начальное отсечение вдоль каждого target-пути, мм>=0; у полного ClosedUniform 0.
Public EndSetback As Double ' Конечное отсечение вдоль каждого target-пути, мм>=0; ProjectFrom фильтрует hit без перераспределения.
Public EndMode As ERebarEndMode ' IncludeBoth/ExcludeBoth/IncludeStart/IncludeEnd/ClosedUniform; последние требуют соответствующей замкнутости.
Public DistributionScope As ERebarDistributionScope ' Допустимо только PerPath: число стержней, расстояния между ними и отступы от концов применяются отдельно к каждой независимой линии группы.
Public RemainderPlacement As ERemainderPlacement ' RemainderStart/End/Center: положение остатка для точного шага.
Public TangentOffset As Double ' Локальное связанное смещение вдоль касательной, мм; signed, не независимый шаг ряда.
Public NormalOffset As Double ' Локальное связанное смещение по левой нормали (-ty,tx), мм; signed.
Public EverySecond As Boolean ' Брать локальные исходные позиции 1, 3, ... отдельно на каждом пути; SuccessfulProjectionOrdinal задаёт отдельный отсчёт успешных проекций.
Public IncomingTangentAtJoint As Boolean ' В точке соединения двух сегментов брать касательную предыдущего сегмента. По умолчанию берётся касательная следующего сегмента.
Public DirectionX As Double ' Глобальная X направления ProjectFrom, используется при SourceNormalSign=0.
Public DirectionY As Double ' Глобальная Y направления ProjectFrom; направление не нулевое.
Public SourceNormalSign As Double ' -1/0/+1: локальная нормаль исходного ряда или глобальное DirectionX/Y при 0.
Public SuccessfulProjectionOrdinal As Boolean ' True: заново нумеровать только стержни, для которых найдено пересечение с целевой линией, подряд по всей группе. False: сохранять исходные номера стержней с возможными пропусками отдельно на каждой линии.
Public NewPublicationRoot As Boolean ' У ProjectFrom начинает новый корень публикации по успешным позициям.
Public SkipMissingProjection As Boolean ' Разрешить пропуск без hit с PROJECTION_MISSING; False выдаёт ошибку.
Public Publication As ERebarPublication ' PublishByRule или PublishPerSourcePosition; влияет на вывод, не геометрический шаг.
Public PositionOrder As ERebarPositionOrder ' Порядок вывода стержней: AlongPath — от начала линии, FromPathMiddle — от её середины. При KeepPublishedOwner также определяет, чей диаметр и материал сохраняются в совпавшей позиции.
Public Coincidence As ERebarCoincidence ' RejectCoincidence/MergeIdentical/KeepPublishedOwner; последнее сохраняет явно выбранного опубликованного владельца.
Public PlacementCheck As ERebarPlacementCheck ' CheckCenter или CheckFullDisk относительно Outer/Opening; ужесточение задаётся явно.
Public OverlapPolicy As ERebarOverlapPolicy ' ReportOverlap или RejectOverlap. Контакт допустим; нормативный минимальный зазор не вводится.
Public SourcePrefix As String ' Префикс SourcePositionId; пустой использует RuleId+'.' либо связь source+suffix.
Public SourceSuffix As String ' Суффикс SourcePositionId, в том числе для связанных рядов.
Public AnnotationGroup As String ' Смысловая принадлежность подписи физического стержня; не расчётная группа трещин СП35.
Public Comment As String ' Инженерное пояснение правила; не selector или ResultCode.
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
' Меняет NormalOffset, TangentOffset и EverySecond; координаты рассчитывает CRebarRuleExecutor.
' Оба диаметра должны быть положительными; отключённый ряд пропускается до вызова.
Public Sub ConfigureDependentContact(ByVal baseDiameter As Double, ByVal diameter As Double, _
        ByVal axis As ERebarOffsetAxis, ByVal directionSign As Double, ByVal everySecond As Boolean, _
        Optional ByVal interveningDiameterSum As Double = 0#)
    If baseDiameter <= 0# Or diameter <= 0# Or interveningDiameterSum < 0# Then _
        SectionRecipeError "CRebarRuleOptions", "Для контактного ряда нужны положительные диаметры и неотрицательная сумма промежуточных диаметров.", SourceReference
    If axis <> RowNormal And axis <> RowTangent Then SectionRecipeError "CRebarRuleOptions", "Неизвестное направление контактного ряда.", SourceReference
    If directionSign <> -1# And directionSign <> 1# Then SectionRecipeError "CRebarRuleOptions", "Знак направления должен быть -1 или 1.", SourceReference
    Dim distance As Double
    distance = baseDiameter / 2# + interveningDiameterSum + diameter / 2#
    NormalOffset = 0#: TangentOffset = 0#
    If axis = RowNormal Then NormalOffset = directionSign * distance Else TangentOffset = directionSign * distance
    Me.EverySecond = everySecond
End Sub

' Копирует параметры в правило расстановки, чтобы дальнейшие изменения объекта
' не изменяли уже объявленное армирование.
Friend Sub CopyTo(ByRef rule As TRebarRule)
    rule.StartSetback = StartSetback: rule.EndSetback = EndSetback
    rule.DistributionScope = DistributionScope
    rule.EndMode = EndMode: rule.RemainderPlacement = RemainderPlacement
    rule.TangentOffset = TangentOffset: rule.NormalOffset = NormalOffset
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
