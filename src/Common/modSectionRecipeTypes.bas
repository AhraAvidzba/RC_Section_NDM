Attribute VB_Name = "modSectionRecipeTypes"
Option Explicit

' Общие данные описания сечения во внутренних единицах.
' Хранят числа, имена и параметры правил без ссылок на Excel, решатель
' или классы конкретных форм сечения.
Public Enum ESectionLoopRole
    LoopOuter = 1
    LoopOpening = 2
End Enum

' Присваивает имена по общему номеру сегмента без отдельных счётчиков отрезков и дуг.
Public Enum EContourNamingPolicy
    PreserveSegmentIDs = 0
    SequentialByKind = 1
End Enum

' Компонента межцентрового смещения зависимого контактного ряда.
Public Enum ERebarOffsetAxis
    RowNormal = 1
    RowTangent = 2
End Enum

Public Enum ERebarEndMode
    IncludeBoth = 1
    ExcludeBoth = 2
    IncludeStart = 3
    IncludeEnd = 4
    ClosedUniform = 5
End Enum

Public Enum ERebarDistribution
    ByCount = 1
    ByMaxStep = 2
    ByFixedStep = 3
    ExplicitPositions = 4
    ProjectFrom = 5
    DependentOffset = 6
End Enum

' Количество/шаг применяется независимо к каждому непрерывному пути группы.
Public Enum ERebarDistributionScope
    PerPath = 1
End Enum

Public Enum ERemainderPlacement
    RemainderStart = 1
    RemainderEnd = 2
    RemainderCenter = 3
End Enum

Public Enum ERebarPublication
    PublishByRule = 1
    PublishPerSourcePosition = 2
End Enum

Public Enum ERebarPositionOrder
    AlongPath = 1
    FromPathMiddle = 2
End Enum

Public Enum ERebarCoincidence
    RejectCoincidence = 1
    MergeIdentical = 2
    KeepPublishedOwner = 3
End Enum

Public Enum ERebarPlacementCheck
    CheckCenter = 1
    CheckFullDisk = 2
End Enum

' Контакт допустим. Отрицательный геометрический зазор всегда диагностируется;
' запрет задается явно, чтобы не изменить исторические зависимые ряды.
Public Enum ERebarOverlapPolicy
    ReportOverlap = 1
    RejectOverlap = 2
End Enum

Public Type TSectionLoop
    PathID As String
    LoopID As String
    Role As ESectionLoopRole
    ParentID As String
    Comment As String
    ArcComment As String
    ContourIDs As Variant
    NamingPolicy As EContourNamingPolicy ' Способ присвоения имён отрезкам и дугам этой границы.
    NamingPrefix As String ' Независимый префикс границы при SequentialByKind.
End Type

' Явная связь основного ряда с бетонной границей и исходной фигурой.
' Имена R2/R3 передаются при объявлении; по суффиксам правил они не определяются.
Public Type TRebarContourLink
    MainRuleID As String ' Правило основного ряда, от которого независимо отсчитываются оба дополнительных ряда.
    LoopID As String ' Наружная граница или отверстие, к которому относится семейство.
    SourcePathID As String ' Исходная замкнутая фигура; у составного сечения может быть отдельным прямоугольником.
    AxisDistance As Double ' Заданный as основного ряда от исходной фигуры в сторону бетона, мм.
    NormalSign As Double ' Направление от Main дальше в бетон по его левой нормали: -1 или 1; 0 означает, что нужны другие однозначные метаданные.
    Row2RuleID As String ' Явное имя второго ряда; соответствующее правило может отсутствовать, если ряд выключен.
    Row3RuleID As String ' Явное имя третьего ряда; соответствующее правило может отсутствовать, если ряд выключен.
    SourceSegments As Variant ' Необязательный массив явных имён исходных участков, если геометрия допускает несколько соответствий.
    SourceReference As String ' Адреса исходных полей базового ограничения, которое используется даже без активного Main.
End Type

Public Type TRebarRule
    RuleID As String
    Target As String
    Mode As ERebarDistribution
    DistributionScope As ERebarDistributionScope
    Enabled As Boolean
    Diameter As Double
    SteelClass As String
    Count As Long
    Step As Double
    StartSetback As Double
    EndSetback As Double
    EndMode As ERebarEndMode
    RemainderPlacement As ERemainderPlacement
    SourceRuleID As String
    TangentOffset As Double
    NormalOffset As Double
    ValidateRowOffset As Boolean ' Проверять диски и криволинейный перенос дополнительных рядов семейства с новым отступом.
    EverySecond As Boolean
    IncomingTangentAtJoint As Boolean
    DirectionX As Double
    DirectionY As Double
    SourceNormalSign As Double
    SuccessfulProjectionOrdinal As Boolean
    NewPublicationRoot As Boolean
    SkipMissingProjection As Boolean
    Publication As ERebarPublication
    PositionOrder As ERebarPositionOrder
    Coincidence As ERebarCoincidence
    PlacementCheck As ERebarPlacementCheck
    OverlapPolicy As ERebarOverlapPolicy
    SourcePrefix As String
    SourceSuffix As String
    AnnotationGroup As String
    Comment As String
    SourceReference As String
    ExplicitStations As Variant
End Type

Public Type TSectionAnnotation
    AnnotationID As String
    Kind As String
    Target As String
    SegmentID As String
    StartAnchorID As String
    EndAnchorID As String
    Name As String
    StartX As Double
    StartY As Double
    EndX As Double
    EndY As Double
    NormalX As Double
    NormalY As Double
    DirectionX As Double
    DirectionY As Double
    Value As Double
    DisplayIncrement As Double
    ReverseDirection As Boolean
    Comment As String
End Type

Public Type TSectionMeshOptions
    StepX As Double ' Ширина основной ячейки бетонной сетки, мм.
    StepY As Double ' Высота основной ячейки бетонной сетки, мм.
    AutoStepFactor As Double ' 0 — использовать StepX/StepY; >0 — одинаковый шаг как доля меньшего габарита бетонной области.
    BoundarySubdivisions As Long ' Число подъячеек вдоль стороны граничной ячейки.
    MaterialID As Long ' Номер материала создаваемых бетонных элементов.
    SourceReference As String ' Сохранённые адреса параметров сетки для пояснения ошибок построения.
End Type

Public Type TRebarPosition
    X As Double
    Y As Double
    TangentX As Double
    TangentY As Double
    Station As Double
    RuleIndex As Long
    RootRuleIndex As Long
    RootOrdinal As Long
    RootPublicationRank As Long
    SourceOrdinal As Long
    SourcePositionIndex As Long ' Позиция непосредственного исходного ряда; 0 у независимого или спроецированного Main.
    IncomingTangentAtJoint As Boolean ' Какую касательную фактически использовали на стыке исходного пути.
    SourcePositionID As String
    PathID As String
End Type

' Выдаёт ошибку ввода, отделяя адреса, место ошибки, причину и действие пользователя.
Public Sub SectionRecipeError(ByVal owner As String, ByVal reason As String, Optional ByVal sourceReference As String = "", _
        Optional ByVal context As String = "", Optional ByVal entityCaption As String = "")
    Err.Raise vbObjectError + 5530, owner, SectionRecipeErrorText(reason, sourceReference, context, entityCaption)
End Sub

' Оформляет готовую причину без изменения адресов, чисел или расчётного смысла.
' Источник данных без ссылки на лист не называется адресом ячейки Excel.
Public Function SectionRecipeErrorText(ByVal reason As String, Optional ByVal sourceReference As String = "", _
        Optional ByVal context As String = "", Optional ByVal entityCaption As String = "") As String
    Dim body As String, action As String, position As Long, text As String
    body = Replace$(Replace$(reason, vbCrLf, vbLf), vbCr, vbLf)
    position = InStr(1, body, "Возможное решение:", vbBinaryCompare)
    If position > 0 Then
        action = TrimMessagePart(Mid$(body, position + Len("Возможное решение:")))
        body = TrimMessagePart(Left$(body, position - 1))
    Else
        position = FirstActionSentence(body)
        If position > 0 Then
            action = TrimMessagePart(Mid$(body, position))
            body = TrimMessagePart(Left$(body, position - 1))
        ElseIf Len(sourceReference) > 0 Then
            action = "Исправьте значение в указанной ячейке."
        End If
    End If
    If Len(sourceReference) > 0 Then
        If InStr(1, sourceReference, "!", vbBinaryCompare) > 0 Then text = "Адреса ячеек:" Else text = "Источник данных:"
        text = text & vbCrLf & sourceReference
    End If
    If Len(context) > 0 Then text = MessageParagraph(text, ContextParagraph(context))
    If Len(entityCaption) > 0 Then body = entityCaption & vbLf & body
    If Len(body) > 0 Then text = MessageParagraph(text, Replace$(body, vbLf, vbCrLf))
    If Len(action) > 0 Then text = MessageParagraph(text, "Возможное решение:" & vbCrLf & Replace$(action, vbLf, vbCrLf))
    SectionRecipeErrorText = text
End Function

' Находит первое действие в начале предложения; условие перед запятой не отрывается от него.
Private Function FirstActionSentence(ByVal text As String) As Long
    Dim verb As Variant, position As Long, prefix As String, lastCharacter As String
    For Each verb In Array("Введите ", "Выберите ", "Исправьте ", "Восстановите ", "Уменьшите ", "Увеличьте ", _
            "Измените ", "Проверьте ", "Укажите ", "Задайте ", "Удалите ", "Заполните ", "Повторите ", _
            "Сохраните ", "Закройте ", "Откройте ", "Обновите ", "Отключите ", "Перезапустите ")
        position = InStr(1, text, CStr(verb), vbTextCompare)
        Do While position > 0
            prefix = RTrim$(Left$(text, position - 1)): lastCharacter = Right$(prefix, 1)
            If position = 1 Or lastCharacter = "." Or lastCharacter = vbLf Then
                If FirstActionSentence = 0 Or position < FirstActionSentence Then FirstActionSentence = position
                Exit Do
            End If
            position = InStr(position + Len(CStr(verb)), text, CStr(verb), vbTextCompare)
        Loop
    Next verb
End Function

' Название сечения и описание стороны выводятся отдельными строками одного абзаца.
Private Function ContextParagraph(ByVal context As String) As String
    Dim position As Long, detail As String
    position = InStr(1, context, ": ", vbBinaryCompare)
    If position > 0 Then
        detail = Mid$(context, position + 2)
        If Len(detail) > 0 Then detail = UCase$(Left$(detail, 1)) & Mid$(detail, 2)
        ContextParagraph = Left$(context, position) & vbCrLf & detail
    Else
        ContextParagraph = context
    End If
End Function

' Добавляет смысловой абзац, оставляя между соседними абзацами пустую строку.
Private Function MessageParagraph(ByVal text As String, ByVal paragraph As String) As String
    MessageParagraph = paragraph
    If Len(text) > 0 Then MessageParagraph = text & vbCrLf & vbCrLf & paragraph
End Function

' Удаляет только пробелы и переводы строк по краям смыслового фрагмента.
Private Function TrimMessagePart(ByVal text As String) As String
    text = Trim$(text)
    Do While Len(text) > 0
        If Left$(text, 1) <> vbLf Then Exit Do
        text = Trim$(Mid$(text, 2))
    Loop
    Do While Len(text) > 0
        If Right$(text, 1) <> vbLf Then Exit Do
        text = Trim$(Left$(text, Len(text) - 1))
    Loop
    TrimMessagePart = text
End Function

' Добавляет название стороны и адреса исходных параметров к геометрической ошибке.
' Исчезнувший участок объясняет как невозможность разместить заданную арматуру;
' непредвиденные ошибки сохраняют свой исходный код и источник.
Public Sub SectionRecipeGeometryError(ByVal owner As String, ByVal sourceReference As String, _
        Optional ByVal rebarCaption As String = "", Optional ByVal path As CGeometryPath = Nothing)
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    If source = "CSectionGeometryQuery" Or source = "CGeometryPath" Then
        If Len(rebarCaption) > 0 Then
            Dim diagnostic As CGeometryDiagnostic
            If source = "CGeometryPath" And Not path Is Nothing Then Set diagnostic = path.LastDiagnostic
            If Not diagnostic Is Nothing Then
                If diagnostic.Code = "SEGMENT_DISAPPEARED" Then reason = _
                    "После скругления или смещения контура на выбранной стороне не осталось участка для размещения арматуры. " & _
                    "Измените размеры сечения, радиусы скругления или расстояние as от края бетона до осей стержней. " & _
                    "Если арматура на этой стороне не требуется, отключите её: задайте диаметр 0 либо оставьте диаметр пустым."
            End If
        End If
        SectionRecipeError owner, reason, sourceReference, rebarCaption
    End If
    Err.Raise number, source, reason
End Sub
