VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CResultReadSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Источник одного просмотра: готовый полный снимок либо адресный файловый reader
' ==========================================================================
' Не является CCalculationSnapshot и не выдаёт частичный набор за полный расчёт.
' v4/v5 читаются полностью по явному контракту совместимости; v6 читает используемые блоки.
Private mPrepared As CCalculationSnapshot, mIndexed As CIndexedSnapshotFile
Private mCommonValidated As Boolean, mValidatedBoundaries As Object

' Подключает уже подготовленный полный набор без повторного чтения носителя.
Public Sub UsePrepared(ByVal snapshot As CCalculationSnapshot)
    mCommonValidated = False: Set mValidatedBoundaries = Nothing
    snapshot.Validate
    Set mPrepared = snapshot: Set mIndexed = Nothing
End Sub

' Открывает текущий файл Config; смена Excel/File не использует прежнюю привязку.
Public Sub OpenWorkbook(ByVal workbook As Object, ByVal store As CCalculationSnapshotStore)
    Set mPrepared = Nothing: Set mIndexed = Nothing: Set mValidatedBoundaries = Nothing: mCommonValidated = False
    Dim path As String: path = store.SelectedReadPath(workbook)
    If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then
        Set mPrepared = store.LoadFile(workbook) ' Штатная ошибка содержит фактическую настройку и ожидаемый путь.
        Exit Sub
    End If
    Set mIndexed = New CIndexedSnapshotFile
    If Not mIndexed.OpenRead(path, store) Then Set mIndexed = Nothing: Set mPrepared = store.LoadFile(workbook)
End Sub

' Возвращает только запрошенную таблицу, проверяя её контракт и RunID до использования.
Public Function Table(ByVal key As String) As Variant
    If Not mPrepared Is Nothing Then Table = mPrepared.Table(key): Exit Function
    Dim values As Variant: values = mIndexed.Table(key)
    If key <> "Snapshot.Kind" Then ValidateSnapshotTableShape key, values
    ValidateSnapshotRunID values, key, RunID
    Table = values
End Function

' Соединяет небольшие общие и LC-зависимые свойства/аннотации, не читая состояния элементов.
Public Function ResultTable(ByVal key As String) As Variant
    If Not mPrepared Is Nothing Then ResultTable = mPrepared.ResultTable(key): Exit Function
    Dim extra As String
    If key = "rngNDMSectionProperties" Then extra = "LoadCase.Properties"
    If key = "rngNDMSectionAnnotations" Then extra = "LoadCase.Annotations"
    If Len(extra) = 0 Or IsGeometryPreview Then ResultTable = Table(key): Exit Function
    Dim common As Variant, cases As Variant: common = Table(key): cases = Table(extra)
    ResultTable = JoinSnapshotTables(common, cases)
End Function

' Метаданные доступности не требуют загрузки подробных состояний.
Public Function SelectElementResults(ByVal combinationID As String, ByVal stateType As String, ByVal indexOnly As Boolean, Optional ByVal profileID As String = "") As Variant
    If Not mPrepared Is Nothing Then
        SelectElementResults = mPrepared.SelectElementResults(combinationID, stateType, indexOnly, profileID)
    Else
        SelectElementResults = mIndexed.SelectElementResults(combinationID, stateType, indexOnly, profileID)
    End If
End Function

' Проверяет все общие таблицы перед их публикацией на Results; состояния остаются адресными.
' Явный GeometryPreview представляет полный геометрический набор без НДС и проверяется целиком.
Public Sub ValidateForDisplay()
    If mCommonValidated Then Exit Sub
    If Not mPrepared Is Nothing Then mPrepared.Validate: mCommonValidated = True: Exit Sub
    If IsGeometryPreview Then
        Set mPrepared = mIndexed.LoadAll: Set mIndexed = Nothing
        mCommonValidated = True: Exit Sub
    End If
    Dim key As Variant, values As Variant, stabilityLabelColumns As Long
    For Each key In Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", _
            "Summary", "Strength.Headers", "Strength.Data", "Cracks.Headers", "Cracks.Data", "Stability.Labels", "Stability.Data", "LoadCase.Properties", "LoadCase.Annotations", "Calculation.Profiles")
        values = Table(CStr(key))
        ValidateStabilitySnapshotColumns CStr(key), values, stabilityLabelColumns
    Next key
    values = TableHeader("rngNDMElementResults"): ValidateSnapshotTableShape "rngNDMElementResults", values
    Dim geometry As Variant, boundaries As Variant
    geometry = Table("rngNDMSectionGeometry")
    If HasTable("rngNDMElementGeometry") Then boundaries = Table("rngNDMElementGeometry")
    Set mValidatedBoundaries = ReadConcreteBoundaryMap(boundaries, geometry)
    mCommonValidated = True
End Sub

' Повторно использует проверенные границы только при точном совпадении исходной геометрии.
Friend Function TryCopyValidatedConcreteBoundaries(ByRef geometry As Variant, ByRef result As Object) As Boolean
    If Not mPrepared Is Nothing Then TryCopyValidatedConcreteBoundaries = mPrepared.TryCopyValidatedConcreteBoundaries(geometry, result): Exit Function
    Dim expected As Variant: expected = Table("rngNDMSectionGeometry")
    If Not SnapshotMatricesEqual(geometry, expected) Then Exit Function
    If mValidatedBoundaries Is Nothing Then
        Dim boundaries As Variant
        If HasTable("rngNDMElementGeometry") Then boundaries = Table("rngNDMElementGeometry")
        Set mValidatedBoundaries = ReadConcreteBoundaryMap(boundaries, expected)
    End If
    Dim copied As Object: Set copied = CreateObject("Scripting.Dictionary"): copied.CompareMode = vbBinaryCompare
    Dim key As Variant, region As CConcreteRegion
    For Each key In mValidatedBoundaries.Keys
        Set region = mValidatedBoundaries.Item(CStr(key)): copied.Add CStr(key), region.Clone
    Next key
    Set result = copied: TryCopyValidatedConcreteBoundaries = True
End Function

' Читает сохранённый профиль, даже если текущий Config больше не содержит его столбца.
Public Function SavedVisualizationProfile(ByVal profileID As String) As CCalculationProfile
    Dim values As Variant: values = Table("Calculation.Profiles")
    Set SavedVisualizationProfile = SnapshotVisualizationProfile(values, profileID)
End Function

' Возвращает шапку элементов из отдельного блока; остальные шапки берёт из общих таблиц.
Public Function TableHeader(ByVal key As String) As Variant
    If Not mPrepared Is Nothing Then TableHeader = mPrepared.TableHeader(key): Exit Function
    If key = "rngNDMElementResults" Then TableHeader = mIndexed.ElementHeader: Exit Function
    Dim values As Variant, header() As Variant, column As Long: values = Table(key)
    ReDim header(1 To 1, 1 To UBound(values, 2))
    For column = 1 To UBound(values, 2): header(1, column) = values(1, column): Next column
    TableHeader = header
End Function

' Высота таблицы берётся из каталога без её загрузки.
Public Function TableRowCount(ByVal key As String) As Long
    If Not mPrepared Is Nothing Then TableRowCount = mPrepared.TableRowCount(key) Else TableRowCount = mIndexed.TableRowCount(key)
End Function

' Наличие необязательных границ определяется метаданными текущего источника.
Public Function HasTable(ByVal key As String) As Boolean
    If Not mPrepared Is Nothing Then HasTable = mPrepared.HasTable(key) Else HasTable = mIndexed.HasTable(key)
End Function

' Перечисляет технические таблицы, которые реально присутствуют в снимке.
Public Function BaseKeys() As Variant
    Dim keys As Variant
    keys = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
    If HasTable("rngNDMElementGeometry") Then ReDim Preserve keys(0 To 6): keys(6) = "rngNDMElementGeometry"
    BaseKeys = keys
End Function

' Признак preview задаётся явной таблицей Kind, а не отсутствием результатов НДС.
Public Property Get IsGeometryPreview() As Boolean
    If Not mPrepared Is Nothing Then IsGeometryPreview = mPrepared.IsGeometryPreview: Exit Property
    If Not HasTable("Snapshot.Kind") Then Exit Property
    Dim values As Variant: values = Table("Snapshot.Kind")
    If UBound(values, 1) <> 1 Or UBound(values, 2) <> 1 Then Err.Raise 5, , "Повреждён обязательный вид снимка."
    If VarType(values(1, 1)) <> vbString Then Err.Raise 5, , "В Snapshot.Kind требуется строка GeometryPreview."
    If CStr(values(1, 1)) <> "GeometryPreview" Then Err.Raise 5, , "Неизвестный вид снимка: " & CStr(values(1, 1))
    IsGeometryPreview = True
End Property

' Исходная единица длины нужна сводке и не берётся из изменённого Config.
Public Function OutputLengthUnit() As String
    Dim values As Variant, row As Long: values = Table("rngNDMSectionProperties")
    For row = 2 To UBound(values, 1)
        If CStr(values(row, 2)) = "ALL" And CStr(values(row, 3)) = "Output.LengthUnit" Then OutputLengthUnit = CStr(values(row, 4)): Exit Function
    Next row
    Err.Raise 5, , "В снимке нет Output.LengthUnit."
End Function

' RunID относится к текущему открытому снимку.
Public Property Get RunID() As Double
    If Not mPrepared Is Nothing Then RunID = mPrepared.RunID Else RunID = mIndexed.RunID
End Property

' Возвращает время сохранённого расчёта.
Public Property Get ElapsedSeconds() As Double
    If Not mPrepared Is Nothing Then ElapsedSeconds = mPrepared.ElapsedSeconds Else ElapsedSeconds = mIndexed.ElapsedSeconds
End Property

' Полный снимок восстанавливается только при явном запросе полного контракта.
Public Function LoadAll() As CCalculationSnapshot
    If Not mPrepared Is Nothing Then Set LoadAll = mPrepared Else Set LoadAll = mIndexed.LoadAll
End Function
