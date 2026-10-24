VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CCalculationSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Полный набор готовых таблиц одного расчёта без решателя и объектов Excel
' ==========================================================================
' Содержит прежние технические таблицы и недостающие строки подробных сводок.
' Числа сохраняются без округления. После подготовки набор только читается;
' восстановление не обращается к модели сечения, материалам или AutoCAD.
Private mTables As Object ' Матрицы значений по стабильным ключам таблиц и шапок.
Private mTableRows As Object ' Высоты неизменяемых матриц, без копирования таблиц при проверке раскладки.
Private mElapsedSeconds As Double ' Полное время запуска, отдельное от содержимого таблиц.
Private mRunID As Double ' Точная дата одного согласованного запуска.
Private mValidated As Boolean ' Повторная проверка того же неизменённого набора не копирует большие массивы.
Private mElements As Variant ' Большая матрица принадлежит снимку; служебный словарь содержит только её ключ.
Private mElementFirst() As Long, mElementLast() As Long ' Непрерывные группы строк в исходном порядке.
Private mElementCase() As String, mElementState() As String ' Точные подписи группы; поиск не зависит от регистра.
Private mElementGroupCount As Long ' Число групп; ноль также означает ещё не построенный индекс.

' Создаёт отдельный набор, не разделяющий состояние с другими расчётами.
Private Sub Class_Initialize()
    Set mTables = CreateObject("Scripting.Dictionary")
    mTables.CompareMode = vbBinaryCompare
    Set mTableRows = CreateObject("Scripting.Dictionary")
    mTableRows.CompareMode = vbBinaryCompare
End Sub

' Устанавливает дату до заполнения таблиц; преобразование в текст не требуется.
Public Property Let RunID(ByVal value As Double)
    mRunID = value
    mValidated = False
End Property

' Возвращает исходную дату для проверки согласованности нижних таблиц.
Public Property Get RunID() As Double
    RunID = mRunID
End Property

' Сохраняет измеренное время, не меняя численные результаты.
Public Property Let ElapsedSeconds(ByVal value As Double)
    mElapsedSeconds = value
End Property

' Возвращает длительность исходного расчёта при повторном открытии снимка.
Public Property Get ElapsedSeconds() As Double
    ElapsedSeconds = mElapsedSeconds
End Property

' Добавляет одну подготовленную матрицу. Повтор ключа означает ошибку сборки.
Public Sub AddTable(ByVal key As String, ByVal values As Variant)
    If Not IsArray(values) Then Err.Raise 5, "CCalculationSnapshot", "Не задана матрица снимка: " & key
    If mTables.Exists(key) Then Err.Raise 5, "CCalculationSnapshot", "Повторная таблица снимка: " & key
    If LBound(values, 1) <> 1 Or LBound(values, 2) <> 1 Then Err.Raise 5, "CCalculationSnapshot", "Неверные границы таблицы: " & key
    If Left$(key, 6) = "rngNDM" Or Left$(key, 9) = "LoadCase." Then TrimTechnicalRows values
    If key = "rngNDMElementResults" Then
        mElements = values
        mTables.Add key, Empty
        mElementGroupCount = 0
    Else
        mTables.Add key, values
    End If
    mTableRows.Add key, UBound(values, 1)
    mValidated = False
End Sub

' Удаляет неиспользованный резерв массивов, как блочная запись прежнего writer.
' Пустые строки пользовательских верхних таблиц сохраняются отдельным контрактом.
Private Sub TrimTechnicalRows(ByRef values As Variant)
    Dim rows As Long: rows = UBound(values, 1)
    Do While rows > 1
        If Len(CStr(values(rows, 1))) > 0 Then Exit Do
        rows = rows - 1
    Loop
    If rows = UBound(values, 1) Then Exit Sub
    Dim trimmed() As Variant, i As Long, j As Long
    ReDim trimmed(1 To rows, 1 To UBound(values, 2))
    For i = 1 To rows
        For j = 1 To UBound(values, 2): trimmed(i, j) = values(i, j): Next j
    Next i
    values = trimmed
End Sub

' Принимает большой готовый массив writer-а одной независимой копией.
' Незанятый резерв при несошедшемся состоянии не попадает в сохранённый снимок.
Friend Sub AddPreparedElementResults(ByRef values() As Variant, ByVal usedRows As Long)
    If mTables.Exists("rngNDMElementResults") Then Err.Raise 5, "CCalculationSnapshot", "Повторная таблица rngNDMElementResults."
    If LBound(values, 1) <> 1 Or LBound(values, 2) <> 1 Or UBound(values, 2) <> 9 Or usedRows < 1 Or usedRows > UBound(values, 1) Then Err.Raise 5, "CCalculationSnapshot", "Неверный размер готовых результатов элементов."
    If usedRows = UBound(values, 1) Then
        mElements = values
    Else
        Dim trimmed() As Variant, row As Long, column As Long
        ReDim trimmed(1 To usedRows, 1 To 9)
        For row = 1 To usedRows
            For column = 1 To 9: trimmed(row, column) = values(row, column): Next column
        Next row
        mElements = trimmed
    End If
    mTables.Add "rngNDMElementResults", Empty
    mTableRows.Add "rngNDMElementResults", usedRows
    mElementGroupCount = 0: mValidated = False
End Sub

' Возвращает настройки просмотра сохранённого профиля, отсутствующего в текущем Config.
' Расчётные модели по этим строкам не создаются: используются только готовые НДС.
Public Function SavedVisualizationProfile(ByVal profileID As String) As CCalculationProfile
    Dim values As Variant: values = Table("Calculation.Profiles")
    Dim column As Long, selected As Long, row As Long
    For row = 1 To UBound(values, 1)
        For column = 3 To UBound(values, 2)
            If StrComp(CStr(values(row, column)), profileID, vbTextCompare) = 0 Then selected = column: Exit For
        Next column
        If selected > 0 Then Exit For
    Next row
    If selected = 0 Then Err.Raise 5, , "В снимке отсутствует сохранённый профиль: " & profileID
    Dim fields As Object: Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For row = 1 To UBound(values, 1)
        If Len(CStr(values(row, 2))) > 0 Then fields.Item(CStr(values(row, 2))) = CStr(values(row, selected))
    Next row
    Dim profile As CCalculationProfile: Set profile = New CCalculationProfile
    profile.Initialize profileID, CStr(fields.Item("Profile.DisplayName")), CStr(fields.Item("Profile.Description")), _
        CStr(fields.Item("Calculation.Strength.DirectState")) = "Yes", CStr(fields.Item("Calculation.Strength.Capacity")) = "Yes", _
        CStr(fields.Item("Calculation.Crack.Width")) = "Yes", CStr(fields.Item("Calculation.Stability.Enabled")) = "Yes"
    profile.SetVisualization CStr(fields.Item("Visualization.State")), CStr(fields.Item("Visualization.Quantity")), _
        CStr(fields.Item("Visualization.StressPrecision")), CStr(fields.Item("Visualization.StrainPrecision"))
    Set SavedVisualizationProfile = profile
End Function

' Получает матрицу без повторного расчёта и чтения пользовательского ввода.
Public Function Table(ByVal key As String) As Variant
    If Not mTables.Exists(key) Then Err.Raise 5, "CCalculationSnapshot", "В снимке отсутствует таблица: " & key
    If key = "rngNDMElementResults" Then Table = mElements: Exit Function
    Dim values As Variant: values = mTables.Item(key)
    If key = "Summary" And mElapsedSeconds > 0# Then values(4, 1) = mElapsedSeconds
    Table = values
End Function

' Передаёт большую таблицу элементов кодеку по ссылке, без независимой копии для записи.
' Кодек только читает массив; публичный Table по-прежнему возвращает независимые данные.
Friend Sub WritePreparedTable(ByVal key As String, ByVal store As CCalculationSnapshotStore, ByVal file As CBinarySnapshotFile)
    If key = "rngNDMElementResults" Then
        store.WritePreparedMatrix file, mElements, TableRowCount("rngNDMSectionGeometry") - 1
    Else
        Dim values As Variant: values = Table(key)
        store.WritePreparedMatrix file, values
    End If
End Sub

' Возвращает размер без копирования всей матрицы ради одного UBound.
Public Function TableRowCount(ByVal key As String) As Long
    If Not mTables.Exists(key) Then Err.Raise 5, "CCalculationSnapshot", "В снимке отсутствует таблица: " & key
    TableRowCount = CLng(mTableRows.Item(key))
End Function

' Выбирает только нужные строки по кратковременному индексу неизменяемого снимка.
' Несмежные повторения сочетания и состояния сохраняют исходный порядок.
Public Function SelectElementResults(ByVal combinationID As String, ByVal stateType As String, ByVal indexOnly As Boolean) As Variant
    If Len(combinationID) = 0 And Len(stateType) = 0 And Not indexOnly Then SelectElementResults = Table("rngNDMElementResults"): Exit Function
    BuildElementIndex
    Dim group As Long, rows As Long, seen As Object, identity As String, include() As Boolean
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    ReDim include(0 To mElementGroupCount)
    For group = 1 To mElementGroupCount
        If (Len(combinationID) = 0 Or StrComp(mElementCase(group), combinationID, vbTextCompare) = 0) And _
                (Len(stateType) = 0 Or StrComp(mElementState(group), stateType, vbTextCompare) = 0) Then
            identity = CStr(Len(mElementCase(group))) & ":" & mElementCase(group) & mElementState(group)
            If Not indexOnly Or Not seen.Exists(identity) Then
                include(group) = True
                If indexOnly Then rows = rows + 1: seen.Item(identity) = True Else rows = rows + mElementLast(group) - mElementFirst(group) + 1
            End If
        End If
    Next group
    Dim selected() As Variant, row As Long, column As Long, used As Long, last As Long
    ReDim selected(1 To rows + 1, 1 To UBound(mElements, 2))
    For column = 1 To UBound(mElements, 2): selected(1, column) = mElements(1, column): Next column
    used = 1
    For group = 1 To mElementGroupCount
        If include(group) Then
            last = mElementLast(group): If indexOnly Then last = mElementFirst(group)
            For row = mElementFirst(group) To last
                used = used + 1
                For column = 1 To UBound(mElements, 2): selected(used, column) = mElements(row, column): Next column
            Next row
        End If
    Next group
    SelectElementResults = selected
End Function

' Один раз на объект снимка находит непрерывные группы; Excel и файлового ввода здесь нет.
Private Sub BuildElementIndex()
    If mElementGroupCount > 0 Or UBound(mElements, 1) < 2 Then Exit Sub
    On Error GoTo Failed
    Dim caseColumn As Long, stateColumn As Long, row As Long, capacity As Long
    caseColumn = ResultColumn(mElements, "LoadCase"): stateColumn = ResultColumn(mElements, "StateType")
    capacity = 32
    ReDim mElementFirst(1 To capacity): ReDim mElementLast(1 To capacity)
    ReDim mElementCase(1 To capacity): ReDim mElementState(1 To capacity)
    Dim loadCase As String, state As String, newGroup As Boolean
    For row = 2 To UBound(mElements, 1)
        loadCase = CStr(mElements(row, caseColumn)): state = CStr(mElements(row, stateColumn))
        newGroup = (mElementGroupCount = 0)
        If Not newGroup Then newGroup = StrComp(loadCase, mElementCase(mElementGroupCount), vbTextCompare) <> 0 Or StrComp(state, mElementState(mElementGroupCount), vbTextCompare) <> 0
        If newGroup Then
            mElementGroupCount = mElementGroupCount + 1
            If mElementGroupCount > capacity Then
                capacity = capacity * 2
                ReDim Preserve mElementFirst(1 To capacity): ReDim Preserve mElementLast(1 To capacity)
                ReDim Preserve mElementCase(1 To capacity): ReDim Preserve mElementState(1 To capacity)
            End If
            mElementFirst(mElementGroupCount) = row
            mElementCase(mElementGroupCount) = loadCase: mElementState(mElementGroupCount) = state
        End If
        mElementLast(mElementGroupCount) = row
    Next row
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    mElementGroupCount = 0
    Erase mElementFirst: Erase mElementLast: Erase mElementCase: Erase mElementState
    Err.Raise number, "CCalculationSnapshot.BuildElementIndex", reason
End Sub

' Проверяет наличие необязательной части до её чтения.
Public Function HasTable(ByVal key As String) As Boolean
    HasTable = mTables.Exists(key)
End Function

' Возвращает стабильный порядок частей для одного формата файла и листа.
Public Function Keys() As Variant
    Keys = mTables.Keys
End Function

' Перечисляет уже существующие технические таблицы; они не дублируются
' в единственном новом диапазоне дополнительных сведений на Results.
Public Function BaseKeys() As Variant
    BaseKeys = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", _
        "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
End Function

' Отличает существующие нижние таблицы от дополнительных матриц сводок.
Public Function IsBaseKey(ByVal key As String) As Boolean
    Dim item As Variant
    For Each item In BaseKeys()
        If CStr(item) = key Then IsBaseKey = True: Exit Function
    Next item
End Function

' Получает сохранённую единицу для точности отображения ширины трещин.
' Текущие единицы Config не меняют ранее рассчитанные значения.
Public Function OutputLengthUnit() As String
    Dim values As Variant, i As Long: values = Table("rngNDMSectionProperties")
    For i = 2 To UBound(values, 1)
        If CStr(values(i, 2)) = "ALL" And CStr(values(i, 3)) = "Output.LengthUnit" Then
            OutputLengthUnit = CStr(values(i, 4)): Exit Function
        End If
    Next i
    Err.Raise 5, "CCalculationSnapshot", "В rngNDMSectionProperties нет единицы длины снимка."
End Function

' Проверяет полный обязательный состав до изменения листа Results.
Public Sub Validate()
    If mValidated Then Exit Sub
    Dim key As Variant, values As Variant, i As Long
    For Each key In Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", _
            "Summary", "Strength.Headers", "Strength.Data", "Cracks.Headers", "Cracks.Data", "Stability.Labels", "Stability.Data", "LoadCase.Properties", "LoadCase.Annotations", "Calculation.Profiles")
        If CStr(key) = "rngNDMElementResults" Then
            If Not HasTable(CStr(key)) Then Err.Raise 5, "CCalculationSnapshot", "Отсутствует rngNDMElementResults."
            ValidateTableShape CStr(key), mElements
        Else
            values = Table(CStr(key)): ValidateTableShape CStr(key), values
        End If
    Next key
    For Each key In BaseKeys()
        If CStr(key) = "rngNDMElementResults" Then
            ValidateRunID mElements, CStr(key)
        Else
            values = Table(CStr(key)): ValidateRunID values, CStr(key)
        End If
    Next key
    For Each key In Array("LoadCase.Properties", "LoadCase.Annotations")
        values = Table(CStr(key))
        For i = 2 To UBound(values, 1)
            If CDbl(values(i, 1)) <> mRunID Then Err.Raise 5, "CCalculationSnapshot", "LC относятся к другому запуску: " & CStr(key)
        Next i
    Next key
    mValidated = True
End Sub

' Проверяет запуск непосредственно в принадлежащем снимку массиве, без копии для обхода.
Private Sub ValidateRunID(ByRef values As Variant, ByVal key As String)
    Dim row As Long
    For row = 2 To UBound(values, 1)
        If Len(CStr(values(row, 1))) > 0 Then
            If CDbl(values(row, 1)) <> mRunID Then Err.Raise 5, "CCalculationSnapshot", "Таблицы относятся к разным запускам: " & key
        End If
    Next row
End Sub

' Проверяет фиксированный контракт матриц до любой записи в Results.
Private Sub ValidateTableShape(ByVal key As String, ByRef values As Variant)
    Dim columns As Long, rows As Long
    Select Case key
        Case "rngNDMElementResults": columns = 9
        Case "rngNDMSectionGeometry": columns = 16
        Case "rngNDMSectionContours": columns = 17
        Case "rngNDMSectionProperties", "LoadCase.Properties": columns = 6
        Case "rngNDMMaterialDiagrams": columns = 11
        Case "rngNDMSectionAnnotations", "LoadCase.Annotations": columns = 24
        Case "Summary": columns = 26
        Case "Strength.Headers": columns = 49: rows = 4
        Case "Strength.Data": columns = 49
        Case "Cracks.Headers": columns = 72: rows = 4
        Case "Cracks.Data": columns = 72
        Case "Stability.Labels": columns = 84: rows = 1
        Case "Stability.Data": columns = 84
        Case "Calculation.Profiles"
            If UBound(values, 2) < 3 Or UBound(values, 1) < 1 Then Err.Raise 5, "CCalculationSnapshot", "Повреждены сохранённые профили расчёта."
            Exit Sub
    End Select
    If UBound(values, 2) <> columns Or UBound(values, 1) < 1 Then Err.Raise 5, "CCalculationSnapshot", "Неверная структура таблицы снимка: " & key
    If rows > 0 Then If UBound(values, 1) <> rows Then Err.Raise 5, "CCalculationSnapshot", "Неверная шапка таблицы снимка: " & key
End Sub

' Восстанавливает прежний контракт технической таблицы для схемы и экспорта.
' Физически общие и LC-зависимые строки хранятся раздельно.
Public Function ResultTable(ByVal key As String) As Variant
    Dim extra As String
    If key = "rngNDMSectionProperties" Then extra = "LoadCase.Properties"
    If key = "rngNDMSectionAnnotations" Then extra = "LoadCase.Annotations"
    If Len(extra) = 0 Then ResultTable = Table(key): Exit Function
    Dim common As Variant, cases As Variant, joined() As Variant, i As Long, j As Long
    common = Table(key): cases = Table(extra)
    ReDim joined(1 To UBound(common, 1) + UBound(cases, 1) - 1, 1 To UBound(common, 2))
    For i = 1 To UBound(common, 1)
        For j = 1 To UBound(common, 2): joined(i, j) = common(i, j): Next j
    Next i
    For i = 2 To UBound(cases, 1)
        For j = 1 To UBound(cases, 2): joined(UBound(common, 1) + i - 1, j) = cases(i, j): Next j
    Next i
    ResultTable = joined
End Function

' Разделяет готовые свойства или аннотации, не вычисляя их заново.
Public Sub AddSeparatedTable(ByVal key As String, ByVal values As Variant)
    Dim extra As String, i As Long, j As Long, commonCount As Long, caseCount As Long, isCase As Boolean
    If key = "rngNDMSectionProperties" Then extra = "LoadCase.Properties" Else extra = "LoadCase.Annotations"
    commonCount = 1: caseCount = 1
    For i = 2 To UBound(values, 1)
        If key = "rngNDMSectionProperties" Then isCase = IsCaseProperty(values, i) Else isCase = Left$(CStr(values(i, 2)), 6) = "CRACK_"
        If isCase Then caseCount = caseCount + 1 Else commonCount = commonCount + 1
    Next i
    Dim common() As Variant, cases() As Variant, c As Long, k As Long
    ReDim common(1 To commonCount, 1 To UBound(values, 2)): ReDim cases(1 To caseCount, 1 To UBound(values, 2))
    For j = 1 To UBound(values, 2): common(1, j) = values(1, j): cases(1, j) = values(1, j): Next j
    c = 1: k = 1
    For i = 2 To UBound(values, 1)
        If key = "rngNDMSectionProperties" Then isCase = IsCaseProperty(values, i) Else isCase = Left$(CStr(values(i, 2)), 6) = "CRACK_"
        If isCase Then
            k = k + 1: For j = 1 To UBound(values, 2): cases(k, j) = values(i, j): Next j
        Else
            c = c + 1: For j = 1 To UBound(values, 2): common(c, j) = values(i, j): Next j
        End If
    Next i
    AddTable key, common: AddTable extra, cases
End Sub

' Точка приложения и расчётная гибкость принадлежат запуску, а не геометрии сечения.
Private Function IsCaseProperty(ByRef values As Variant, ByVal row As Long) As Boolean
    Dim key As String: key = CStr(values(row, 3))
    IsCaseProperty = CStr(values(row, 2)) <> "ALL" Or key = "LoadReferenceX" Or key = "LoadReferenceY" Or Left$(key, Len("SP35.Slenderness")) = "SP35.Slenderness"
End Function
