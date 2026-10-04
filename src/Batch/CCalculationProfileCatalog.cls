VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalculationProfileCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' КАТАЛОГ РАСЧЕТНЫХ ПРОФИЛЕЙ
' ==========================================================================
' Каталог читает вертикальную таблицу rngCalculationProfiles: параметры идут
' по строкам, а профили PR1/PR2/... - по столбцам. Код не привязан к четырем
' профилям: число профилей определяется таблицей и не зашито в расчетное ядро.
'
' Класс не решает, какие проверки запускать и в каком порядке. Он только
' возвращает CCalculationProfile по ProfileId, назначенному сочетанию нагрузок.

Private mProfiles() As CCalculationProfile ' Прочитанные профили по столбцам таблицы.
Private mProfileIds() As String ' Стабильные ID профилей.
Private mCount As Long ' Количество столбцов профилей в таблице.

' Загружает каталог из именованного диапазона книги.
Public Sub LoadFromWorkbook(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 3980, "CCalculationProfileCatalog", "Книга Excel не передана."
    LoadFromRange workbook.Names.Item("rngCalculationProfiles").RefersToRange
End Sub

' Читает вертикальную таблицу профилей. Структурные строки в квадратных
' скобках пропускаются. Каталог доступен только после успешного чтения всех
' столбцов; при отказе частичные данные удаляются, исходная ошибка сохраняется.
Public Sub LoadFromRange(ByVal source As Object)
    If source Is Nothing Then Err.Raise vbObjectError + 3981, "CCalculationProfileCatalog", "Диапазон rngCalculationProfiles не найден."
    Clear
    On Error GoTo Failed

    Dim data As Variant
    data = source.Value2
    If Not IsArray(data) Then Err.Raise vbObjectError + 3987, "CCalculationProfileCatalog", _
        "Таблица расчетных профилей должна содержать заголовок, строки параметров и минимум один столбец профиля."
    If UBound(data, 1) < 2 Or UBound(data, 2) < 3 Then Err.Raise vbObjectError + 3987, "CCalculationProfileCatalog", _
        "Таблица расчетных профилей должна содержать заголовок, строки параметров и минимум один столбец профиля."

    Dim headerRow As Long
    headerRow = FindProfileHeaderRow(data)
    If headerRow = 0 Then Err.Raise vbObjectError + 3982, "CCalculationProfileCatalog", "В rngCalculationProfiles не найдена строка ProfileId."

    Dim profileColumns() As Long
    Dim profileIds() As String
    Dim profileCount As Long
    CollectProfileColumns data, headerRow, profileColumns, profileIds, profileCount
    If profileCount = 0 Then Err.Raise vbObjectError + 3983, "CCalculationProfileCatalog", "В rngCalculationProfiles нет столбцов PR*."

    ReDim mProfiles(1 To profileCount)
    ReDim mProfileIds(1 To profileCount)

    Dim profileIndex As Long
    For profileIndex = 1 To profileCount
        mProfileIds(profileIndex) = profileIds(profileIndex)
        Set mProfiles(profileIndex) = BuildProfile(source, data, profileColumns(profileIndex), profileIds(profileIndex))
    Next profileIndex
    mCount = profileCount
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, errorSource As String
    errorNumber = Err.Number: description = Err.Description: errorSource = Err.Source
    Clear
    Err.Raise errorNumber, errorSource, description
End Sub

' Создает профиль по одному столбцу PR и сохраняет адреса его параметров
' для понятной ошибки ввода. Адреса не передаются в расчетное ядро.
Private Function BuildProfile(ByVal source As Object, ByRef data As Variant, _
        ByVal profileColumn As Long, ByVal profileId As String) As CCalculationProfile
    On Error GoTo InvalidValue
    Dim values As Object
    Dim locations As Object
    Dim activeKey As String ' Параметр, который сейчас проверяется; пустой для уже адресной ошибки.
    Set values = CreateObject("Scripting.Dictionary")
    Set locations = CreateObject("Scripting.Dictionary")
    values.CompareMode = 1
    locations.CompareMode = 1

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        Dim keyText As String
        keyText = Trim$(CStr(data(rowIndex, 2)))
        If Len(keyText) > 0 Then
            locations(keyText) = source.Worksheet.Name & "!" & source.Cells(rowIndex, profileColumn).Address(False, False)
            values(keyText) = ProfileCellText(data(rowIndex, profileColumn), keyText, profileId, locations)
        End If
    Next rowIndex

    Dim profile As CCalculationProfile
    Set profile = New CCalculationProfile
    profile.Initialize profileId, ProfileValue(values, "Profile.DisplayName"), _
        ProfileValue(values, "Profile.Description"), _
        YesNo(ProfileValue(values, "Calculation.Strength.DirectState"), "Calculation.Strength.DirectState", profileId, locations), _
        YesNo(ProfileValue(values, "Calculation.Strength.Capacity"), "Calculation.Strength.Capacity", profileId, locations), _
        YesNo(ProfileValue(values, "Calculation.Crack.Width"), "Calculation.Crack.Width", profileId, locations), _
        YesNo(ProfileValue(values, "Calculation.Stability.Enabled"), "Calculation.Stability.Enabled", profileId, locations)

    Dim spec As CMaterialModelSpec
    Set spec = SpecFromValues(values, locations, profileId, "MaterialModel.Strength")
    If spec.IsComplete Then profile.SetStrengthSpec spec

    Set spec = SpecFromValues(values, locations, profileId, "MaterialModel.CrackInitiation")
    If spec.IsComplete Then profile.SetCrackInitiationSpec spec

    Set spec = SpecFromValues(values, locations, profileId, "MaterialModel.CrackedState")
    If spec.IsComplete Then profile.SetCrackedStateSpec spec

    activeKey = "MaterialModel.Stability.ValueSet"
    If profile.StabilityEnabled Then RequireProfileValue values, activeKey
    profile.SetStabilityValueSet ProfileValue(values, activeKey)

    activeKey = "Visualization.State"
    RequireProfileValue values, activeKey
    profile.SetVisualization ProfileValue(values, activeKey), vbNullString
    activeKey = "Visualization.Quantity"
    RequireProfileValue values, activeKey
    profile.SetVisualization vbNullString, ProfileValue(values, activeKey)
    activeKey = "Visualization.StressPrecision"
    RequireProfileValue values, activeKey
    profile.SetVisualization vbNullString, vbNullString, ProfileValue(values, activeKey)
    activeKey = "Visualization.StrainPrecision"
    RequireProfileValue values, activeKey
    profile.SetVisualization vbNullString, vbNullString, vbNullString, ProfileValue(values, activeKey)

    Set BuildProfile = profile
    Exit Function
InvalidValue:
    Dim errorNumber As Long, description As String
    errorNumber = Err.Number: description = Err.Description
    If Len(activeKey) = 0 Then Err.Raise errorNumber, "CCalculationProfileCatalog", description
    RaiseProfileInputError errorNumber, description, activeKey, profileId, locations
End Function

' Собирает MaterialModelSpec теми же enum-парсерами, что использует ядро.
' Каждая ошибка называет конкретное поле; полностью пустая неактивная модель
' остается ненастроенной и не получает произвольную диаграмму по умолчанию.
Private Function SpecFromValues(ByVal values As Object, ByVal locations As Object, _
        ByVal profileId As String, ByVal prefix As String) As CMaterialModelSpec
    On Error GoTo InvalidValue
    Dim activeKey As String
    Dim spec As CMaterialModelSpec
    Set spec = New CMaterialModelSpec

    Dim valueSetText As String
    Dim concreteDiagramText As String
    Dim concreteTensionText As String
    Dim steelDiagramText As String
    valueSetText = ProfileValue(values, prefix & ".ValueSet")
    concreteDiagramText = ProfileValue(values, prefix & ".ConcreteDiagram")
    concreteTensionText = ProfileValue(values, prefix & ".ConcreteTension")
    steelDiagramText = ProfileValue(values, prefix & ".SteelDiagram")

    If Len(valueSetText) = 0 And Len(concreteDiagramText) = 0 And _
            Len(concreteTensionText) = 0 And Len(steelDiagramText) = 0 Then
        Set SpecFromValues = spec
        Exit Function
    End If
    activeKey = prefix & ".ValueSet"
    Call MaterialValueSetFromText(valueSetText)
    activeKey = prefix & ".ConcreteDiagram"
    Call ConcreteDiagramKindFromText(concreteDiagramText)
    activeKey = prefix & ".ConcreteTension"
    Call ConcreteTensionKindFromText(concreteTensionText)
    activeKey = prefix & ".SteelDiagram"
    Call SteelDiagramKindFromText(steelDiagramText)
    spec.Initialize valueSetText, concreteDiagramText, concreteTensionText, steelDiagramText
    Set SpecFromValues = spec
    Exit Function
InvalidValue:
    RaiseProfileInputError Err.Number, Err.Description, activeKey, profileId, locations
End Function

' Различает строку, пустоту и ошибку Excel до преобразования в String.
' Ошибка формулы не становится неизвестным selector-ом или значением default.
Private Function ProfileCellText(ByVal value As Variant, ByVal key As String, _
        ByVal profileId As String, ByVal locations As Object) As String
    Dim errorNumber As Long
    errorNumber = vbObjectError + 3989
    If Left$(key, 12) = "Calculation." Then errorNumber = vbObjectError + 3988
    If IsError(value) Then RaiseProfileInputError errorNumber, "Ячейка содержит ошибку формулы Excel.", key, profileId, locations
    If IsNull(value) Then RaiseProfileInputError errorNumber, "Значение ячейки не определено.", key, profileId, locations
    ProfileCellText = Trim$(CStr(value))
End Function

' Проверяет обязательность поля из Config. Optional-параметры программного
' API профиля не разрешают молча заменить пустой пользовательский ввод.
Private Sub RequireProfileValue(ByVal values As Object, ByVal key As String)
    If Len(ProfileValue(values, key)) = 0 Then Err.Raise vbObjectError + 3989, _
        "CCalculationProfileCatalog", "Обязательное значение не задано; выберите или введите его в таблице профилей."
End Sub

' Добавляет адресную навигацию к исходной причине ввода, сохраняя машинный
' номер ошибки enum/параметра. Статус результата по этому тексту не назначается.
Private Sub RaiseProfileInputError(ByVal errorNumber As Long, ByVal description As String, _
        ByVal key As String, ByVal profileId As String, ByVal locations As Object)
    Dim location As String
    location = "лист Config, таблица rngCalculationProfiles"
    If Not locations Is Nothing Then
        If locations.Exists(key) Then location = CStr(locations(key))
    End If
    Err.Raise errorNumber, "CCalculationProfileCatalog", "Профиль " & profileId & _
        ": параметр " & key & " (" & location & "). " & description & _
        " Проверьте указанную ячейку или восстановите строку параметра в таблице профилей."
End Sub

' Возвращает строковое значение параметра профиля по ключу, если такая строка есть в таблице.
Private Function ProfileValue(ByVal values As Object, ByVal keyText As String) As String
    If values.Exists(keyText) Then ProfileValue = Trim$(CStr(values(keyText)))
End Function

' Находит строку, где после служебных колонок идут PR1/PR2/... .
Private Function FindProfileHeaderRow(ByRef data As Variant) As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        For colIndex = 3 To UBound(data, 2)
            If IsProfileIdText(CStr(data(rowIndex, colIndex))) Then
                FindProfileHeaderRow = rowIndex
                Exit Function
            End If
        Next colIndex
    Next rowIndex
End Function

' Собирает все столбцы PR*, чтобы reader не был жестко привязан к PR1-PR4.
Private Sub CollectProfileColumns(ByRef data As Variant, ByVal headerRow As Long, _
        ByRef profileColumns() As Long, ByRef profileIds() As String, ByRef profileCount As Long)
    Dim colIndex As Long
    For colIndex = 3 To UBound(data, 2)
        Dim profileId As String
        profileId = Trim$(CStr(data(headerRow, colIndex)))
        If IsProfileIdText(profileId) Then
            profileCount = profileCount + 1
            ReDim Preserve profileColumns(1 To profileCount)
            ReDim Preserve profileIds(1 To profileCount)
            profileColumns(profileCount) = colIndex
            profileIds(profileCount) = profileId
        End If
    Next colIndex
End Sub

' Проверяет, что заголовок столбца выглядит как расчетный ProfileId вида PR1, PR2 и т.п.
Private Function IsProfileIdText(ByVal value As String) As Boolean
    value = UCase$(Trim$(value))
    IsProfileIdText = (Len(value) >= 3 And Left$(value, 2) = "PR" And IsNumeric(Mid$(value, 3)))
End Function

' Читает обязательный переключатель расчета профиля. Все утвержденные
' Boolean aliases сохраняются; пустое/неизвестное значение является ошибкой
' Config, а не неявным отключением запрошенного пользователем расчета.
Private Function YesNo(ByVal value As String, ByVal key As String, ByVal profileId As String, _
        ByVal locations As Object) As Boolean
    value = UCase$(Trim$(value))
    Select Case value
        Case "YES", "TRUE", "1", "ДА": YesNo = True
        Case "NO", "FALSE", "0", "НЕТ": YesNo = False
        Case Else
            RaiseProfileInputError vbObjectError + 3988, _
                "Выберите Yes или No; прочитано: """ & value & """.", key, profileId, locations
    End Select
End Function

' Очищает ранее прочитанный каталог перед новой загрузкой из workbook.
Private Sub Clear()
    Erase mProfiles
    Erase mProfileIds
    mCount = 0
End Sub

' Возвращает профиль по стабильному ID столбца.
Public Function ProfileById(ByVal profileId As String) As CCalculationProfile
    Dim index As Long
    index = FindProfile(profileId)
    If index = 0 Then Err.Raise vbObjectError + 3984, "CCalculationProfileCatalog", _
        "Расчетный профиль " & profileId & " не найден в rngCalculationProfiles."
    Set ProfileById = mProfiles(index)
End Function

' Проверяет наличие профиля без выбрасывания ошибки.
Public Function HasProfile(ByVal profileId As String) As Boolean
    HasProfile = (FindProfile(profileId) > 0)
End Function

' Ищет профиль сначала по стабильному ID столбца, затем по пользовательскому имени профиля.
Private Function FindProfile(ByVal profileId As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mProfileIds(i), Trim$(profileId), vbTextCompare) = 0 Then
            FindProfile = i
            Exit Function
        End If
    Next i
    For i = 1 To mCount
        If Not mProfiles(i) Is Nothing Then
            If StrComp(mProfiles(i).DisplayName, Trim$(profileId), vbTextCompare) = 0 Then
                FindProfile = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

' Возвращает стабильный ID профиля по позиции каталога. Неверный индекс
' является нарушением контракта доступа, а не поводом выбрать другой профиль.
Public Property Get ProfileId(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3985, "CCalculationProfileCatalog", "Индекс профиля вне диапазона."
    ProfileId = mProfileIds(index)
End Property

' Возвращает профиль по проверенному индексу, сохраняя его идентичность.
' Здесь не создается запасной профиль и не меняются настройки материала.
Public Property Get ProfileAt(ByVal index As Long) As CCalculationProfile
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3986, "CCalculationProfileCatalog", "Индекс профиля вне диапазона."
    Set ProfileAt = mProfiles(index)
End Property
