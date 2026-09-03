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
' профилям: добавление PR5 в будущем не должно менять расчетное ядро.
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
' скобках пропускаются, а профильные столбцы остаются доступными по ProfileId.
Public Sub LoadFromRange(ByVal source As Object)
    If source Is Nothing Then Err.Raise vbObjectError + 3981, "CCalculationProfileCatalog", "Диапазон rngCalculationProfiles не найден."
    Clear

    Dim data As Variant
    data = source.Value2

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
    mCount = profileCount

    Dim profileIndex As Long
    For profileIndex = 1 To profileCount
        mProfileIds(profileIndex) = profileIds(profileIndex)
        Set mProfiles(profileIndex) = BuildProfile(data, profileColumns(profileIndex), profileIds(profileIndex))
    Next profileIndex
End Sub

' Создает профиль по одному столбцу PR.
Private Function BuildProfile(ByRef data As Variant, ByVal profileColumn As Long, ByVal profileId As String) As CCalculationProfile
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = 1

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        Dim keyText As String
        keyText = Trim$(CStr(data(rowIndex, 2)))
        If Len(keyText) > 0 Then values(keyText) = Trim$(CStr(data(rowIndex, profileColumn)))
    Next rowIndex

    Dim profile As CCalculationProfile
    Set profile = New CCalculationProfile
    profile.Initialize profileId, ProfileValue(values, "Profile.DisplayName"), _
        ProfileValue(values, "Profile.Description"), _
        YesNo(ProfileValue(values, "Calculation.Strength.DirectState")), _
        YesNo(ProfileValue(values, "Calculation.Strength.Capacity")), _
        YesNo(ProfileValue(values, "Calculation.Crack.Width")), _
        YesNo(ProfileValue(values, "Calculation.Stability.Enabled"))

    Dim spec As CMaterialModelSpec
    Set spec = SpecFromValues(values, "MaterialModel.Strength")
    If spec.IsComplete Then profile.SetStrengthSpec spec

    Set spec = SpecFromValues(values, "MaterialModel.CrackInitiation")
    If spec.IsComplete Then profile.SetCrackInitiationSpec spec

    Set spec = SpecFromValues(values, "MaterialModel.CrackedState")
    If spec.IsComplete Then profile.SetCrackedStateSpec spec

    profile.SetStabilityValueSet ProfileValue(values, "MaterialModel.Stability.ValueSet")

    profile.SetVisualization ProfileValue(values, "Visualization.State"), _
        ProfileValue(values, "Visualization.Quantity")

    Set BuildProfile = profile
End Function

' Собирает MaterialModelSpec из четырех строк одного префикса.
Private Function SpecFromValues(ByVal values As Object, ByVal prefix As String) As CMaterialModelSpec
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
    spec.Initialize valueSetText, concreteDiagramText, concreteTensionText, steelDiagramText
    Set SpecFromValues = spec
End Function

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

Private Function IsProfileIdText(ByVal value As String) As Boolean
    value = UCase$(Trim$(value))
    IsProfileIdText = (Len(value) >= 3 And Left$(value, 2) = "PR" And IsNumeric(Mid$(value, 3)))
End Function

Private Function YesNo(ByVal value As String) As Boolean
    value = UCase$(Trim$(value))
    YesNo = (value = "YES" Or value = "TRUE" Or value = "1" Or value = "ДА")
End Function

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

Private Function FindProfile(ByVal profileId As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mProfileIds(i), Trim$(profileId), vbTextCompare) = 0 Then
            FindProfile = i
            Exit Function
        End If
    Next i
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get ProfileId(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3985, "CCalculationProfileCatalog", "Индекс профиля вне диапазона."
    ProfileId = mProfileIds(index)
End Property

Public Property Get ProfileAt(ByVal index As Long) As CCalculationProfile
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3986, "CCalculationProfileCatalog", "Индекс профиля вне диапазона."
    Set ProfileAt = mProfiles(index)
End Property
