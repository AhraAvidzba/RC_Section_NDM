VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalculationProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Расчётный профиль сочетания нагрузок
' ==========================================================================
' Определяет, какие проверки выполнять и какие модели материалов использовать.
' Порядок расчётов и запись результатов задаёт CBatchSectionCalculator.
' ProfileId соответствует заголовку PR1/PR2/... в rngCalculationProfiles;
' профиль применяется, когда его имя указано в строке сочетания нагрузок.

Private mProfileId As String ' Стабильный ID столбца профиля, например PR1.
Private mDisplayName As String ' Человекочитаемое имя профиля.
Private mDescription As String ' Короткое описание для пользователя.
Private mInputLocations As Object ' Копии адресов полей профиля; используется только для пояснений, без ссылок на Excel.
Private mStrengthDirectState As Boolean ' Запрошено StrengthState.
Private mStrengthCapacity As Boolean ' Запрошена несущая способность.
Private mCrackWidth As Boolean ' Запрошен расчет раскрытия нормальных трещин.
Private mStabilityEnabled As Boolean ' Запрошена устойчивость; ее отказ не отменяет остальные проверки сочетания.
Private mStabilityValueSet As EMaterialValueSet ' Набор характеристик материала для устойчивости без диаграмм.
Private mHasStabilityValueSet As Boolean ' True, если строка MaterialModel.Stability.ValueSet задана в профиле.
Private mStrengthSpec As CMaterialModelSpec ' Материальная модель прочности.
Private mCrackInitiationSpec As CMaterialModelSpec ' Материальная модель проверки образования нормальной трещины.
Private mCrackedStateSpec As CMaterialModelSpec ' Материальная модель НДС с раскрытой трещиной.
Private mVisualizationState As ESectionStateType ' Какой StateType показывать на схеме и в AutoCAD.
Private mVisualizationQuantity As EVisualizationQuantity ' Stress или Strain для визуализации.
Private mHasVisualizationState As Boolean ' Визуальная настройка задана в профиле.
Private mHasVisualizationQuantity As Boolean ' Визуальная величина задана в профиле.
Private mVisualizationStressPrecision As Long ' Знаки после запятой для подписей Stress.
Private mVisualizationStrainPrecision As Long ' Знаки после запятой для подписей Strain.
Private mHasVisualizationStressPrecision As Boolean ' True, если точность Stress задана в профиле.
Private mHasVisualizationStrainPrecision As Boolean ' True, если точность Strain задана в профиле.

Private Const DEFAULT_VISUALIZATION_STRESS_PRECISION As Long = 1 ' Число знаков после запятой для напряжений, если точность не задана.
Private Const DEFAULT_VISUALIZATION_STRAIN_PRECISION As Long = 6 ' Число знаков после запятой для деформаций, если точность не задана.

' Создаёт пустые наборы параметров материалов для трёх назначений: прочности,
' образования трещины и НДС с трещинами.
Private Sub Class_Initialize()
    Set mStrengthSpec = New CMaterialModelSpec
    Set mCrackInitiationSpec = New CMaterialModelSpec
    Set mCrackedStateSpec = New CMaterialModelSpec
End Sub

' Заполняет основные параметры профиля. Модели материалов передаются отдельно,
' поскольку для трёх видов расчёта у них могут быть разные настройки.
Public Sub Initialize(ByVal profileId As String, ByVal displayName As String, ByVal description As String, _
        ByVal strengthDirectState As Boolean, ByVal strengthCapacity As Boolean, _
        ByVal crackWidth As Boolean, Optional ByVal stabilityEnabled As Boolean = False)
    Set mInputLocations = Nothing
    mProfileId = Trim$(profileId)
    mDisplayName = Trim$(displayName)
    mDescription = Trim$(description)
    mStrengthDirectState = strengthDirectState
    mStrengthCapacity = strengthCapacity
    mCrackWidth = crackWidth
    mStabilityEnabled = stabilityEnabled
End Sub

' Сохраняет спецификацию модели прочности.
Public Sub SetStrengthSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mStrengthSpec = spec.Clone
End Sub

' Сохраняет спецификацию модели образования трещины для обязательной предварительной проверки.
Public Sub SetCrackInitiationSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mCrackInitiationSpec = spec.Clone
End Sub

' Сохраняет спецификацию модели НДС с раскрытой трещиной.
Public Sub SetCrackedStateSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mCrackedStateSpec = spec.Clone
End Sub

' Сохраняет набор характеристик для устойчивости. Диаграммы здесь не нужны:
' расчет D, Ncr и eta использует R/E, а не sigma-epsilon график.
Public Sub SetStabilityValueSet(ByVal valueSetText As String)
    If Len(Trim$(valueSetText)) = 0 Then Exit Sub
    mStabilityValueSet = MaterialValueSetFromText(valueSetText)
    mHasStabilityValueSet = True
End Sub

' Сохраняет настройки схемы. Отсутствие выбранного состояния не блокирует расчёт:
' ошибка возникает при попытке показать это состояние на схеме или экспортировать в AutoCAD.
Public Sub SetVisualization(ByVal stateText As String, ByVal quantityText As String, _
        Optional ByVal stressPrecisionText As String = vbNullString, _
        Optional ByVal strainPrecisionText As String = vbNullString)
    If Len(Trim$(stateText)) > 0 Then
        mVisualizationState = SectionStateTypeFromText(stateText)
        mHasVisualizationState = True
    End If
    If Len(Trim$(quantityText)) > 0 Then
        mVisualizationQuantity = VisualizationQuantityFromText(quantityText)
        mHasVisualizationQuantity = True
    End If
    If Len(Trim$(stressPrecisionText)) > 0 Then
        mVisualizationStressPrecision = ParseVisualizationPrecision(stressPrecisionText, _
            "Visualization.StressPrecision")
        mHasVisualizationStressPrecision = True
    End If
    If Len(Trim$(strainPrecisionText)) > 0 Then
        mVisualizationStrainPrecision = ParseVisualizationPrecision(strainPrecisionText, _
            "Visualization.StrainPrecision")
        mHasVisualizationStrainPrecision = True
    End If
End Sub

' Проверяет наличие имени и хотя бы одной запрошенной проверки.
' Профиль используется, только если назначен сочетанию нагрузок;
' само заполнение столбца Config его не включает.
Public Function IsConfigured() As Boolean
    IsConfigured = (Len(mDisplayName) > 0 And (mStrengthDirectState Or mStrengthCapacity Or mCrackWidth Or mStabilityEnabled))
End Function

' Проверяет профиль, назначенный сочетанию. Наличие состояния для Visualization.State
' проверяется при выводе схемы, поэтому его отсутствие не останавливает пакетный расчёт.
Public Function ValidateForCalculation(ByRef errorText As String) As Boolean
    ValidateForCalculation = False
    If Len(mProfileId) = 0 Then
        errorText = ValidationMessage("ProfileId", "Не задан номер профиля ProfileId.")
        Exit Function
    End If
    If Not IsConfigured Then
        errorText = ValidationMessage("Profile.DisplayName;Calculation.Strength.DirectState;Calculation.Strength.Capacity;Calculation.Crack.Width;Calculation.Stability.Enabled", _
            "Задайте имя профиля и включите хотя бы одну расчётную проверку.")
        Exit Function
    End If
    If (mStrengthDirectState Or mStrengthCapacity) And Not mStrengthSpec.IsComplete Then
        errorText = ValidationMessage("MaterialModel.Strength.ValueSet;MaterialModel.Strength.ConcreteDiagram;MaterialModel.Strength.SteelDiagram;MaterialModel.Strength.ConcreteTension", _
            "Не полностью задана модель материалов для прочности (MaterialModel.Strength). Заполните набор характеристик, диаграммы бетона и арматуры, способ учёта растяжения бетона.")
        Exit Function
    End If
    If mStabilityEnabled And Not mHasStabilityValueSet Then
        errorText = ValidationMessage("MaterialModel.Stability.ValueSet", "Для устойчивости задайте набор характеристик материалов MaterialModel.Stability.ValueSet.")
        Exit Function
    End If
    If mCrackWidth And Not mCrackedStateSpec.IsComplete Then
        errorText = ValidationMessage("MaterialModel.CrackedState.ValueSet;MaterialModel.CrackedState.ConcreteDiagram;MaterialModel.CrackedState.SteelDiagram;MaterialModel.CrackedState.ConcreteTension", _
            "Не полностью задана модель материалов для НДС с трещинами (MaterialModel.CrackedState). Заполните её обязательные параметры.")
        Exit Function
    End If
    If mCrackWidth And mCrackedStateSpec.ConcreteTension <> ctkIgnore Then
        errorText = ValidationMessage("MaterialModel.CrackedState.ConcreteTension", "Для НДС с трещинами выберите Ignore: растянутый бетон в этой модели не учитывается.")
        Exit Function
    End If
    If mCrackWidth And Not mCrackInitiationSpec.IsComplete Then
        errorText = ValidationMessage("MaterialModel.CrackInitiation.ValueSet;MaterialModel.CrackInitiation.ConcreteDiagram;MaterialModel.CrackInitiation.SteelDiagram;MaterialModel.CrackInitiation.ConcreteTension", _
            "Для расчёта ширины трещин задайте модель материалов стадии образования трещины (MaterialModel.CrackInitiation).")
        Exit Function
    End If
    If mCrackWidth And mCrackInitiationSpec.IsComplete And mCrackInitiationSpec.ConcreteTension <> ctkUseDiagram Then
        errorText = ValidationMessage("MaterialModel.CrackInitiation.ConcreteTension", "Для определения нагрузки образования трещины выберите UseDiagram: требуется растянутая ветвь диаграммы бетона.")
        Exit Function
    End If
    ValidateForCalculation = True
End Function

' Получает только строковые адреса из ридера профилей и делает независимую копию.
Friend Sub SetInputLocations(ByVal locations As Object)
    Dim key As Variant
    Set mInputLocations = CreateObject("Scripting.Dictionary")
    mInputLocations.CompareMode = vbTextCompare
    For Each key In locations.Keys
        mInputLocations(CStr(key)) = CStr(locations(key))
    Next key
End Sub

' Добавляет к нарушенному условию профиль и сохранённые адреса связанных полей.
' Допустимость профиля определяется выше; пояснение не участвует в выборе расчётных ветвей.
Private Function ValidationMessage(ByVal keys As String, ByVal reason As String) As String
    Dim key As Variant
    ValidationMessage = "Профиль " & mProfileId & ": " & reason
    If mInputLocations Is Nothing Then
        ValidationMessage = ValidationMessage & " Параметры: " & keys & "."
        Exit Function
    End If
    For Each key In Split(keys, ";")
        If mInputLocations.Exists(CStr(key)) Then
            ValidationMessage = ValidationMessage & " " & CStr(key) & ": " & mInputLocations(key) & "."
        Else
            ValidationMessage = ValidationMessage & " Параметр " & CStr(key) & ": восстановите его строку в таблице профилей."
        End If
    Next key
End Function

' Стабильный ID столбца профиля, например PR1.
Public Property Get ProfileId() As String
    ProfileId = mProfileId
End Property

' Человекочитаемое имя профиля.
Public Property Get DisplayName() As String
    DisplayName = mDisplayName
End Property

' Короткое описание для пользователя.
Public Property Get Description() As String
    Description = mDescription
End Property

' Показывает, запрошено ли НДС по модели прочности при заданной нагрузке.
Public Property Get StrengthDirectStateEnabled() As Boolean
    StrengthDirectStateEnabled = mStrengthDirectState
End Property

' Запрошена несущая способность.
Public Property Get StrengthCapacityEnabled() As Boolean
    StrengthCapacityEnabled = mStrengthCapacity
End Property

' Запрошен расчет раскрытия нормальных трещин.
Public Property Get CrackWidthEnabled() As Boolean
    CrackWidthEnabled = mCrackWidth
End Property

' Запрошена устойчивость; ее отказ не отменяет остальные проверки сочетания.
Public Property Get StabilityEnabled() As Boolean
    StabilityEnabled = mStabilityEnabled
End Property

' Возвращает выбранный набор характеристик материалов для устойчивости.
' При программном создании профиля без явного выбора используется ULS.
' Обязательный ввод Config проверяется отдельно и этим значением не заменяется.
Public Property Get StabilityValueSet() As EMaterialValueSet
    If mHasStabilityValueSet Then
        StabilityValueSet = mStabilityValueSet
    Else
        StabilityValueSet = mvsULS
    End If
End Property

' Название набора нормативных параметров для устойчивости, выбранного профилем.
Public Property Get StabilityValueSetText() As String
    StabilityValueSetText = MaterialValueSetToText(StabilityValueSet)
End Property

' Материальная модель прочности.
Public Property Get StrengthSpec() As CMaterialModelSpec
    Set StrengthSpec = mStrengthSpec
End Property

' Материальная модель проверки образования нормальной трещины.
Public Property Get CrackInitiationSpec() As CMaterialModelSpec
    Set CrackInitiationSpec = mCrackInitiationSpec
End Property

' Материальная модель НДС с раскрытой трещиной.
Public Property Get CrackedStateSpec() As CMaterialModelSpec
    Set CrackedStateSpec = mCrackedStateSpec
End Property

' Выбирает сохраненное состояние для визуализации, не запуская его расчет.
' Если поле не задано через программный API, стандартный выбор - StrengthState.
Public Property Get VisualizationState() As ESectionStateType
    If mHasVisualizationState Then
        VisualizationState = mVisualizationState
    Else
        VisualizationState = sstStrengthState
    End If
End Property

' Выбранное профилем состояние НДС для схемы.
Public Property Get VisualizationStateText() As String
    VisualizationStateText = SectionStateTypeToText(VisualizationState)
End Property

' Возвращает визуализируемую величину. Для программно созданного профиля
' без явного выбора используется напряжение; проверка Config выполняется отдельно.
Public Property Get VisualizationQuantity() As EVisualizationQuantity
    If mHasVisualizationQuantity Then
        VisualizationQuantity = mVisualizationQuantity
    Else
        VisualizationQuantity = vqStress
    End If
End Property

' Выбранное профилем поле схемы: напряжения или деформации.
Public Property Get VisualizationQuantityText() As String
    VisualizationQuantityText = VisualizationQuantityToText(VisualizationQuantity)
End Property

' Возвращает число знаков после запятой для подписей напряжений.
' Численные результаты не округляет; при программном создании без явного
' выбора использует значение по умолчанию.
Public Property Get VisualizationStressPrecision() As Long
    If mHasVisualizationStressPrecision Then
        VisualizationStressPrecision = mVisualizationStressPrecision
    Else
        VisualizationStressPrecision = DEFAULT_VISUALIZATION_STRESS_PRECISION
    End If
End Property

' Возвращает число знаков после запятой для подписей деформаций.
' При программном создании без явного выбора использует значение по умолчанию;
' расчётное НДС не округляется.
Public Property Get VisualizationStrainPrecision() As Long
    If mHasVisualizationStrainPrecision Then
        VisualizationStrainPrecision = mVisualizationStrainPrecision
    Else
        VisualizationStrainPrecision = DEFAULT_VISUALIZATION_STRAIN_PRECISION
    End If
End Property

' Возвращает точность числовых подписей для выбранной визуализируемой величины.
Public Function VisualizationPrecisionForQuantity(ByVal quantity As EVisualizationQuantity) As Long
    If quantity = vqStrain Then
        VisualizationPrecisionForQuantity = VisualizationStrainPrecision
    Else
        VisualizationPrecisionForQuantity = VisualizationStressPrecision
    End If
End Function

' Читает пользовательское количество знаков после запятой для Stress/Strain.
' Ограничение защищает Excel-подписи от чрезмерно длинных строк.
Private Function ParseVisualizationPrecision(ByVal valueText As String, ByVal settingName As String) As Long
    If Not IsNumeric(valueText) Then Err.Raise vbObjectError + 3991, "CCalculationProfile", _
        settingName & " должен быть целым числом от 0 до 10."

    Dim numericValue As Double
    numericValue = CDbl(valueText)
    If numericValue < 0# Or numericValue > 10# Then _
        Err.Raise vbObjectError + 3992, "CCalculationProfile", _
            settingName & " должен быть целым числом от 0 до 10."
    If numericValue <> Fix(numericValue) Then _
        Err.Raise vbObjectError + 3992, "CCalculationProfile", _
            settingName & " должен быть целым числом от 0 до 10."
    ParseVisualizationPrecision = CLng(numericValue)
End Function
