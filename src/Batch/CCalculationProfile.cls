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
' РАСЧЕТНЫЙ ПРОФИЛЬ LC
' ==========================================================================
' Профиль отвечает только на вопрос "что считать" для сочетания: прямое НДС
' по прочности, несущую способность, раскрытие трещин и какие материальные
' модели использовать для этих расчетов. Порядок выполнения, переиспользование
' состояний и запись Results остаются ответственностью batch-orchestrator-а.
'
' ProfileId задается заголовком столбца PR1/PR2/... в rngCalculationProfiles.
' Отдельного IsActive нет: профиль используется только тогда, когда его
' ProfileId назначен строке сочетания.

Private mProfileId As String ' Стабильный ID столбца профиля, например PR1.
Private mDisplayName As String ' Человекочитаемое имя профиля.
Private mDescription As String ' Короткое описание для пользователя.
Private mStrengthDirectState As Boolean ' Запрошено StrengthState.
Private mStrengthCapacity As Boolean ' Запрошена несущая способность.
Private mCrackWidth As Boolean ' Запрошен расчет раскрытия нормальных трещин.
Private mStabilityEnabled As Boolean ' Запрошен фильтр продольного изгиба/устойчивости перед НДМ.
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

Private Const DEFAULT_VISUALIZATION_STRESS_PRECISION As Long = 1
Private Const DEFAULT_VISUALIZATION_STRAIN_PRECISION As Long = 6

' Создает пустые спецификации материалов для трех расчетных ролей профиля.
Private Sub Class_Initialize()
    Set mStrengthSpec = New CMaterialModelSpec
    Set mCrackInitiationSpec = New CMaterialModelSpec
    Set mCrackedStateSpec = New CMaterialModelSpec
End Sub

' Заполняет простые поля профиля. MaterialModelSpec задаются отдельными
' методами, потому что профиль хранит три независимые роли материальных моделей.
Public Sub Initialize(ByVal profileId As String, ByVal displayName As String, ByVal description As String, _
        ByVal strengthDirectState As Boolean, ByVal strengthCapacity As Boolean, _
        ByVal crackWidth As Boolean, Optional ByVal stabilityEnabled As Boolean = False)
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

' Сохраняет настройки визуализации. Их несоответствие фактически полученным
' состояниям не блокирует расчет: ошибка появится только при построении схемы
' или AutoCAD export по отсутствующему StateType.
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

' Профиль считается настроенным только если задано имя и есть хотя бы один
' запрошенный расчет. Неназначенные профили не участвуют в batch, даже если
' в таблице профилей заполнены все их параметры.
Public Function IsConfigured() As Boolean
    IsConfigured = (Len(mDisplayName) > 0 And (mStrengthDirectState Or mStrengthCapacity Or mCrackWidth Or mStabilityEnabled))
End Function

' Валидирует профиль только для фактически назначенного LC. Visualization.State
' здесь не проверяется на наличие будущего состояния: это ошибка визуализации,
' а не причина блокировать batch-расчет.
Public Function ValidateForCalculation(ByRef errorText As String) As Boolean
    ValidateForCalculation = False
    If Len(mProfileId) = 0 Then
        errorText = "ProfileId профиля не задан."
        Exit Function
    End If
    If Not IsConfigured Then
        errorText = "Профиль " & mProfileId & " не настроен: задайте имя и включите хотя бы один расчет."
        Exit Function
    End If
    If (mStrengthDirectState Or mStrengthCapacity) And Not mStrengthSpec.IsComplete Then
        errorText = "Профиль " & mProfileId & ": не задана MaterialModel.Strength."
        Exit Function
    End If
    If mStabilityEnabled And Not mHasStabilityValueSet Then
        errorText = "Профиль " & mProfileId & ": для устойчивости задайте MaterialModel.Stability.ValueSet."
        Exit Function
    End If
    If mCrackWidth And Not mCrackedStateSpec.IsComplete Then
        errorText = "Профиль " & mProfileId & ": не задана MaterialModel.CrackedState."
        Exit Function
    End If
    If mCrackWidth And mCrackedStateSpec.ConcreteTension <> ctkIgnore Then
        errorText = "Профиль " & mProfileId & ": MaterialModel.CrackedState.ConcreteTension должен быть Ignore."
        Exit Function
    End If
    If mCrackWidth And Not mCrackInitiationSpec.IsComplete Then
        errorText = "Профиль " & mProfileId & ": для Calculation.Crack.Width нужна MaterialModel.CrackInitiation."
        Exit Function
    End If
    If mCrackWidth And mCrackInitiationSpec.IsComplete And mCrackInitiationSpec.ConcreteTension <> ctkUseDiagram Then
        errorText = "Профиль " & mProfileId & ": MaterialModel.CrackInitiation.ConcreteTension должен быть UseDiagram."
        Exit Function
    End If
    ValidateForCalculation = True
End Function

Public Property Get ProfileId() As String
    ProfileId = mProfileId
End Property

Public Property Get DisplayName() As String
    DisplayName = mDisplayName
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Get StrengthDirectStateEnabled() As Boolean
    StrengthDirectStateEnabled = mStrengthDirectState
End Property

Public Property Get StrengthCapacityEnabled() As Boolean
    StrengthCapacityEnabled = mStrengthCapacity
End Property

Public Property Get CrackWidthEnabled() As Boolean
    CrackWidthEnabled = mCrackWidth
End Property

Public Property Get StabilityEnabled() As Boolean
    StabilityEnabled = mStabilityEnabled
End Property

' Возвращает заданный набор характеристик устойчивости. Для профиля, созданного
' программно без этого поля, используется ULS; чтение обязательного поля Config
' отдельно проверяется reader-ом и не подменяет невалидный ввод этим значением.
Public Property Get StabilityValueSet() As EMaterialValueSet
    If mHasStabilityValueSet Then
        StabilityValueSet = mStabilityValueSet
    Else
        StabilityValueSet = mvsULS
    End If
End Property

Public Property Get StabilityValueSetText() As String
    StabilityValueSetText = MaterialValueSetToText(StabilityValueSet)
End Property

Public Property Get StrengthSpec() As CMaterialModelSpec
    Set StrengthSpec = mStrengthSpec
End Property

Public Property Get CrackInitiationSpec() As CMaterialModelSpec
    Set CrackInitiationSpec = mCrackInitiationSpec
End Property

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

Public Property Get VisualizationQuantityText() As String
    VisualizationQuantityText = VisualizationQuantityToText(VisualizationQuantity)
End Property

' Возвращают точность подписей напряжений и деформаций, не меняя численные
' результаты. Для незаданных полей программного API используются именованные defaults.
Public Property Get VisualizationStressPrecision() As Long
    If mHasVisualizationStressPrecision Then
        VisualizationStressPrecision = mVisualizationStressPrecision
    Else
        VisualizationStressPrecision = DEFAULT_VISUALIZATION_STRESS_PRECISION
    End If
End Property

' Возвращает точность подписей деформаций из профиля либо программный default,
' если это поле не задано через API; расчетное НДС не округляет.
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

    ParseVisualizationPrecision = CLng(CDbl(valueText))
    If ParseVisualizationPrecision < 0 Or ParseVisualizationPrecision > 10 Then _
        Err.Raise vbObjectError + 3992, "CCalculationProfile", _
            settingName & " должен быть целым числом от 0 до 10."
End Function
