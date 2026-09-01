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
Private mStrengthSpec As CMaterialModelSpec ' Материальная модель прочности.
Private mCrackInitiationSpec As CMaterialModelSpec ' Материальная модель Mcrc для Auto psi_s.
Private mCrackedStateSpec As CMaterialModelSpec ' Материальная модель НДС с раскрытой трещиной.
Private mVisualizationState As ESectionStateType ' Какой StateType показывать на схеме и в AutoCAD.
Private mVisualizationQuantity As EVisualizationQuantity ' Stress или Strain для визуализации.
Private mHasVisualizationState As Boolean ' Визуальная настройка задана в профиле.
Private mHasVisualizationQuantity As Boolean ' Визуальная величина задана в профиле.

Private Sub Class_Initialize()
    Set mStrengthSpec = New CMaterialModelSpec
    Set mCrackInitiationSpec = New CMaterialModelSpec
    Set mCrackedStateSpec = New CMaterialModelSpec
End Sub

' Заполняет простые поля профиля. MaterialModelSpec задаются отдельными
' методами, потому что у пустых резервных профилей они могут отсутствовать.
Public Sub Initialize(ByVal profileId As String, ByVal displayName As String, ByVal description As String, _
        ByVal strengthDirectState As Boolean, ByVal strengthCapacity As Boolean, ByVal crackWidth As Boolean)
    mProfileId = Trim$(profileId)
    mDisplayName = Trim$(displayName)
    mDescription = Trim$(description)
    mStrengthDirectState = strengthDirectState
    mStrengthCapacity = strengthCapacity
    mCrackWidth = crackWidth
End Sub

Public Sub SetStrengthSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mStrengthSpec = spec.Clone
End Sub

Public Sub SetCrackInitiationSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mCrackInitiationSpec = spec.Clone
End Sub

Public Sub SetCrackedStateSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Exit Sub
    Set mCrackedStateSpec = spec.Clone
End Sub

Public Sub SetVisualization(ByVal stateText As String, ByVal quantityText As String)
    If Len(Trim$(stateText)) > 0 Then
        mVisualizationState = SectionStateTypeFromText(stateText)
        mHasVisualizationState = True
    End If
    If Len(Trim$(quantityText)) > 0 Then
        mVisualizationQuantity = VisualizationQuantityFromText(quantityText)
        mHasVisualizationQuantity = True
    End If
End Sub

' Профиль считается настроенным только если задано имя и есть хотя бы один
' запрошенный расчет. Пустые PR3/PR4 остаются резервом и не проходят сюда.
Public Function IsConfigured() As Boolean
    IsConfigured = (Len(mDisplayName) > 0 And (mStrengthDirectState Or mStrengthCapacity Or mCrackWidth))
End Function

' Валидирует профиль только для фактически назначенного LC. Visualization.State
' здесь не проверяется на наличие будущего состояния: это ошибка визуализации,
' а не причина блокировать batch-расчет.
Public Function ValidateForCalculation(ByVal psiMode As String, ByRef errorText As String) As Boolean
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
    If mCrackWidth And Not mCrackedStateSpec.IsComplete Then
        errorText = "Профиль " & mProfileId & ": не задана MaterialModel.CrackedState."
        Exit Function
    End If
    If mCrackWidth And StrComp(Trim$(psiMode), "Auto", vbTextCompare) = 0 And _
            Not mCrackInitiationSpec.IsComplete Then
        errorText = "Профиль " & mProfileId & ": для SLS.Crack.PsiMode=Auto нужна MaterialModel.CrackInitiation."
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

Public Property Get StrengthSpec() As CMaterialModelSpec
    Set StrengthSpec = mStrengthSpec
End Property

Public Property Get CrackInitiationSpec() As CMaterialModelSpec
    Set CrackInitiationSpec = mCrackInitiationSpec
End Property

Public Property Get CrackedStateSpec() As CMaterialModelSpec
    Set CrackedStateSpec = mCrackedStateSpec
End Property

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
