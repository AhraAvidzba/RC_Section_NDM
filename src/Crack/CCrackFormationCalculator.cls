VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackFormationCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Orchestration-слой проверки образования нормальной трещины
' ==========================================================================
' CCrackFormationCalculator изолирует этап formation от формульного расчета
' ширины. На текущем шаге миграции он переиспользует существующий
' CCrackWidthCalculator как рабочий контекст поиска, но наружу возвращает
' отдельный CCrackFormationResult. Общая численная логика UltimateStrain /
' LoadMultiplier остается в CLimitSearchCoordinator и связанных CLimitSearch*
' классах, а не копируется здесь параллельно.

' Запускает проверку образования нормальной трещины и возвращает result-объект.
' Calculator передается снаружи уже настроенным, чтобы не дублировать чтение
' Config и не менять существующий численный pipeline.
Public Function CheckFormation(ByVal crack As CCrackWidthCalculator, ByVal statusPolicy As CBatchStatusPolicy, _
        ByVal section As CSectionModel, ByVal materialProvider As CMaterialModelProvider, _
        ByVal crackedStateSpec As CMaterialModelSpec, ByVal crackInitiationSpec As CMaterialModelSpec, _
        ByVal targetN As Double, ByVal targetMx As Double, ByVal targetMy As Double, _
        ByVal loadStateForClassification As CSectionLoadState, ByVal centralReferenceX As Double, _
        ByVal centralReferenceY As Double) As CCrackFormationResult
    If statusPolicy Is Nothing Then Set statusPolicy = New CBatchStatusPolicy

    crack.CheckFormationBeforeWidth section, materialProvider, crackedStateSpec, _
        crackInitiationSpec, targetN, targetMx, targetMy, loadStateForClassification, _
        centralReferenceX, centralReferenceY

    Dim result As CCrackFormationResult
    Set result = New CCrackFormationResult
    result.InitializeFromCalculator crack, statusPolicy.CrackFormationMetaFromCalculator(crack)
    Set CheckFormation = result
End Function
