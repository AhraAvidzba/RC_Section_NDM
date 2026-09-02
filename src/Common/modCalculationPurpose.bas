Attribute VB_Name = "modCalculationPurpose"
Option Explicit

' ==========================================================================
' ТИПИЗИРОВАННЫЕ РОЛИ МАТЕРИАЛЬНЫХ МОДЕЛЕЙ И NAMED STATES
' ==========================================================================
' Модуль держит общие enum-ы и преобразование текста из Config/Results в
' типизированные значения. Здесь нет формул СП и построения диаграмм:
' физическую модель строит CMaterialModelProvider, а solver получает уже
' готовый CMaterialDiagram.

Public Enum ECalculationPurpose
    cpStrength = 1        ' I ГПС: расчет прочности и несущей способности.
    cpMcrc = 2            ' II ГПС: состояние образования трещины.
    cpCrackedNDS = 3      ' II ГПС: НДС раскрытой трещины без растянутого бетона.
    cpStateSolution = 4   ' Прямой НДС заданного LC; физическая база передается отдельно.
End Enum

Public Enum EMaterialValueSet
    mvsULS = 1            ' Характеристики I ГПС: Rb/Rbt/Rs/Rsc.
    mvsSLS = 2            ' Характеристики II ГПС: Rb,ser/Rbt,ser/Rs,ser/Rsc,ser.
End Enum

Public Enum EConcreteDiagramKind
    cdkTwoLine = 1        ' Двухлинейная диаграмма бетона.
    cdkThreeLine = 2      ' Трехлинейная диаграмма бетона.
End Enum

Public Enum EConcreteTensionKind
    ctkIgnore = 1         ' Растянутый бетон не работает.
    ctkUseDiagram = 2     ' Растянутый бетон работает по выбранной диаграмме.
End Enum

Public Enum ESteelDiagramKind
    sdkTwoLine = 1        ' Двухлинейная диаграмма арматуры.
    sdkThreeLine = 2      ' Трехлинейная диаграмма арматуры.
End Enum

Public Enum ESectionStateType
    sstStrengthState = 1          ' Прямое НДС по модели прочности.
    sstCapacityState = 2          ' Предельное НДС, найденное capacity solver-ом.
    sstBeforeMcrcState = 3        ' НДС на пороге Mcrc до выключения растянутого бетона.
    sstAfterMcrcState = 4         ' НДС при lambda_crc сразу после выключения растянутого бетона.
    sstCrackedState = 5           ' Текущее НДС с раскрытой трещиной.
End Enum

Public Enum EVisualizationQuantity
    vqStress = 1          ' Для схемы/AutoCAD показываются напряжения.
    vqStrain = 2          ' Для схемы/AutoCAD показываются деформации.
End Enum

' Преобразует пользовательский или snapshot-текст в роль материальной модели.
Public Function PurposeFromText(ByVal purposeText As String) As ECalculationPurpose
    Select Case LCase$(Trim$(purposeText))
        Case "strength", "group1", "uls", "i", "1"
            PurposeFromText = cpStrength
        Case "mcrc"
            PurposeFromText = cpMcrc
        Case "crackednds", "a_crc", "crack", "group2", "sls", "ii", "2"
            PurposeFromText = cpCrackedNDS
        Case "statesolution", "state", "directstate"
            PurposeFromText = cpStateSolution
        Case Else
            Err.Raise vbObjectError + 3301, "modCalculationPurpose", _
                "Неизвестная расчетная цель material provider: " & purposeText
    End Select
End Function

' Возвращает канонический текст для Results, диагностики и справки.
Public Function PurposeToText(ByVal purpose As ECalculationPurpose) As String
    Select Case purpose
        Case cpStrength
            PurposeToText = "Strength"
        Case cpMcrc
            PurposeToText = "Mcrc"
        Case cpCrackedNDS
            PurposeToText = "CrackedNDS"
        Case cpStateSolution
            PurposeToText = "StateSolution"
        Case Else
            Err.Raise vbObjectError + 3302, "modCalculationPurpose", _
                "Неизвестный ECalculationPurpose."
    End Select
End Function

' Возвращает смысловую метку material model role для профильной архитектуры.
Public Function MaterialModelRoleToText(ByVal purpose As ECalculationPurpose) As String
    Select Case purpose
        Case cpStrength
            MaterialModelRoleToText = "Strength"
        Case cpMcrc
            MaterialModelRoleToText = "CrackInitiation"
        Case cpCrackedNDS
            MaterialModelRoleToText = "CrackedState"
        Case Else
            Err.Raise vbObjectError + 3310, "modCalculationPurpose", _
                "Для роли материальной модели требуется физическая расчетная цель."
    End Select
End Function

' Преобразует ValueSet из профиля в enum.
Public Function MaterialValueSetFromText(ByVal valueText As String) As EMaterialValueSet
    Select Case LCase$(Trim$(valueText))
        Case "uls", "i", "group1", "strength"
            MaterialValueSetFromText = mvsULS
        Case "sls", "ii", "group2", "service"
            MaterialValueSetFromText = mvsSLS
        Case Else
            Err.Raise vbObjectError + 3311, "modCalculationPurpose", _
                "ValueSet материальной модели должен быть ULS или SLS."
    End Select
End Function

Public Function MaterialValueSetToText(ByVal valueSet As EMaterialValueSet) As String
    Select Case valueSet
        Case mvsULS
            MaterialValueSetToText = "ULS"
        Case mvsSLS
            MaterialValueSetToText = "SLS"
        Case Else
            Err.Raise vbObjectError + 3312, "modCalculationPurpose", _
                "Неизвестный MaterialValueSet."
    End Select
End Function

Public Function ConcreteDiagramKindFromText(ByVal valueText As String) As EConcreteDiagramKind
    Select Case LCase$(Trim$(valueText))
        Case "threeline", "three"
            ConcreteDiagramKindFromText = cdkThreeLine
        Case "twoline", "two"
            ConcreteDiagramKindFromText = cdkTwoLine
        Case Else
            Err.Raise vbObjectError + 3313, "modCalculationPurpose", _
                "Диаграмма бетона должна быть TwoLine или ThreeLine."
    End Select
End Function

Public Function ConcreteDiagramKindToText(ByVal diagramKind As EConcreteDiagramKind) As String
    Select Case diagramKind
        Case cdkTwoLine
            ConcreteDiagramKindToText = "TwoLine"
        Case cdkThreeLine
            ConcreteDiagramKindToText = "ThreeLine"
        Case Else
            Err.Raise vbObjectError + 3314, "modCalculationPurpose", _
                "Неизвестный тип диаграммы бетона."
    End Select
End Function

Public Function ConcreteTensionKindFromText(ByVal valueText As String) As EConcreteTensionKind
    Select Case LCase$(Trim$(valueText))
        Case "usediagram", "use"
            ConcreteTensionKindFromText = ctkUseDiagram
        Case "ignore", "no", "none"
            ConcreteTensionKindFromText = ctkIgnore
        Case Else
            Err.Raise vbObjectError + 3315, "modCalculationPurpose", _
                "Растянутый бетон должен быть Ignore или UseDiagram."
    End Select
End Function

Public Function ConcreteTensionKindToText(ByVal tensionKind As EConcreteTensionKind) As String
    Select Case tensionKind
        Case ctkIgnore
            ConcreteTensionKindToText = "Ignore"
        Case ctkUseDiagram
            ConcreteTensionKindToText = "UseDiagram"
        Case Else
            Err.Raise vbObjectError + 3316, "modCalculationPurpose", _
                "Неизвестный режим растянутого бетона."
    End Select
End Function

Public Function SteelDiagramKindFromText(ByVal valueText As String) As ESteelDiagramKind
    Select Case LCase$(Trim$(valueText))
        Case "threeline", "three"
            SteelDiagramKindFromText = sdkThreeLine
        Case "twoline", "two"
            SteelDiagramKindFromText = sdkTwoLine
        Case Else
            Err.Raise vbObjectError + 3317, "modCalculationPurpose", _
                "Диаграмма арматуры должна быть TwoLine или ThreeLine."
    End Select
End Function

Public Function SteelDiagramKindToText(ByVal diagramKind As ESteelDiagramKind) As String
    Select Case diagramKind
        Case sdkTwoLine
            SteelDiagramKindToText = "TwoLine"
        Case sdkThreeLine
            SteelDiagramKindToText = "ThreeLine"
        Case Else
            Err.Raise vbObjectError + 3318, "modCalculationPurpose", _
                "Неизвестный тип диаграммы арматуры."
    End Select
End Function

Public Function SectionStateTypeFromText(ByVal valueText As String) As ESectionStateType
    Select Case LCase$(Trim$(valueText))
        Case "strengthstate", "strength", "прочность"
            SectionStateTypeFromText = sstStrengthState
        Case "capacitystate", "capacity", "предельное ндс"
            SectionStateTypeFromText = sstCapacityState
        Case "beforemcrcstate", "beforemcrc", "mcrc"
            SectionStateTypeFromText = sstBeforeMcrcState
        Case "aftermcrcstate", "aftermcrc"
            SectionStateTypeFromText = sstAfterMcrcState
        Case "crackedstate", "crack", "crackednds", "ндс при трещинах"
            SectionStateTypeFromText = sstCrackedState
        Case Else
            Err.Raise vbObjectError + 3319, "modCalculationPurpose", _
                "Visualization.State должен быть StrengthState, CapacityState, BeforeMcrcState, AfterMcrcState или CrackedState."
    End Select
End Function

Public Function SectionStateTypeToText(ByVal stateType As ESectionStateType) As String
    Select Case stateType
        Case sstStrengthState
            SectionStateTypeToText = "StrengthState"
        Case sstCapacityState
            SectionStateTypeToText = "CapacityState"
        Case sstBeforeMcrcState
            SectionStateTypeToText = "BeforeMcrcState"
        Case sstAfterMcrcState
            SectionStateTypeToText = "AfterMcrcState"
        Case sstCrackedState
            SectionStateTypeToText = "CrackedState"
        Case Else
            Err.Raise vbObjectError + 3320, "modCalculationPurpose", "Неизвестный StateType."
    End Select
End Function

Public Function MaterialRoleFromStateType(ByVal stateType As ESectionStateType) As ECalculationPurpose
    Select Case stateType
        Case sstStrengthState, sstCapacityState
            MaterialRoleFromStateType = cpStrength
        Case sstBeforeMcrcState
            MaterialRoleFromStateType = cpMcrc
        Case sstAfterMcrcState, sstCrackedState
            MaterialRoleFromStateType = cpCrackedNDS
        Case Else
            Err.Raise vbObjectError + 3321, "modCalculationPurpose", _
                "Для StateType не определена роль material model."
    End Select
End Function

Public Function VisualizationQuantityFromText(ByVal valueText As String) As EVisualizationQuantity
    Select Case LCase$(Trim$(valueText))
        Case "stress", "напряжения"
            VisualizationQuantityFromText = vqStress
        Case "strain", "деформации"
            VisualizationQuantityFromText = vqStrain
        Case Else
            Err.Raise vbObjectError + 3322, "modCalculationPurpose", _
                "Visualization.Quantity должен быть Stress или Strain."
    End Select
End Function

Public Function VisualizationQuantityToText(ByVal quantity As EVisualizationQuantity) As String
    Select Case quantity
        Case vqStress
            VisualizationQuantityToText = "Stress"
        Case vqStrain
            VisualizationQuantityToText = "Strain"
        Case Else
            Err.Raise vbObjectError + 3323, "modCalculationPurpose", _
                "Неизвестная Visualization.Quantity."
    End Select
End Function

' Проверяет, что цель является физической material model без numerical extension.
Public Function IsPhysicalPurpose(ByVal purpose As ECalculationPurpose) As Boolean
    IsPhysicalPurpose = (purpose = cpStrength Or purpose = cpMcrc Or purpose = cpCrackedNDS)
End Function

' Проверяет, что цель может быть базой для прямого StateSolution.
Public Function IsStateBasePurpose(ByVal purpose As ECalculationPurpose) As Boolean
    IsStateBasePurpose = (purpose = cpStrength Or purpose = cpCrackedNDS)
End Function
