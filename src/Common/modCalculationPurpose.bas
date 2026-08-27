Attribute VB_Name = "modCalculationPurpose"
Option Explicit

' ==========================================================================
' ТИПИЗИРОВАННЫЕ ЦЕЛИ РАСЧЕТА МАТЕРИАЛЬНОЙ МОДЕЛИ
' ==========================================================================
' Модуль держит единый enum для выбора диаграммы материала и небольшие
' функции преобразования текста из Config/Results в типизированное значение.
' Здесь нет формул СП и нет построения диаграмм: физический выбор остается в
' CMaterialModelProvider, а solver получает уже готовый CMaterialDiagram.

Public Enum ECalculationPurpose
    cpStrength = 1        ' I ГПС: расчет прочности и несущей способности.
    cpMcrc = 2            ' II ГПС: состояние образования трещины.
    cpCrackedNDS = 3      ' II ГПС: НДС раскрытой трещины без растянутого бетона.
    cpStateSolution = 4   ' Прямой НДС заданного LC; требует отдельной физической базы.
End Enum

' Преобразует пользовательский или snapshot-текст в enum. Алиасы оставлены
' только на границе с Excel, чтобы внутри расчетного кода не ходили свободные строки.
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

' Возвращает канонический текст для записи в Results, диагностику и справку.
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

' True для целей, которые являются физической базой диаграммы без numerical extension.
Public Function IsPhysicalPurpose(ByVal purpose As ECalculationPurpose) As Boolean
    IsPhysicalPurpose = (purpose = cpStrength Or purpose = cpMcrc Or purpose = cpCrackedNDS)
End Function

' True для физических целей, которыми разрешено подпитать прямой StateSolution.
Public Function IsStateBasePurpose(ByVal purpose As ECalculationPurpose) As Boolean
    IsStateBasePurpose = (purpose = cpStrength Or purpose = cpCrackedNDS)
End Function

' Выбирает физическую базу StateSolution по группе сочетания.
' Group1 получает Strength, Group2 получает CrackedNDS.
Public Function StateBasePurposeForCalculationType(ByVal calculationType As String) As ECalculationPurpose
    Select Case LCase$(Trim$(calculationType))
        Case "group2", "2", "sls", "crack", "crackonly"
            StateBasePurposeForCalculationType = cpCrackedNDS
        Case Else
            StateBasePurposeForCalculationType = cpStrength
    End Select
End Function
