Attribute VB_Name = "modResultStatus"
Option Explicit

' ==========================================================================
' ТИПИЗИРОВАННЫЕ СТАТУСЫ И КОДЫ РЕЗУЛЬТАТОВ
' ==========================================================================
' Модуль задает общий внутренний язык результата расчета. Эти enum-ы не
' являются пользовательским выводом: наружные строки OK/FAIL/NumFail и цвета
' формирует CResultStatusPolicy. Инженерные классы должны передавать дальше
' именно InternalStatus + ResultCode, чтобы writer-ы не парсили StopReason.

Public Enum EResultInternalStatus
    rsUnset = 0                  ' Результат еще не сформирован; в завершенном расчете это архитектурная ошибка.
    rsNotRequested = 1           ' Расчет отключен профилем или настройкой пользователя.
    rsNotApplicable = 2          ' Проверка физически или методически неприменима для данного LC.
    rsBlockedByDependency = 3    ' Этап не запускался, потому что обязательный предыдущий результат отсутствует.
    rsSuccess = 4                ' Расчет корректно завершен и проверка прошла либо дала допустимый физический результат.
    rsCheckFailed = 5            ' Равновесие или данные получены, но инженерная проверка не проходит.
    rsNumericalFailure = 6       ' Численная процедура не сошлась там, где действительно искалось равновесие или предел.
    rsInvalidInput = 7           ' Некорректные пользовательские исходные данные конкретного расчета.
    rsInvalidConfiguration = 8   ' Некорректная настройка Config или несовместимые параметры расчетного профиля.
    rsSuccessWithWarning = 9     ' Корректный результат получен через fallback или с важным предупреждением.
    rsInternalError = 10         ' Непредусмотренная внутренняя ошибка программы.
End Enum

Public Enum EResultCode
    rcNone = 0                       ' Нет уточняющего кода; применимо для простого OK.
    rcNotRequested = 1               ' Расчет не был запрошен профилем.
    rcNotApplicable = 2              ' Расчет неприменим к текущему состоянию.
    rcBlockedByDependency = 3        ' Нет обязательного предыдущего результата.
    rcCheckPassed = 4                ' Инженерная проверка пройдена.
    rcCheckFailed = 5                ' Инженерная проверка не пройдена без численной ошибки.
    rcPhysicalLimitExceeded = 6      ' НДС найдено, но вышло за физическую диаграмму материала.
    rcInitialStateBeyondLimit = 7    ' Базовая точка траектории уже за предельным состоянием; для Capacity это BaseFail.
    rcCriterionNotReached = 8        ' Search корректно установил, что критерий не достигается на выбранной траектории.
    rcSearchBoundReached = 9         ' Достигнута техническая граница поиска без доказательства физического недостижения.
    rcNumericalFailure = 10          ' Общая численная несходимость solver/search.
    rcSingularTangent = 11           ' Вырожденная матрица/касательная в численном решении.
    rcInvalidInput = 12              ' Ошибка пользовательских входных данных.
    rcInvalidConfiguration = 13      ' Ошибка настроек Config или профиля.
    rcInternalError = 14             ' Непредусмотренная внутренняя ошибка.
    rcCrackWidthExceeded = 15        ' Ширина нормальной трещины больше допуска.
    rcLongitudinalCrackExceeded = 16 ' Проверка продольных трещин не проходит.
    rcCrackNotFormed = 17            ' Нормальная трещина физически не образуется.
    rcStateOutsidePhysicalRange = 18 ' Равновесие найдено только за физическим диапазоном диаграммы.
    rcSolverDidNotReturnState = 19   ' Оркестратор не получил объект solver-а для требуемого named-state.
End Enum

Public Enum EResultKind
    rkGeneric = 0                ' Общий результат без специальной инженерной трактовки.
    rkDirectState = 1            ' Прямое НДС заданного сочетания.
    rkCapacity = 2               ' Поиск предельной несущей способности.
    rkCrackFormation = 3         ' Поиск образования нормальной трещины.
    rkCrackWidth = 4             ' Формульный расчет ширины нормальной трещины.
    rkLongitudinalCrack = 5      ' Проверка продольных трещин по сжатому бетону.
    rkStability = 6              ' Проверка продольного изгиба и устойчивости.
End Enum

' Возвращает каноническое имя внутреннего статуса для diagnostic/report.
Public Function ResultInternalStatusToText(ByVal value As EResultInternalStatus) As String
    Select Case value
        Case rsUnset: ResultInternalStatusToText = "rsUnset"
        Case rsNotRequested: ResultInternalStatusToText = "rsNotRequested"
        Case rsNotApplicable: ResultInternalStatusToText = "rsNotApplicable"
        Case rsBlockedByDependency: ResultInternalStatusToText = "rsBlockedByDependency"
        Case rsSuccess: ResultInternalStatusToText = "rsSuccess"
        Case rsCheckFailed: ResultInternalStatusToText = "rsCheckFailed"
        Case rsNumericalFailure: ResultInternalStatusToText = "rsNumericalFailure"
        Case rsInvalidInput: ResultInternalStatusToText = "rsInvalidInput"
        Case rsInvalidConfiguration: ResultInternalStatusToText = "rsInvalidConfiguration"
        Case rsSuccessWithWarning: ResultInternalStatusToText = "rsSuccessWithWarning"
        Case rsInternalError: ResultInternalStatusToText = "rsInternalError"
        Case Else: ResultInternalStatusToText = "rsInternalError"
    End Select
End Function

' Возвращает каноническое имя машинного кода результата.
Public Function ResultCodeToText(ByVal value As EResultCode) As String
    Select Case value
        Case rcNone: ResultCodeToText = "NONE"
        Case rcNotRequested: ResultCodeToText = "NOT_REQUESTED"
        Case rcNotApplicable: ResultCodeToText = "NOT_APPLICABLE"
        Case rcBlockedByDependency: ResultCodeToText = "BLOCKED_BY_DEPENDENCY"
        Case rcCheckPassed: ResultCodeToText = "CHECK_PASSED"
        Case rcCheckFailed: ResultCodeToText = "CHECK_FAILED"
        Case rcPhysicalLimitExceeded: ResultCodeToText = "PHYSICAL_LIMIT_EXCEEDED"
        Case rcInitialStateBeyondLimit: ResultCodeToText = "INITIAL_STATE_BEYOND_LIMIT"
        Case rcCriterionNotReached: ResultCodeToText = "CRITERION_NOT_REACHED"
        Case rcSearchBoundReached: ResultCodeToText = "SEARCH_BOUND_REACHED"
        Case rcNumericalFailure: ResultCodeToText = "NUMERICAL_FAILURE"
        Case rcSingularTangent: ResultCodeToText = "SINGULAR_TANGENT"
        Case rcInvalidInput: ResultCodeToText = "INVALID_INPUT"
        Case rcInvalidConfiguration: ResultCodeToText = "INVALID_CONFIGURATION"
        Case rcInternalError: ResultCodeToText = "INTERNAL_ERROR"
        Case rcCrackWidthExceeded: ResultCodeToText = "CRACK_WIDTH_EXCEEDED"
        Case rcLongitudinalCrackExceeded: ResultCodeToText = "LONGITUDINAL_CRACK_EXCEEDED"
        Case rcCrackNotFormed: ResultCodeToText = "CRACK_NOT_FORMED"
        Case rcStateOutsidePhysicalRange: ResultCodeToText = "STATE_OUTSIDE_PHYSICAL_RANGE"
        Case rcSolverDidNotReturnState: ResultCodeToText = "SOLVER_DID_NOT_RETURN_STATE"
        Case Else: ResultCodeToText = "INTERNAL_ERROR"
    End Select
End Function

' Возвращает каноническое имя вида результата для отчета и отладки.
Public Function ResultKindToText(ByVal value As EResultKind) As String
    Select Case value
        Case rkDirectState: ResultKindToText = "DirectState"
        Case rkCapacity: ResultKindToText = "Capacity"
        Case rkCrackFormation: ResultKindToText = "CrackFormation"
        Case rkCrackWidth: ResultKindToText = "CrackWidth"
        Case rkLongitudinalCrack: ResultKindToText = "LongitudinalCrack"
        Case rkStability: ResultKindToText = "Stability"
        Case Else: ResultKindToText = "Generic"
    End Select
End Function
