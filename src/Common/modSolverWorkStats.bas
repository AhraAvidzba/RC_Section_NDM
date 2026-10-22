Attribute VB_Name = "modSolverWorkStats"
Option Explicit

' ==========================================================================
' Единые счетчики фактической работы решателей
' ==========================================================================
' Счетчики меняются в самих численных циклах, включая неудачные попытки.
' Расчёт сочетаний сохраняет копию счётчиков для сводки; следующие расчёты её не меняют.

Private mSectionEquilibriumSolveCount As Long ' Количество фактических запусков CSectionSolver.Solve после валидации входа.
Private mSectionEquilibriumIterationCount As Long ' Итерации Newton/Secant на всех ступенях и стартах НДС.
Private mUltimateStrainStartCount As Long ' Начальные плоскости общего UltimateStrain и реальные теплые перезапуски.
Private mUltimateStrainIterationCount As Long ' Число итераций общего поиска предела без дополнительных оценок Якобиана и подбора шага.

' Сбрасывает все четыре счётчика перед новым пакетным расчётом.
Public Sub ResetSectionEquilibriumSolveCount()
    mSectionEquilibriumSolveCount = 0
    mSectionEquilibriumIterationCount = 0
    mUltimateStrainStartCount = 0
    mUltimateStrainIterationCount = 0
End Sub

' Регистрирует одну реальную попытку CSectionSolver подобрать равновесие.
Public Sub RecordSectionEquilibriumSolve()
    mSectionEquilibriumSolveCount = mSectionEquilibriumSolveCount + 1
End Sub

' Возвращает число запусков CSectionSolver.Solve с последнего сброса.
Public Function SectionEquilibriumSolveCount() As Long
    SectionEquilibriumSolveCount = mSectionEquilibriumSolveCount
End Function

' Одна итерация фактически начатого цикла равновесия, а не интегрирование пробы.
Public Sub RecordSectionEquilibriumIteration()
    mSectionEquilibriumIterationCount = mSectionEquilibriumIterationCount + 1
End Sub

' Возвращает число итераций равновесия на всех ступенях и повторных стартах после последнего сброса.
Public Function SectionEquilibriumIterationCount() As Long
    SectionEquilibriumIterationCount = mSectionEquilibriumIterationCount
End Function

' Каждая новая попытка с плоскости, включая повторную, считается отдельно.
Public Sub RecordUltimateStrainStart()
    mUltimateStrainStartCount = mUltimateStrainStartCount + 1
End Sub

' Возвращает число фактически начатых стартов поиска UltimateStrain, включая повторные плоскости.
Public Function UltimateStrainStartCount() As Long
    UltimateStrainStartCount = mUltimateStrainStartCount
End Function

' Регистрирует итерацию общего поиска UltimateStrain; отдельные пробы Якобиана и длины шага сюда не входят.
Public Sub RecordUltimateStrainIteration()
    mUltimateStrainIterationCount = mUltimateStrainIterationCount + 1
End Sub

' Возвращает число итераций общего поиска UltimateStrain с последнего сброса счётчиков.
Public Function UltimateStrainIterationCount() As Long
    UltimateStrainIterationCount = mUltimateStrainIterationCount
End Function
