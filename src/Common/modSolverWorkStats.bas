Attribute VB_Name = "modSolverWorkStats"
Option Explicit

' ==========================================================================
' Единый счетчик тяжелых запусков CSectionSolver
' ==========================================================================
' Модуль хранит один общий для расчетного запуска счетчик фактических попыток
' найти равновесное НДС через CSectionSolver.Solve. Верхние уровни batch,
' capacity и crack не должны восстанавливать это число через свои итерации:
' источником правды является сам CSectionSolver.

Private mSectionEquilibriumSolveCount As Long ' Количество фактических запусков CSectionSolver.Solve после валидации входа.

' Сбрасывает счетчик перед новым пакетным расчетом.
Public Sub ResetSectionEquilibriumSolveCount()
    mSectionEquilibriumSolveCount = 0
End Sub

' Регистрирует одну реальную попытку CSectionSolver подобрать равновесие.
Public Sub RecordSectionEquilibriumSolve()
    mSectionEquilibriumSolveCount = mSectionEquilibriumSolveCount + 1
End Sub

' Возвращает число запусков CSectionSolver.Solve с последнего сброса.
Public Function SectionEquilibriumSolveCount() As Long
    SectionEquilibriumSolveCount = mSectionEquilibriumSolveCount
End Function
