Attribute VB_Name = "modTorsionPostTables"
Option Explicit

' ==========================================================================
' Точка входа сборщика для подготовки таблиц кручения
' ==========================================================================
' Запись, очистку и оформление выполняет CTorsionSummaryWriter.
' Процедура доступна сборщику через Application.Run; НДМ не решает.

' Создаёт пустые самостоятельные выходы при сборке, не меняя настройки.
Public Sub InitializeTorsionPostOutput()
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    Dim snapshot As New CCalculationSnapshot, results As New Collection
    Dim writer As New CTorsionSummaryWriter
    writer.CaptureSummary snapshot, results, units
    Dim warning As String: warning = writer.RestoreSummary(ThisWorkbook, snapshot)
    If Len(warning) > 0 Then Err.Raise 5, , warning
End Sub
