Attribute VB_Name = "modWorkbookPrintLayout"
Option Explicit

' ==========================================================================
' Активация сохранённых границ печати в локализованном Excel
' ==========================================================================
' XLSM хранит стандартное имя Print_Area. Если Excel прочитал его как обычное
' локальное имя, активирует те же границы через PageSetup текущей локали.
' Ячейки, размеры строк/столбцов и пользовательские границы не меняет.

' При открытии использует только уже заданную область; новую область не назначает.
Public Sub ActivateDeclaredPrintAreas(ByVal workbook As Workbook)
    Dim sheet As Worksheet, declared As Name, area As Range
    Dim wasSaved As Boolean, modified As Boolean, reference As String, nameText As String
    Dim number As Long, description As String, removed As Boolean
    wasSaved = workbook.Saved
    On Error GoTo Failed
    For Each sheet In workbook.Worksheets
        If Len(sheet.PageSetup.PrintArea) = 0 Then
            Set declared = DeclaredAreaName(sheet)
            If Not declared Is Nothing Then
                Set area = declared.RefersToRange
                If area.Parent Is sheet Then
                    reference = declared.RefersTo: nameText = declared.Name
                    declared.Delete: removed = True ' Не оставляем обычное имя рядом с локализованным служебным именем при сохранении.
                    sheet.PageSetup.PrintArea = area.Address(True, True, xlA1)
                    removed = False: modified = True
                End If
            End If
        End If
    Next sheet
    If modified And wasSaved Then workbook.Saved = True ' Активация тех же границ не создаёт новых пользовательских данных.
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If removed Then
        On Error Resume Next
        workbook.Names.Add Name:=nameText, RefersTo:=reference
        On Error GoTo 0
    End If
    Err.Raise number, "modWorkbookPrintLayout.ActivateDeclaredPrintAreas", description
End Sub

' Находит только локальное стандартное имя области печати данного листа.
Private Function DeclaredAreaName(ByVal sheet As Worksheet) As Name
    Dim item As Name, leaf As String
    For Each item In sheet.Names
        leaf = Mid$(item.Name, InStrRev(item.Name, "!") + 1)
        If StrComp(leaf, "Print_Area", vbTextCompare) = 0 Or _
           StrComp(leaf, "_xlnm.Print_Area", vbTextCompare) = 0 Then
            Set DeclaredAreaName = item
            Exit Function
        End If
    Next item
End Function
