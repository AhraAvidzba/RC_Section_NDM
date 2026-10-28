VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarGroupAnnotationBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общие подписи количества и диаметров арматуры
' ==========================================================================
' Собирает подпись стержней каждой заданной группы с разбивкой по диаметрам.
' Стержень учитывается только в своей основной группе подписи. Дополнительные
' связи совпавшей позиции не добавляют арматуру и не увеличивают количество в подписи.

' Добавляет подписи по готовым группам стержней и их диаметрам.
' Использует переданные места привязки; координаты стержней и состав их групп не меняет.
Public Sub AddRebarLabels(ByVal annotations As CSectionAnnotations, ByVal rebars As CRebarLayout, _
        ByRef anchors() As TSectionAnnotation, ByVal anchorCount As Long)
    If annotations Is Nothing Then Err.Raise vbObjectError + 5290, "CRebarGroupAnnotationBuilder", "Коллекция аннотаций не передана."
    If rebars Is Nothing Then Exit Sub
    If rebars.Count <= 0 Or anchorCount <= 0 Then Exit Sub

    Dim groupedDiameters As Object
    Set groupedDiameters = CreateObject("Scripting.Dictionary")
    groupedDiameters.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To rebars.Count
        Dim groupName As String
        groupName = rebars.BarAnnotationGroupName(i)
        If Len(groupName) > 0 Then
            Dim diameterText As String
            diameterText = FormatNumberInvariant(rebars.Diameter(i))
            Dim key As String
            key = groupName & "|" & diameterText
            If groupedDiameters.Exists(key) Then
                groupedDiameters(key) = CLng(groupedDiameters(key)) + 1
            Else
                groupedDiameters.Add key, 1
            End If
        End If
    Next i

    If groupedDiameters.Count = 0 Then Exit Sub

    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    labels.CompareMode = vbTextCompare

    Dim keyVariant As Variant
    For Each keyVariant In groupedDiameters.Keys
        Dim parts As Variant
        parts = Split(CStr(keyVariant), "|")
        Dim itemText As String
        itemText = CStr(groupedDiameters(keyVariant)) & ChrW$(216) & CStr(parts(1))
        If labels.Exists(CStr(parts(0))) Then
            labels(CStr(parts(0))) = CStr(labels(CStr(parts(0)))) & " + " & itemText
        Else
            labels.Add CStr(parts(0)), itemText
        End If
    Next keyVariant

    For i = 1 To anchorCount
        groupName = Trim$(anchors(i).Name)
        If labels.Exists(groupName) Then
            annotations.AddRebarLabel "REBAR_" & groupName, _
                anchors(i).StartX, anchors(i).StartY, _
                anchors(i).EndX, anchors(i).EndY, _
                anchors(i).NormalX, anchors(i).NormalY, CStr(labels(groupName)), boundaryIsLocal:=anchors(i).BoundaryIsLocal
            labels.Remove groupName ' Несколько привязок одной группы не создают повторную подпись.
        End If
    Next i
End Sub

' Форматирует диаметр стержня для групповой подписи без зависимости от
' регионального десятичного разделителя Windows.
Private Function FormatNumberInvariant(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.###"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatNumberInvariant = textValue
End Function
