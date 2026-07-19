Attribute VB_Name = "CAD"
Option Explicit
Public startRow As Integer, StartColumn As Integer
Public objApp As AcadApplication
Public objDoc As AcadDocument
Public cmdecho As Integer
'установка начального столбца и строки
Public Sub SetStartValue()
    startRow = 10
    StartColumn = 0
End Sub
Public Function setVerticalRange(Column As String, startRow As Integer, count As Integer) As String
Dim endRow As Integer
    endRow = startRow + count - 1
    setVerticalRange = Column & CStr(startRow) & ":" & Column & CStr(endRow)
End Function
'подключение к КАД-программе
Public Sub cadConnect()
    Set objApp = GetObject(, "AutoCAD.Application")
    Set objDoc = objApp.ActiveDocument
    cmdecho = objDoc.GetVariable("CMDECHO")
    Call objDoc.SetVariable("CMDECHO", 0)
End Sub
'отключение от КАД-программы
Public Sub cadDisconnect()
    Call objDoc.SetVariable("CMDECHO", cmdecho)
    Set objApp = Nothing
    Set objDoc = Nothing
End Sub
'сделать слой текущим, при необходимости создать слой
Public Sub SetActiveLayer(Name As String, CreateLayer As Boolean, color As Integer)
Dim layCurrent As AcadLayer
    On Error Resume Next
    Set layCurrent = objDoc.Layers(Name)
    If CreateLayer = True Then
        If layCurrent Is Nothing Then
            Set layCurrent = objDoc.Layers.Add(Name)
        End If
        layCurrent.color = color
        objDoc.ActiveLayer = layCurrent
    Else
        Exit Sub
    End If
End Sub
'удаление объектов на слое
Public Sub ClearLayer(ObjectType As String, LayerName As String)
    On Error GoTo ErrorHandler
Dim osSet As AcadSelectionSet
Dim FilterType(1) As Integer
Dim FilterData(1) As Variant
    FilterType(0) = 0
    FilterData(0) = ObjectType
    FilterType(1) = 8
    FilterData(1) = LayerName
    With objDoc.SelectionSets
        While .count > 0
            .Item(0).Delete
        Wend
        Set osSet = .Add("$NewOne$")
    End With
    osSet.Select acSelectionSetAll, , , FilterType, FilterData
Dim ent As AcadEntity
    For Each ent In osSet
        ent.Delete
    Next
    DeleteSSet
ErrorHandler:
    Exit Sub
End Sub
'удаление всех наборов
Public Sub DeleteSSet()
    On Error GoTo ErrorHandler
     With objDoc.SelectionSets
        While .count > 0
            .Item(0).Delete
        Wend
    End With
ErrorHandler:
    Exit Sub
End Sub
'добавить текст
Public Sub AddText(ByVal TextValue As String, ByVal X As Double, ByVal Y As Double, ByVal TextSize As Double)
Dim TextPosition(0 To 2) As Double
    TextPosition(0) = X
    TextPosition(1) = Y
    TextPosition(2) = 0
Dim objText As AcadText
    Set objText = objDoc.ModelSpace.AddText(TextValue, TextPosition, TextSize)
    objText.Alignment = acAlignmentMiddleCenter
    objText.TextAlignmentPoint = TextPosition
    Set objText = Nothing
End Sub
'добавить штриховку
Public Sub AddHatch(objHandle As String, patternName As String)
    On Error GoTo ErrorHandler
Dim autocadObj As AcadObject
Dim hatchObj As AcadHatch
Dim patternType As Long
    patternType = acHatchPatternTypePreDefined
    Set autocadObj = objDoc.HandleToObject(objHandle)
    Set hatchObj = objDoc.ModelSpace.AddHatch(patternType, patternName, True)
        'hatchObj.PatternScale = 2 'масштаб штриховки
    Dim outerLoop(0 To 0) As AcadObject
    Set outerLoop(0) = autocadObj
    hatchObj.AppendOuterLoop (outerLoop)
    hatchObj.Evaluate
    Set autocadObj = Nothing
    Set hatchObj = Nothing
    Set outerLoop(0) = Nothing
ErrorHandler:
    Exit Sub
End Sub
'добавить 3D
Public Sub Add3D(objHandle As String, height As Double)
    On Error GoTo ErrorHandler
Dim autocadObj As AcadObject
    Set autocadObj = objDoc.HandleToObject(objHandle)
    Dim regionObj As AcadRegion
    Set regionObj = autocadObj
    Dim solidObj As Acad3DSolid
    Set solidObj = objDoc.ModelSpace.AddExtrudedSolid(regionObj, height, 0)
    Set autocadObj = Nothing
    Set regionObj = Nothing
    Set solidObj = Nothing
ErrorHandler:
    Exit Sub
End Sub
'обработка ошибок и вывод предупреждения
Public Sub ErrorHandlerSub(Number As Long)
    If Err.Number <> 0 And Err.Number <> 40040 Then
    Select Case Err.Number
        Case 13
            Call MsgBox("Для макроса указана некорретная библиотека AutoCAD. Переназначьте библиотеку согласно инструкции к программе", vbCritical, "Ошибка")
        Case 18
            Call MsgBox("Операция была приостановлена пользователем", vbInformation, "Завершение")
            Application.StatusBar = ""
        Case -2145386484
            Call MsgBox("Не найден объект с указателем. Проверьте корректность файла, при необходимости обновите геометрию", vbCritical, "Ошибка")
        Case 429
            Call MsgBox("Откройте AutoCAD и загрузите чертеж", vbCritical, "Ошибка")
        Case -2145320900
            Call MsgBox("Откройте чертеж в AutoCAD", vbCritical, "Ошибка")
        Case -2145386413
            Call MsgBox("Слой, на котором должны быть расположены объекты, заблокирован. Разблокируйте слой и запустите команду заново", vbCritical, "Ошибка")
        Case Else
            Call MsgBox("При выполнении возникла ошибка номер " & Err.Number, vbCritical, "Ошибка")
    End Select
    End If
End Sub
'создание нужных слоёв для задания сечения
Sub AddConstructLayers()
 On Error GoTo ErrorHandler
Application.EnableCancelKey = xlErrorHandler
cadConnect
   Call SetActiveLayer("Concrete", True, 252)
   Call SetActiveLayer("Reinf", True, 1)
   Call SetActiveLayer("Prestress", True, 6)
   Call SetActiveLayer("User", True, 4)
   Call SetActiveLayer("Concrete_points", True, 2)
   Call SetActiveLayer("Reinf_points", True, 3)
   Call SetActiveLayer("Prestress_points", True, 7)
   Call SetActiveLayer("User_points", True, 5)
   objDoc.Regen True
cadDisconnect
Call MsgBox("Необходимые для задания сечения слои созданы", vbInformation, "Завершено")

ErrorHandler:
    Exit Sub
End Sub
Sub GetAcadConcreteButton()
On Error GoTo ErrorHandler
    Application.DisplayStatusBar = True
    Application.EnableCancelKey = xlErrorHandler
    Dim acadRange, handleRange, weightRange, inertiaRange As String
    'диапазоны ячеек на листе
    acadRange = "A10:E10010"
    handleRange = "A10:A10010"
    weightRange = "Y10:Y10010"
    inertiaRange = "U10:V10010"
    'переделано без ссылки на имя листа
    ActiveSheet.EnableCalculation = False
    ActiveSheet.Range(acadRange).Clear
    ActiveSheet.Range(acadRange).HorizontalAlignment = xlCenter
    ActiveSheet.Range(inertiaRange).Clear
    ActiveSheet.Range(inertiaRange).HorizontalAlignment = xlCenter
    ActiveSheet.Range(weightRange).Clear
    ActiveSheet.Range(weightRange).HorizontalAlignment = xlCenter
    ActiveSheet.Range(handleRange).NumberFormat = "@"
    Dim osSet_reg, osSet_point As AcadSelectionSet
    SetStartValue
    cadConnect
    Dim FilterType_reg(0) As Integer
    Dim FilterData_reg(0) As Variant
    Dim FilterType_point(0) As Integer
    Dim FilterData_point(0) As Variant
    FilterType_reg(0) = 0
    FilterData_reg(0) = "REGION"
    FilterType_point(0) = 0
    FilterData_point(0) = "POINT"
    With objDoc.SelectionSets
        While .count > 0
            .Item(0).Delete
        Wend
        Set osSet_reg = .Add("$NewOne$")
        Set osSet_point = .Add("$NewTwo$")
    End With
    osSet_reg.Select acSelectionSetAll, , , FilterType_reg, FilterData_reg
    osSet_point.Select acSelectionSetAll, , , FilterType_point, FilterData_point
    Dim i, nCount As Integer
    nCount = osSet_reg.count + osSet_point.count
    If nCount > 10000 Then
        Call MsgBox("Количество объектов больше 10000. Уменьшите количество объектов", vbInformation, "Завершено")
        Exit Sub
    End If
    Dim ent As AcadEntity
    Dim minExt As Variant
    Dim maxExt As Variant
    Dim minX, maxX As Double
    Dim minY, maxY As Double
    Dim dimX, dimY As Double
    Dim sLayerName As String
    Dim objHandleArr() As String
    Dim AreaArr() As Double
    Dim XArr() As Double
    Dim YArr() As Double
    Dim JxArr() As Double 'главные центральные моменты инерции
    Dim JyArr() As Double
    Dim typeArr() As String
    Dim weightArr() As Double
    ReDim objHandleArr(1 To nCount, 1 To 1)
    ReDim AreaArr(1 To nCount, 1 To 1)
    ReDim XArr(1 To nCount, 1 To 1)
    ReDim YArr(1 To nCount, 1 To 1)
    ReDim JxArr(1 To nCount, 1 To 1)
    ReDim JyArr(1 To nCount, 1 To 1)
    ReDim typeArr(1 To nCount, 1 To 1)
    ReDim weightArr(1 To nCount, 1 To 1)
    Dim userWeight As Double
    'пользовательский вес
    userWeight = ActiveSheet.Cells(22, 50).Value2
    
    i = 1
    For Each ent In osSet_reg
        Dim regionObj As AcadRegion
        Set regionObj = ent
        sLayerName = UCase(regionObj.Layer)
        If sLayerName = "CONCRETE" Or sLayerName = "REINF" Or sLayerName = "USER" Or sLayerName = "PRESTRESS" Then
            objHandleArr(i, 1) = regionObj.Handle
            'площадь в см^2
            AreaArr(i, 1) = regionObj.Area / 100#
            XArr(i, 1) = regionObj.Centroid(0)
            YArr(i, 1) = regionObj.Centroid(1)
            'моменты инерции в см^4
            JxArr(i, 1) = regionObj.PrincipalMoments(0) / 10000#
            JyArr(i, 1) = regionObj.PrincipalMoments(1) / 10000#
            If sLayerName = "CONCRETE" Then
                typeArr(i, 1) = "c"
                weightArr(i, 1) = 2.5 'тс/м^3 объёмный вес бетона
            ElseIf sLayerName = "REINF" Then
                typeArr(i, 1) = "r"
                weightArr(i, 1) = 7.85 'тс/м^3 объёмный вес ненапрягаемой арматуры
            ElseIf sLayerName = "USER" Then
                typeArr(i, 1) = "u"
                weightArr(i, 1) = userWeight
            Else
                typeArr(i, 1) = "p"
                weightArr(i, 1) = 7.85 'тс/м^3 объёмный вес напрягаемой арматуры
            End If
            regionObj.GetBoundingBox minExt, maxExt
            If i = 1 Then
                minX = minExt(0)
                maxX = maxExt(0)
                minY = minExt(1)
                maxY = maxExt(1)
            End If
            If minExt(0) < minX Then minX = minExt(0)
            If maxExt(0) > maxX Then maxX = maxExt(0)
            If minExt(1) < minY Then minY = minExt(1)
            If maxExt(1) > maxY Then maxY = maxExt(1)
        End If
        Application.StatusBar = "Получено " & i & " объектов из " & nCount
        i = i + 1
    Next
    For Each ent In osSet_point
        Dim pointObj As AcadPoint
        Set pointObj = ent
        sLayerName = UCase(pointObj.Layer)
        If sLayerName = "CONCRETE_POINTS" Or sLayerName = "REINF_POINTS" Or sLayerName = "USER_POINTS" Or sLayerName = "PRESTRESS_POINTS" Then
            objHandleArr(i, 1) = pointObj.Handle
            AreaArr(i, 1) = 0
            XArr(i, 1) = pointObj.Coordinates(0)
            YArr(i, 1) = pointObj.Coordinates(1)
            If sLayerName = "CONCRETE_POINTS" Then
                typeArr(i, 1) = "cp"
                weightArr(i, 1) = 0
            ElseIf sLayerName = "REINF_POINTS" Then
                typeArr(i, 1) = "rp"
                weightArr(i, 1) = 0
            ElseIf sLayerName = "USER_POINTS" Then
                typeArr(i, 1) = "up"
                weightArr(i, 1) = 0
            Else
                typeArr(i, 1) = "pp"
                weightArr(i, 1) = 0
            End If
        End If
        Application.StatusBar = "Получено " & i & " объектов из " & nCount
        i = i + 1
    Next
    dimX = maxX - minX
    dimY = maxY - minY
    'переписано на активный лист
    ActiveSheet.Range(setVerticalRange("A", startRow, nCount)).Value2 = objHandleArr
    ActiveSheet.Range(setVerticalRange("B", startRow, nCount)).Value2 = AreaArr
    ActiveSheet.Range(setVerticalRange("C", startRow, nCount)).Value2 = typeArr
    ActiveSheet.Range(setVerticalRange("D", startRow, nCount)).Value2 = XArr
    ActiveSheet.Range(setVerticalRange("E", startRow, nCount)).Value2 = YArr
    ActiveSheet.Range(setVerticalRange("U", startRow, nCount)).Value2 = JxArr
    ActiveSheet.Range(setVerticalRange("V", startRow, nCount)).Value2 = JyArr
    ActiveSheet.Range(setVerticalRange("Y", startRow, nCount)).Value2 = weightArr
    'куда писать габаритные размеры сечения
    ActiveSheet.Cells(10, 35).Value2 = dimX
    ActiveSheet.Cells(11, 35).Value2 = dimY
    
    Set osSet_reg = Nothing
    Set osSet_point = Nothing
    DeleteSSet
    cadDisconnect
    Application.StatusBar = ""
    ActiveSheet.EnableCalculation = True
    Calculate
    Call MsgBox("Перенос геометрии из AutoCAD выполнен. Получено " & nCount & " объектов.", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
'отобразить сжатую зону
Sub ShowCompressed()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    cadConnect
    SetStartValue
    Call SetActiveLayer("Compressed", True, 161)
    Dim i As Integer
    Dim EType As String
    Dim e As Double
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "c" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
        ElseIf EType = "r" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
        ElseIf EType = "u" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
        ElseIf EType = "p" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
        End If
        If e < 0 Then
            Call AddHatch(ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2, "SOLID")
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
cadDisconnect
Call MsgBox("Элементы сжатой зоны показаны синим цветом на слое Compressed", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowDestroyedCompression()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim ebult, esult, euult, erult, Eult As Double
    'нужно ячейки проверять
    ebult = ActiveSheet.Cells(63, 47).Value2 'бетон
    esult = (-1) * ActiveSheet.Cells(72, 47).Value2 'ненапрягаемая арматура
    erult = (-1) * ActiveSheet.Cells(80, 47).Value2 'напрягаемая арматура
    euult = (-1) * ActiveSheet.Cells(88, 47).Value2 'пользовательский материал
    cadConnect
    SetStartValue
    Call SetActiveLayer("Destroyed_Compression", True, 1)
    Dim i As Integer
    Dim objHandle As String
    Dim EType As String
    Dim e As Double
    Dim autocadObj As AcadObject
    Dim hatchObj As AcadHatch
    Dim patternName As String
    Dim patternType As Long
    Dim outerLoop(0 To 0) As AcadObject
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "c" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
            Eult = ebult
        ElseIf EType = "r" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
            Eult = esult
        ElseIf EType = "u" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
            Eult = euult
        ElseIf EType = "p" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
            Eult = erult
        End If
        If e < 0# And e < Eult Then
            Call AddHatch(ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2, "SOLID")
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
   Loop
objDoc.Regen True
cadDisconnect
Call MsgBox("Элементы, относительные деформации сжатия в которых превышают предельные, показаны красным цветом на слое Destroyed_Compression", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowDestroyedTension()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim ebtult, esult, euult, erult, Eult As Double
    'проверить ячейки
    If IsNumeric(ActiveSheet.Cells(64, 47).Value2) Then
       ebtult = ActiveSheet.Cells(64, 47).Value2 'бетон
    Else
       ebtult = 0#
    End If
    esult = ActiveSheet.Cells(72, 47).Value2  'ненапрягаемая арматура
    erult = ActiveSheet.Cells(80, 47).Value2  'напрягаемая арматура
    euult = ActiveSheet.Cells(88, 47).Value2  'пользовательский материал
    cadConnect
    SetStartValue
    Call SetActiveLayer("Destroyed_Tension", True, 2)
    Dim i As Integer
    Dim objHandle As String
    Dim EType As String
    Dim e As Double
    Dim autocadObj As AcadObject
    Dim hatchObj As AcadHatch
    Dim patternName As String
    Dim patternType As Long
    Dim outerLoop(0 To 0) As AcadObject
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "c" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
            Eult = ebtult
        ElseIf EType = "r" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
            Eult = esult
        ElseIf EType = "u" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
            Eult = euult
        ElseIf EType = "p" Then
            e = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
            Eult = erult
        End If
        If e > 0# And e > Eult Then
            Call AddHatch(ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2, "SOLID")
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
objDoc.Regen True
cadDisconnect
Call MsgBox("Элементы, относительные деформации растяжения в которых превышают предельные, показаны жёлтым цветом на слое Destroyed_Tension", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
'показать напряжения в материале
Sub ShowElementStress(TextSize As Double, StressColumn As Integer, typeMat As String)
    SetStartValue
    cadConnect
    Dim i As Integer
    Dim EType As String
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
            If EType = typeMat Then
                Dim StressValue As Double
                StressValue = ActiveSheet.Cells(startRow + i, StartColumn + StressColumn).Value2
                Dim sStressValue As String
                sStressValue = CStr(Format(StressValue, "###0.0"))
                If Len(sStressValue) > 5 Then sStressValue = Left(sStressValue, 5)
                If StressValue >= 0# Then
                    Call SetActiveLayer("Anno_Positive", True, 10)
                Else
                    Call SetActiveLayer("Anno_Negative", True, 4)
                End If
                Call AddText(sStressValue, ActiveSheet.Cells(startRow + i, StartColumn + 4).Value2, _
                    ActiveSheet.Cells(startRow + i, StartColumn + 5).Value2, TextSize)
            End If
            i = i + 1
            EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
    cadDisconnect
End Sub
Sub ShowReinfStress()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    'проверить ячейку
    TextSize = ActiveSheet.Cells(30, 59).Value2
    
    Call ShowElementStress(TextSize, 14, "r")
    Call ShowElementStress(TextSize, 16, "p")
    Call MsgBox("Напряжения в арматуре добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowConcreteStress()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    TextSize = ActiveSheet.Cells(29, 59).Value2
    Call ShowElementStress(TextSize, 13, "c")
    Call MsgBox("Напряжения в бетоне добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowUserStress()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    TextSize = ActiveSheet.Cells(29, 59).Value2
    Call ShowElementStress(TextSize, 15, "u")
    Call MsgBox("Напряжения в пользовательском материале добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
'показать усилия в материале
Sub ShowElementForce(TextSize As Double, StressColumn As Integer, typeMat As String)
SetStartValue
cadConnect
Dim i As Integer
Dim objHandle As String
Dim EType As String
i = 0
EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
Do While EType <> ""
        If EType = typeMat Then
            objHandle = ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2
            Dim StressValue, AreaValue, ForceValue As Double
            StressValue = ActiveSheet.Cells(startRow + i, StartColumn + StressColumn).Value2
            AreaValue = ActiveSheet.Cells(startRow + i, StartColumn + 2).Value2
            'сила в т при напряжениях в кгс/см^2 и площади в см^2
            ForceValue = StressValue * AreaValue / 1000
            Dim sForceValue As String
            sForceValue = CStr(Format(ForceValue, "###0.0"))
            If Len(sForceValue) > 6 Then sForceValue = Left(sForceValue, 6)
            If ForceValue >= 0# Then
                Call SetActiveLayer("Anno_Positive", True, 10)
            Else
                Call SetActiveLayer("Anno_Negative", True, 4)
            End If
            Call AddText(sForceValue, ActiveSheet.Cells(startRow + i, StartColumn + 4).Value2, _
                ActiveSheet.Cells(startRow + i, StartColumn + 5).Value2, TextSize)
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
Loop
objDoc.Regen True
cadDisconnect
End Sub
Sub ShowReinfForce()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    'проверить ссылку на ячейку
    TextSize = ActiveSheet.Cells(30, 59).Value2
    Call ShowElementForce(TextSize, 14, "r")
    Call ShowElementForce(TextSize, 16, "p")
    Call MsgBox("Усилия в арматуре добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowUserForce()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    'проверить ссылку на ячейку
    TextSize = ActiveSheet.Cells(30, 59).Value2
    Call ShowElementForce(TextSize, 15, "u")
    Call MsgBox("Усилия в пользовательском материале добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowPointStress()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize As Double
    'проверить ссылку на ячейку
    TextSize = ActiveSheet.Cells(30, 59).Value2
    
    SetStartValue
    cadConnect
    Dim i As Integer
    Dim EType As String
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "cp" Or EType = "rp" Or EType = "up" Or EType = "pp" Then
            Dim StressValue As Double
            If EType = "cp" Then
                StressValue = ActiveSheet.Cells(startRow + i, StartColumn + 13).Value2
            ElseIf EType = "rp" Then
                StressValue = ActiveSheet.Cells(startRow + i, StartColumn + 14).Value2
            ElseIf EType = "up" Then
                StressValue = ActiveSheet.Cells(startRow + i, StartColumn + 15).Value2
            ElseIf EType = "pp" Then
                StressValue = ActiveSheet.Cells(startRow + i, StartColumn + 16).Value2
            End If
            Dim sStressValue As String
            sStressValue = CStr(Format(StressValue, "###0.0"))
            If Len(sStressValue) > 5 Then sStressValue = Left(sStressValue, 5)
            If StressValue >= 0# Then
                Call SetActiveLayer("Anno_Positive", True, 10)
            Else
                Call SetActiveLayer("Anno_Negative", True, 4)
            End If
            Call AddText(sStressValue, ActiveSheet.Cells(startRow + i, StartColumn + 4).Value2, _
                ActiveSheet.Cells(startRow + i, StartColumn + 5).Value2, TextSize)
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
cadDisconnect
Call MsgBox("Напряжения в точках добавлены на чертёж", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowDeformation()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize, eScale As Double
    'проверить ссылку на ячейку
    TextSize = ActiveSheet.Cells(29, 59).Value2
    eScale = ActiveSheet.Cells(31, 59).Value2
    
    SetStartValue
    cadConnect
    Dim i As Integer
    Dim EType As String
    Dim eValue As Double
    Dim seValue As String
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "c" Or EType = "r" Or EType = "u" Or EType = "p" Then
            If EType = "c" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
            ElseIf EType = "r" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
            ElseIf EType = "u" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
            ElseIf EType = "p" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
            End If
            eValue = eValue * eScale
            seValue = CStr(Format(eValue, "###0.0"))
            If Len(seValue) > 5 Then seValue = Left(seValue, 5)
            If eValue >= 0# Then
                Call SetActiveLayer("Anno_Positive", True, 10)
            Else
                Call SetActiveLayer("Anno_Negative", True, 4)
            End If
            Call AddText(seValue, ActiveSheet.Cells(startRow + i, StartColumn + 4).Value2, _
                ActiveSheet.Cells(startRow + i, StartColumn + 5).Value2, TextSize)
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
cadDisconnect
Call MsgBox("Относительные деформации отображены на чертеже", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowPointDeformation()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim TextSize, eScale As Double
    'проверить ссылку на ячейку
    TextSize = ActiveSheet.Cells(29, 59).Value2
    eScale = ActiveSheet.Cells(31, 59).Value2
    
    SetStartValue
    cadConnect
    Dim i As Integer
    Dim EType As String
    Dim eValue As Double
    Dim seValue As String
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "cp" Or EType = "rp" Or EType = "up" Or EType = "pp" Then
            If EType = "cp" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
            ElseIf EType = "rp" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
            ElseIf EType = "up" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
            ElseIf EType = "pp" Then
                eValue = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
            End If
            eValue = eValue * eScale
            seValue = CStr(Format(eValue, "###0.0"))
            If Len(seValue) > 5 Then seValue = Left(seValue, 5)
            If eValue >= 0# Then
                Call SetActiveLayer("Anno_Positive", True, 10)
            Else
                Call SetActiveLayer("Anno_Negative", True, 4)
            End If
            Call AddText(seValue, ActiveSheet.Cells(startRow + i, StartColumn + 4).Value2, _
                ActiveSheet.Cells(startRow + i, StartColumn + 5).Value2, TextSize)
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
cadDisconnect
Call MsgBox("Относительные деформации отображены на чертеже", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
'показать напряжения столбиками
Sub ShowColumnsStress(StressScale As Double, StressColumn As Integer, typeMat As String)
SetStartValue
cadConnect
Dim i As Integer
Dim objHandle As String
Dim EType As String
Dim Stress As Double
i = 0
EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
Do While EType <> ""
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    If EType = typeMat Then
        Stress = ActiveSheet.Cells(startRow + i, StartColumn + StressColumn).Value2
        If Stress <> 0 Then
            If Stress < 0 Then
                Call SetActiveLayer("Columns_Stress_Comp", True, 161)
            Else
                Call SetActiveLayer("Columns_Stress_Tens", True, 20)
            End If
            Call Add3D(ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2, Stress * StressScale)
        End If
    End If
    i = i + 1
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
Loop
objDoc.Regen True
cadDisconnect
End Sub
Sub ShowColumnsStressConcrete()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim StressScale As Double
    'проверить ссылку на ячейку
    StressScale = ActiveSheet.Cells(27, 59).Value2
    
    Call ShowColumnsStress(StressScale, 13, "c")
    Call MsgBox("Напряжения в бетоне отображены объектами на слоях Columns_Stress_Comp и Columns_Stress_Tens", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowColumnsStressUser()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim StressScale As Double
    'проверить ссылку на ячейку
    StressScale = ActiveSheet.Cells(27, 59).Value2
    
    Call ShowColumnsStress(StressScale, 15, "u")
    Call MsgBox("Напряжения в пользовательском материале отображены объектами на слоях Columns_Stress_Comp и Columns_Stress_Tens", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ShowColumnsStrain()
    On Error GoTo ErrorHandler
    Application.EnableCancelKey = xlErrorHandler
    Dim StrainScale As Double
    'проверить ссылку на ячейку
    StrainScale = ActiveSheet.Cells(26, 59).Value2

    SetStartValue
    cadConnect
    Call SetActiveLayer("Columns_strain", True, 140)
    Dim i As Integer
    Dim objHandle As String
    Dim EType As String
    Dim Strain, Strain_r, Strain_c As Double
    i = 0
    EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Do While EType <> ""
        If EType = "c" Or EType = "r" Or EType = "u" Or EType = "p" Then
            objHandle = ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2
            If EType = "c" Then
                Strain = ActiveSheet.Cells(startRow + i, StartColumn + 9).Value2
            ElseIf EType = "r" Then
                Strain = ActiveSheet.Cells(startRow + i, StartColumn + 10).Value2
            ElseIf EType = "u" Then
                Strain = ActiveSheet.Cells(startRow + i, StartColumn + 26).Value2
            ElseIf EType = "p" Then
                Strain = ActiveSheet.Cells(startRow + i, StartColumn + 27).Value2
            End If
            If Strain <> 0 Then
                Call Add3D(ActiveSheet.Cells(startRow + i, StartColumn + 1).Value2, Strain * StrainScale)
            End If
        End If
        i = i + 1
        EType = ActiveSheet.Cells(startRow + i, StartColumn + 3).Value2
    Loop
    objDoc.Regen True
cadDisconnect
Call MsgBox("Относительные деформации отображены объектами на слое Columns_strain", vbInformation, "Завершено")
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub
Sub ClearDWG()
On Error GoTo ErrorHandler
Application.EnableCancelKey = xlErrorHandler
cadConnect

Call ClearLayer("Hatch", "Compressed")
Call ClearLayer("Hatch", "Destroyed_Compression")
Call ClearLayer("Hatch", "Destroyed_Tension")

Call ClearLayer("Text", "Anno")
Call ClearLayer("Text", "Anno_Positive")
Call ClearLayer("Text", "Anno_Negative")

Call ClearLayer("3DSolid", "Columns_Stress_Comp")
Call ClearLayer("3DSolid", "Columns_Stress_Tens")
Call ClearLayer("3DSolid", "Columns_Strain")

objDoc.Regen True
cadDisconnect
ErrorHandler:
    ErrorHandlerSub (Err.Number)
End Sub


