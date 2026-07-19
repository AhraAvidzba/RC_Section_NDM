Attribute VB_Name = "Functions"
Option Explicit
Function StressStrainConcrete(e, tension, concDiag, concClass, eb0, eb1, eb1red, ebt0, ebt1, ebt1red, ebt2, Eb, Ebint, Rb, Rbt, h)

If concDiag = "двухлинейная" Then
        StressStrainConcrete = Concrete2Linear(e, tension, Eb, Rb, Rbt, eb1red, ebt1red, ebt2)

    ElseIf concDiag = "трёхлинейная" Then
        StressStrainConcrete = Concrete3Linear(e, tension, Eb, Rb, Rbt, eb0, eb1, ebt0, ebt1, ebt2)

    ElseIf concDiag = "криволинейная" Then
        StressStrainConcrete = ConcreteCurved(e, tension, Ebint, Eb, Rb, Rbt, concClass, h)
End If

End Function

Function StressStrainReinf(e, reinfDiag, es0, es2, Es, Rs, Rsc)

If reinfDiag = "Двухлинейная (физический предел текучести)" Then
            StressStrainReinf = Reinf2Linear(e, Es, Rs, Rsc, es2, 0, 1)

        ElseIf reinfDiag = "Трехлинейная (условный предел текучести)" Then
            StressStrainReinf = Reinf3Linear(e, Es, Rs, Rsc, es0, es2, 0, 1)

End If

End Function
Function StressStrainPreReinf(e, reinfDiag, es0, es2, Es, Rs, Rsc, PreStress, gammaSP)

If reinfDiag = "Двухлинейная (физический предел текучести)" Then
            StressStrainPreReinf = Reinf2Linear(e, Es, Rs, Rsc, es2, PreStress, gammaSP)

        ElseIf reinfDiag = "Трехлинейная (условный предел текучести)" Then
            StressStrainPreReinf = Reinf3Linear(e, Es, Rs, Rsc, es0, es2, PreStress, gammaSP)

End If

End Function
Private Function Concrete2Linear(ByVal e As Double, ByVal tension As String, _
        ByVal Eb As Double, ByVal Rb As Double, ByVal Rbt As Double, ByVal eb1red As Double, _
        ByVal ebt1red As Double, ByVal ebt2 As Double) As Double
        
    Dim Ebred, Ebtred As Double
    
    Ebred = Rb / eb1red
    Ebtred = Rbt / ebt1red
        
    If e > 0 Then
        If tension = "нет" Then
            Concrete2Linear = 0
        Else
            If Abs(e) <= ebt1red Then
                Concrete2Linear = Ebred * e
            Else
                If Abs(e) <= ebt2 Then
                    Concrete2Linear = Rbt
                Else
                    Concrete2Linear = 0
                End If
            End If
        End If
    Else
        If Abs(e) <= eb1red Then
            Concrete2Linear = Ebred * e
        Else
            Concrete2Linear = -Rb
        End If
    End If
End Function
Private Function Concrete3Linear(ByVal e As Double, ByVal tension As String, _
        ByVal Eb As Double, ByVal Rb As Double, ByVal Rbt As Double, ByVal eb0 As Double, _
        ByVal eb1 As Double, ByVal ebt0 As Double, ByVal ebt1 As Double, ByVal ebt2 As Double) As Double
        
    If e > 0 Then
        If tension = "нет" Then
            Concrete3Linear = 0
        Else
            If Abs(e) <= ebt1 Then
                Concrete3Linear = Eb * e
            Else
                If Abs(e) <= ebt0 Then
                    Concrete3Linear = ((1 - (0.6 * Rbt / Rbt)) * ((Abs(e) - ebt1) / (ebt0 - ebt1)) + (0.6 * Rbt / Rbt)) * Rbt
                Else
                    If Abs(e) <= ebt2 Then
                        Concrete3Linear = Rbt
                    Else
                        Concrete3Linear = 0
                    End If
                End If
            End If
        End If
    Else
        If Abs(e) <= eb1 Then
            Concrete3Linear = Eb * e
        Else
            If Abs(e) <= eb0 Then
                Concrete3Linear = -((1 - (0.6 * Rb / Rb)) * ((Abs(e) - eb1) / (eb0 - eb1)) + (0.6 * Rb / Rb)) * Rb
            Else
                Concrete3Linear = -Rb
            End If
        End If
    End If
End Function
Private Function ConcreteCurved(ByVal e As Double, ByVal tension As String, ByVal E0 As Double, _
        ByVal Eb As Double, ByVal Rb As Double, ByVal Rbt As Double, ByVal concClass As String, ByVal h As Double) As Double

Dim Sigma_, e_, Lambda As Double
Dim w1, w2, v_, v0 As Double
Dim eE As Double
Dim M1, m2, m3 As Double
Dim Class As Double
Dim gamma_btq, h_e As Double
h_e = 30

Class = CDbl(Right(concClass, 2))

If e < 0 Then

    Lambda = 1
    e_ = -(Class / E0) * Lambda * (1 + 0.75 * Lambda * Class / 60 + 0.2 * Lambda / Class) / (0.12 + Class / 60 + 0.2 / Class)
    Sigma_ = -Rb

    v_ = Sigma_ / (e_ * E0)
    eE = e * E0

    If Abs(e) <= Abs(e_) Then
        v0 = 1
        w1 = 2 - 2.5 * v_
        w2 = 1 - w1

        M1 = (w1 * eE ^ 2 * v0 ^ 2 - 2 * w1 * eE ^ 2 * v0 * v_ + w1 * eE ^ 2 * v_ ^ 2 - 2 * Sigma_ * eE * v_) / (2 * Sigma_)
    
        m2 = (eE * (v0 - v_) * Math.Sqr(eE ^ 2 * v0 ^ 2 * w1 ^ 2 + 4 * w2 * eE ^ 2 * v0 ^ 2 - 2 * eE ^ 2 * v0 * v_ * w1 ^ 2 - _
            8 * w2 * eE ^ 2 * v0 * v_ + eE ^ 2 * v_ ^ 2 * w1 ^ 2 - 4 * eE * Sigma_ * v_ * w1 + 4 * Sigma_ ^ 2)) / (2 * Sigma_)
        
        m3 = w2 * eE ^ 2 * v0 ^ 2 - 2 * w2 * eE ^ 2 * v0 * v_ + w2 * eE ^ 2 * v_ ^ 2 + Sigma_ ^ 2
        
        ConcreteCurved = (-1) * Sigma_ ^ 2 * (M1 + m2) / m3
    Else
        v0 = 2.05 * v_
        w1 = 1.95 * v_ - 0.138
        w2 = 1 - w1
    
        M1 = (w1 * eE ^ 2 * v0 ^ 2 - 2 * w1 * eE ^ 2 * v0 * v_ + w1 * eE ^ 2 * v_ ^ 2 - 2 * Sigma_ * eE * v_) / (2 * Sigma_)
    
        m2 = (eE * (v0 - v_) * Math.Sqr(eE ^ 2 * v0 ^ 2 * w1 ^ 2 + 4 * w2 * eE ^ 2 * v0 ^ 2 - 2 * eE ^ 2 * v0 * v_ * w1 ^ 2 - _
        8 * w2 * eE ^ 2 * v0 * v_ + eE ^ 2 * v_ ^ 2 * w1 ^ 2 - 4 * eE * Sigma_ * v_ * w1 + 4 * Sigma_ ^ 2)) / (2 * Sigma_)
    
        m3 = w2 * eE ^ 2 * v0 ^ 2 - 2 * w2 * eE ^ 2 * v0 * v_ + w2 * eE ^ 2 * v_ ^ 2 + Sigma_ ^ 2
        
        ConcreteCurved = (-1) * Sigma_ ^ 2 * (M1 + m2) / m3
        
        If ConcreteCurved / Sigma_ < 0.85 Then ConcreteCurved = 0
    End If
    
Else
    If tension = "Нет" Then
        ConcreteCurved = 0
    Else
        h = h / 10
        
        gamma_btq = 2.07 - ((h / h_e) ^ (1 / 5))
        
        If gamma_btq < 0.9 Then gamma_btq = 0.9
    
        Sigma_ = Rbt * gamma_btq
        v_ = (0.6 + (0.15 * (Rbt / 2.5))) / gamma_btq
    
        e_ = (Sigma_) / (E0 * v_)
        eE = e * E0
    
        If e <= e_ Then
            v0 = 1
            w1 = 2 - 2.5 * v_
            w2 = 1 - w1

            M1 = (w1 * eE ^ 2 * v0 ^ 2 - 2 * w1 * eE ^ 2 * v0 * v_ + w1 * eE ^ 2 * v_ ^ 2 - 2 * Sigma_ * eE * v_) / (2 * Sigma_)
    
            m2 = (eE * (v0 - v_) * Math.Sqr(eE ^ 2 * v0 ^ 2 * w1 ^ 2 + 4 * w2 * eE ^ 2 * v0 ^ 2 - 2 * eE ^ 2 * v0 * v_ * w1 ^ 2 - _
                8 * w2 * eE ^ 2 * v0 * v_ + eE ^ 2 * v_ ^ 2 * w1 ^ 2 - 4 * eE * Sigma_ * v_ * w1 + 4 * Sigma_ ^ 2)) / (2 * Sigma_)
        
            m3 = w2 * eE ^ 2 * v0 ^ 2 - 2 * w2 * eE ^ 2 * v0 * v_ + w2 * eE ^ 2 * v_ ^ 2 + Sigma_ ^ 2
        
            ConcreteCurved = (-1) * Sigma_ ^ 2 * (M1 - m2) / m3
        Else
            v0 = 2.05 * v_
            w1 = 1.95 * v_ - 0.138
            w2 = 1 - w1
    
            M1 = (w1 * eE ^ 2 * v0 ^ 2 - 2 * w1 * eE ^ 2 * v0 * v_ + w1 * eE ^ 2 * v_ ^ 2 - 2 * Sigma_ * eE * v_) / (2 * Sigma_)
    
            m2 = (eE * (v0 - v_) * Math.Sqr(eE ^ 2 * v0 ^ 2 * w1 ^ 2 + 4 * w2 * eE ^ 2 * v0 ^ 2 - 2 * eE ^ 2 * v0 * v_ * w1 ^ 2 - _
            8 * w2 * eE ^ 2 * v0 * v_ + eE ^ 2 * v_ ^ 2 * w1 ^ 2 - 4 * eE * Sigma_ * v_ * w1 + 4 * Sigma_ ^ 2)) / (2 * Sigma_)
    
            m3 = w2 * eE ^ 2 * v0 ^ 2 - 2 * w2 * eE ^ 2 * v0 * v_ + w2 * eE ^ 2 * v_ ^ 2 + Sigma_ ^ 2
        
            ConcreteCurved = (-1) * Sigma_ ^ 2 * (M1 - m2) / m3
        
            If ConcreteCurved / Sigma_ < 0.85 Then ConcreteCurved = 0
            
        End If
        
    End If
    
End If

End Function
Function Reinf2Linear(ByVal e As Double, _
        ByVal Es As Double, ByVal Rs As Double, ByVal Rsc As Double, ByVal es2 As Double, _
        ByVal PreStress As Double, ByVal gammaSP As Double) As Double
        
Dim es0 As Double
Dim e_calc As Double

e_calc = e + (PreStress * gammaSP) / Es
        
If e_calc < 0 Then
    es0 = Rsc / Es
    If Abs(e_calc) < Abs(es0) Then
        Reinf2Linear = e_calc * Es
    Else
        Reinf2Linear = -Rsc
    End If
Else
    es0 = Rs / Es
    If e_calc < es0 Then
        Reinf2Linear = e_calc * Es
    Else
        Reinf2Linear = Rs
    End If
End If

End Function
Private Function Reinf3Linear(ByVal e As Double, _
        ByVal Es As Double, ByVal Rs As Double, ByVal Rsc As Double, ByVal es0 As Double, ByVal es2 As Double, _
        ByVal PreStress As Double, ByVal gammaSP As Double) As Double
        
Dim es1 As Double
Dim Sigma_s1, Sigma_s2 As Double
Dim eRed As Double
Dim e_calc As Double

e_calc = e + (PreStress * gammaSP) / Es

If e_calc < 0 Then

    es1 = 0.9 * Rsc / Es

    If Abs(e_calc) < Abs(es1) Then
        Reinf3Linear = e_calc * Es
    Else
        Sigma_s1 = 0.9 * Rsc
        eRed = ((1 - (Sigma_s1 / Rsc)) * ((Abs(e_calc) - es1) / (es0 - es1))) + (Sigma_s1 / Rsc)
        Reinf3Linear = -eRed * Rsc
        If Abs(Reinf3Linear) > Abs(1.1 * Rsc) Then Reinf3Linear = -1.1 * Rsc
    End If
Else

    es1 = 0.9 * Rs / Es

    If e_calc < es1 Then
        Reinf3Linear = e_calc * Es
    Else
        Sigma_s1 = 0.9 * Rs
        eRed = ((1 - (Sigma_s1 / Rs)) * ((e_calc - es1) / (es0 - es1))) + (Sigma_s1 / Rs)
        Reinf3Linear = eRed * Rs
        If Reinf3Linear > 1.1 * Rs Then Reinf3Linear = 1.1 * Rs
    End If
End If

End Function
Function User2Linear(ByVal e As Double, _
        ByVal Es As Double, ByVal Rs As Double, ByVal Rsc As Double, PreStress As Double, gammaUP) As Double

Dim es0, e_calc As Double

e_calc = e + (PreStress * gammaUP) / Es

If e_calc < 0 Then
    es0 = Rsc / Es
    
    If Abs(e_calc) < Abs(es0) Then
        User2Linear = e_calc * Es
    Else
        User2Linear = -Rsc
    End If

Else
    es0 = Rs / Es
    
    If e_calc < es0 Then
        User2Linear = e_calc * Es
    Else
        User2Linear = Rs
    End If

End If

End Function

Private Function Interpolation(ByVal X As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

    Interpolation = y1 + ((y2 - y1) / (x2 - x1)) * X

End Function
Private Function CheckTolerance(ByVal X As Double, ByVal xn As Double, ByVal tolerance_percent As Double, ByVal tolerance_value As Double) As Boolean

    CheckTolerance = False

    If Abs(xn) < 15 Then
        If Abs(X - xn) < tolerance_value Then CheckTolerance = True
    Else
        If Abs((X - xn) / Abs(xn)) < tolerance_percent Then CheckTolerance = True
    End If

End Function
Sub Iteration()
On Error GoTo ErrorHandler
Application.DisplayStatusBar = True
Application.EnableCancelKey = xlErrorHandler
Dim MaxIteration, iterCode As Integer
'проверять ссылку на ячейку
MaxIteration = ActiveSheet.Cells(45, 44).Value2
iterCode = MainIteration(ByVal MaxIteration, ByVal True, ByVal True)
ErrorHandler:
    If Err.Number <> 0 And Err.Number <> 40040 Then
    Select Case Err.Number
        Case 18
            Call MsgBox("Выполнение расчёта остановлено пользователем", vbInformation, "Завершение")
            Calculate
            Application.StatusBar = ""
        Case Else
            Call MsgBox("При выполнении возникла ошибка номер " & Err.Number, vbCritical, "Ошибка")
    End Select
    End If
Call ClearDWG
End Sub

Sub BatchIteration()

Dim sTerm As String
Dim gamma_b, gamma_s, gamma_u, gamma_sp As Double
Dim dDlit, N, Mx, My As Double
Dim i, j, StartString As Integer
Dim iterCode, MaxIteration As Integer
Dim customValueInput(0 To 2) As String
Dim customValueOutput(0 To 2) As String
Dim resultRange As String

On Error GoTo ErrorHandler

Application.EnableCancelKey = xlErrorHandler

resultRange = "N3:Q1002"

ActiveWorkbook.Worksheets("Batch").Range(resultRange).ClearContents
ActiveWorkbook.Worksheets("Batch").Range(resultRange).Font.Bold = False
ActiveWorkbook.Worksheets("Batch").Range(resultRange).Font.color = RGB(0, 0, 0)

StartString = 3

For j = 0 To 2
    
    customValueInput(j) = ""
    customValueOutput(j) = ""
    
    If ActiveWorkbook.Worksheets("Batch").Cells(2, 11 + j).Value2 <> "" Then
        
        customValueInput(j) = ActiveWorkbook.Worksheets("Batch").Cells(2, 11 + j).Value2
        
    End If
    
    If ActiveWorkbook.Worksheets("Batch").Cells(2, 15 + j).Value2 <> "" Then
    
        customValueOutput(j) = ActiveWorkbook.Worksheets("Batch").Cells(2, 15 + j).Value2
        
    End If
    
Next

MaxIteration = ActiveWorkbook.Worksheets("Calculation").Cells(5, 4).Value2

i = 0

sTerm = CStr(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 2).Value2)

gamma_b = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 3).Value2)
gamma_s = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 4).Value2)
gamma_sp = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 5).Value2)
gamma_u = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 6).Value2)

dDlit = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 7).Value2)

N = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 8).Value2
Mx = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 9).Value2
My = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 10).Value2

Do While dDlit <> 0 Or N <> 0 Or Mx <> 0 Or My <> 0

    If sTerm <> "" Then
        ActiveWorkbook.Worksheets("Calculation").Cells(10, 4).Value2 = sTerm
    End If
    
    If gamma_b <> 0# Then
        ActiveWorkbook.Worksheets("Calculation").Cells(11, 4).Value2 = gamma_b
        ActiveWorkbook.Worksheets("Calculation").Cells(12, 4).Value2 = gamma_b
    End If
    
    If gamma_s <> 0# Then
        ActiveWorkbook.Worksheets("Calculation").Cells(28, 4).Value2 = gamma_s
    End If
    
    If gamma_sp <> 0# Then
        ActiveWorkbook.Worksheets("Calculation").Cells(37, 4).Value2 = gamma_sp
    End If
    
    If gamma_u <> 0# Then
        ActiveWorkbook.Worksheets("Calculation").Cells(46, 4).Value2 = gamma_u
    End If

    If dDlit <> 0# Then
        ActiveWorkbook.Worksheets("Calculation").Cells(60, 4).Value2 = dDlit
    End If

    ActiveWorkbook.Worksheets("Calculation").Cells(63, 2).Value2 = N
    ActiveWorkbook.Worksheets("Calculation").Cells(64, 2).Value2 = Mx
    ActiveWorkbook.Worksheets("Calculation").Cells(65, 2).Value2 = My

    For j = 0 To 2

        If customValueInput(j) <> "" Then
        
            If ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 11 + j).Value2 <> "" Then
        
                ActiveWorkbook.Worksheets("Calculation").Range(customValueInput(j)) = _
                    ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 11 + j).Value2
                
            End If

        End If

    Next

    iterCode = FunctionIteration(ByVal MaxIteration, ByVal False, ByVal False)

    If iterCode = 0 Then

        ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Value2 = _
            ActiveWorkbook.Worksheets("Calculation").Cells(155, 4).Value2

        If ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Value2 > 1 Then

            ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Font.Bold = True
            ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Font.color = RGB(250, 0, 0)

        End If


        For j = 0 To 2
            If customValueOutput(j) <> "" Then
                ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 15 + j).Value2 = _
                    ActiveWorkbook.Worksheets("Calculation").Range(customValueOutput(j))
            End If
        Next

    ElseIf iterCode = 1 Or iterCode = 2 Then

        ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Value2 = "Решение не найдено; k>>1"
        ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Font.Bold = True
        ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 14).Font.color = RGB(250, 0, 0)

        For j = 0 To 2
            If customValueOutput(j) <> "" Then
                ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 15 + j).Value2 = "-"
                ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 15 + j).Font.Bold = True
                ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 15 + j).Font.color = RGB(250, 0, 0)
            End If
        Next

    End If

    i = i + 1

    sTerm = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 2).Value2

    gamma_b = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 3).Value2
    gamma_s = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 4).Value2
    gamma_sp = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 5).Value2
    gamma_u = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 6).Value2

    dDlit = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 7).Value2

    N = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 8).Value2
    Mx = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 9).Value2
    My = ActiveWorkbook.Worksheets("Batch").Cells(StartString + i, 10).Value2

Loop

ErrorHandler:

    If Err.Number <> 0 And Err.Number <> 40040 Then

        Select Case Err.Number
            Case 18
                Call MsgBox("Выполнение расчета остановлено пользователем", vbInformation, "Завершение")
                Calculate
                Application.StatusBar = ""
            Case 1004
                Call MsgBox("Неправильно указан адрес ячейки входного или выходного параметра", vbCritical, "Ошибка")
            Case Else
                Call MsgBox("При выполнении возникла ошибка номер " & Err.Number, vbCritical, "Ошибка")
        End Select

    End If

End Sub

Sub UseCurrent()

Dim sTerm As String
Dim gamma_b, gamma_s, gamma_u, gamma_sp As Double
Dim dDlit, N, Mx, My As Double
Dim j, aRow As Integer
Dim customValueInput(0 To 2) As String

On Error GoTo ErrorHandler


aRow = ActiveCell.Row


For j = 0 To 2
    
    customValueInput(j) = ""
    
    If ActiveWorkbook.Worksheets("Batch").Cells(2, 11 + j).Value2 <> "" Then
        
        customValueInput(j) = ActiveWorkbook.Worksheets("Batch").Cells(2, 11 + j).Value2
        
    End If
    
Next

sTerm = CStr(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 2).Value2)

gamma_b = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 3).Value2)
gamma_s = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 4).Value2)
gamma_sp = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 5).Value2)
gamma_u = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 6).Value2)

dDlit = CDbl(ActiveWorkbook.Worksheets("Batch").Cells(aRow, 7).Value2)

If sTerm <> "" Then
    ActiveWorkbook.Worksheets("Calculation").Cells(10, 4).Value2 = sTerm
End If
    
If gamma_b <> 0# Then
    ActiveWorkbook.Worksheets("Calculation").Cells(11, 4).Value2 = gamma_b
    ActiveWorkbook.Worksheets("Calculation").Cells(12, 4).Value2 = gamma_b
End If
    
If gamma_s <> 0# Then
    ActiveWorkbook.Worksheets("Calculation").Cells(28, 4).Value2 = gamma_s
End If
    
If gamma_sp <> 0# Then
    ActiveWorkbook.Worksheets("Calculation").Cells(37, 4).Value2 = gamma_sp
End If
    
If gamma_u <> 0# Then
    ActiveWorkbook.Worksheets("Calculation").Cells(46, 4).Value2 = gamma_u
End If

If dDlit <> 0# Then
    ActiveWorkbook.Worksheets("Calculation").Cells(60, 4).Value2 = dDlit
End If

N = ActiveWorkbook.Worksheets("Batch").Cells(aRow, 8).Value2
Mx = ActiveWorkbook.Worksheets("Batch").Cells(aRow, 9).Value2
My = ActiveWorkbook.Worksheets("Batch").Cells(aRow, 10).Value2

If N = 0 And Mx = 0 And My = 0 Then

    Call MsgBox("Текущая строка с пустыми параметрами", vbCritical, "Ошибка")
    Exit Sub

Else

    ActiveWorkbook.Worksheets("Calculation").Cells(63, 2).Value2 = N
    ActiveWorkbook.Worksheets("Calculation").Cells(64, 2).Value2 = Mx
    ActiveWorkbook.Worksheets("Calculation").Cells(65, 2).Value2 = My
    
    For j = 0 To 2
    
        If customValueInput(j) <> "" Then
        
            If ActiveWorkbook.Worksheets("Batch").Cells(aRow, 11 + j).Value2 <> "" Then
    
                ActiveWorkbook.Worksheets("Calculation").Range(customValueInput(j)) = _
                    ActiveWorkbook.Worksheets("Batch").Cells(aRow, 11 + j).Value2
                
            End If
    
        End If
    
    Next
    
    ActiveWorkbook.Worksheets("Calculation").Activate
End If

ErrorHandler:

    If Err.Number <> 0 And Err.Number <> 40040 Then

        Select Case Err.Number
            Case 1004
                Call MsgBox("Неправильно указан адрес ячейки входного или выходного параметра", vbCritical, "Ошибка")
            Case Else
                Call MsgBox("При выполнении возникла ошибка номер " & Err.Number, vbCritical, "Ошибка")
        End Select

    End If


End Sub
'основная функция итерации
Function FunctionIteration(ByVal MaxIteration As Integer, _
            ByVal ShowDialog As Boolean, ByVal StatusBar As Boolean) As Integer

Dim N, Mx, My As Double

N = ActiveWorkbook.Worksheets("Calculation").Cells(63, 3).Value2
Mx = ActiveWorkbook.Worksheets("Calculation").Cells(64, 3).Value2
My = ActiveWorkbook.Worksheets("Calculation").Cells(65, 3).Value2

Dim tolerance_percent, tolerance_value As Double

tolerance_percent = ActiveWorkbook.Worksheets("Calculation").Cells(3, 4).Value2
tolerance_value = ActiveWorkbook.Worksheets("Calculation").Cells(3, 5).Value2

Dim rx, ry, E0 As Double
Dim Ncalc, Mxcalc, Mycalc As Double

Dim D11, D22, D12, D13, D23, D33 As Double
Dim i As Integer
Dim sFini As String

FunctionIteration = 0

i = 1

D11 = ActiveWorkbook.Worksheets("Calculation").Cells(113, 4).Value2
D22 = ActiveWorkbook.Worksheets("Calculation").Cells(114, 4).Value2
D12 = ActiveWorkbook.Worksheets("Calculation").Cells(115, 4).Value2
D13 = ActiveWorkbook.Worksheets("Calculation").Cells(116, 4).Value2
D23 = ActiveWorkbook.Worksheets("Calculation").Cells(117, 4).Value2
D33 = ActiveWorkbook.Worksheets("Calculation").Cells(118, 4).Value2

rx = (D23 ^ 2 * Mx + D12 * D33 * My - D13 * D23 * My - D22 * D33 * Mx - D12 * D23 * N + D13 * D22 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

ry = (D13 ^ 2 * My - D11 * D33 * My + D12 * D33 * Mx - D13 * D23 * Mx + D11 * D23 * N - D12 * D13 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

E0 = (D12 ^ 2 * N + D11 * D23 * My - D12 * D13 * My - D12 * D23 * Mx + D13 * D22 * Mx - D11 * D22 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

ActiveWorkbook.Worksheets("Calculation").Cells(109, 5).Value2 = rx
ActiveWorkbook.Worksheets("Calculation").Cells(110, 5).Value2 = ry
ActiveWorkbook.Worksheets("Calculation").Cells(111, 5).Value2 = E0

Calculate

Ncalc = ActiveWorkbook.Worksheets("Calculation").Cells(63, 4).Value2
Mxcalc = ActiveWorkbook.Worksheets("Calculation").Cells(64, 4).Value2
Mycalc = ActiveWorkbook.Worksheets("Calculation").Cells(65, 4).Value2


Dim NTolerance, MxTolerance, MyTolerance As Boolean

NTolerance = CheckTolerance(Ncalc, N, tolerance_percent, tolerance_value)
MxTolerance = CheckTolerance(Mxcalc, Mx, tolerance_percent, tolerance_value)
MyTolerance = CheckTolerance(Mycalc, My, tolerance_percent, tolerance_value)

Do While NTolerance = False Or _
    MxTolerance = False Or _
    MyTolerance = False
    
    If StatusBar = True Then
        Application.StatusBar = "Выполняется итерация " & i
    End If

    i = i + 1
    
    If i = MaxIteration Then
        If ShowDialog = True Then
            Dim answer As Integer
            Dim sQuestion As String
            
            sQuestion = "Было выполнено " & i & " итераций, а решение не было найдено. Это может означать, что при введенных усилиях несущая способность сечения не обеспечена и решение уравнений не может быть найдено либо введены очень жесткие допуски. Остановить расчет?"
            
            answer = MsgBox(sQuestion, vbYesNo + vbQuestion, "Вопрос")
            If answer = vbYes Then
                
                If StatusBar = True Then
                    Application.StatusBar = ""
                End If
                
                Exit Function
            End If
        Else
            FunctionIteration = 1
            Exit Function
        End If
    End If
    
    D11 = ActiveWorkbook.Worksheets("Calculation").Cells(113, 5).Value2
    D22 = ActiveWorkbook.Worksheets("Calculation").Cells(114, 5).Value2
    D12 = ActiveWorkbook.Worksheets("Calculation").Cells(115, 5).Value2
    D13 = ActiveWorkbook.Worksheets("Calculation").Cells(116, 5).Value2
    D23 = ActiveWorkbook.Worksheets("Calculation").Cells(117, 5).Value2
    D33 = ActiveWorkbook.Worksheets("Calculation").Cells(118, 5).Value2
        
    rx = (D23 ^ 2 * Mx + D12 * D33 * My - D13 * D23 * My - D22 * D33 * Mx - D12 * D23 * N + D13 * D22 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

    ry = (D13 ^ 2 * My - D11 * D33 * My + D12 * D33 * Mx - D13 * D23 * Mx + D11 * D23 * N - D12 * D13 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

    E0 = (D12 ^ 2 * N + D11 * D23 * My - D12 * D13 * My - D12 * D23 * Mx + D13 * D22 * Mx - D11 * D22 * N) / _
        (D33 * D12 ^ 2 - 2 * D12 * D13 * D23 + D22 * D13 ^ 2 + D11 * D23 ^ 2 - D11 * D22 * D33)

    If Abs(rx) > 1 Or Abs(ry) > 1 Or Abs(E0) > 1 Then
        If ShowDialog = True Then
            MsgBox "Задача нелинейных уравнений расходится. Как правило, это означает, что при введенных усилиях несущая способность сечения не обеспечена и решение уравнений не может быть найдено. Необходимо уменьшить величины усилий или увеличить допуск на расчет"
            
            If StatusBar = True Then
                Application.StatusBar = ""
            End If
            
            Exit Function
        Else
            FunctionIteration = 2
            Exit Function
        End If
    End If
    
    ActiveWorkbook.Worksheets("Calculation").Cells(109, 5).Value2 = rx
    ActiveWorkbook.Worksheets("Calculation").Cells(110, 5).Value2 = ry
    ActiveWorkbook.Worksheets("Calculation").Cells(111, 5).Value2 = E0

    Calculate
    
    N = ActiveWorkbook.Worksheets("Calculation").Cells(63, 3).Value2
    Mx = ActiveWorkbook.Worksheets("Calculation").Cells(64, 3).Value2
    My = ActiveWorkbook.Worksheets("Calculation").Cells(65, 3).Value2
    
    Ncalc = ActiveWorkbook.Worksheets("Calculation").Cells(63, 4).Value2
    Mxcalc = ActiveWorkbook.Worksheets("Calculation").Cells(64, 4).Value2
    Mycalc = ActiveWorkbook.Worksheets("Calculation").Cells(65, 4).Value2
    
    NTolerance = CheckTolerance(Ncalc, N, tolerance_percent, tolerance_value)
    MxTolerance = CheckTolerance(Mxcalc, Mx, tolerance_percent, tolerance_value)
    MyTolerance = CheckTolerance(Mycalc, My, tolerance_percent, tolerance_value)
    
Loop

If ShowDialog = True Then

    sFini = "Расчет окончен, выполнено " & i & " итераций. Требуемые параметры точности достигнуты."

    Call MsgBox(sFini, vbInformation, "Завершено")

End If

If StatusBar = True Then

    Application.StatusBar = ""
    
End If

End Function
'решение матричного уравнения: Md-матрица коэффициентов, V-вектор свободных членов
Function Fmatrix(Md As Range, V As Range)
  Fmatrix = WorksheetFunction.MMult(WorksheetFunction.MInverse(Md), V)
End Function
'диаграмма деформирования бетона
Function SigmaEpsConcrete(EPS#, Ecomp As Range, Etens As Range, Eb!, Rb!, Rbt!, concDiag As Variant, tensflag As Variant)
Dim Ebred!, epsb1#, Ebtred!, epsbt1#, sigmab1!, sigmabt1!
Select Case concDiag
'***************************************
  Case "двухлинейная"
  Ebred = Abs(Rb / Ecomp(2))
  epsb1 = -Rb / Ebred
  If EPS <= 0 Then 'при сжатии
     If EPS >= epsb1 Then
        SigmaEpsConcrete = Ebred * EPS
     Else
        SigmaEpsConcrete = -Rb
     End If
  Else 'при растяжении
     If tensflag = "нет" Then
        SigmaEpsConcrete = 0#
     Else
        Ebtred = Rbt / Etens(2)
        epsbt1 = Rbt / Ebtred
        If EPS <= epsbt1 Then
           SigmaEpsConcrete = Ebtred * EPS
        Else
           If EPS > epsbt1 And EPS <= Etens(4) Then
              SigmaEpsConcrete = Rbt
           Else
              SigmaEpsConcrete = 0#
           End If
        End If
     End If
  End If
'***************************************
  Case "трёхлинейная"
  sigmab1 = -0.6 * Rb
  epsb1 = sigmab1 / Eb
  If EPS <= 0 Then 'при сжатии
     If EPS >= epsb1 Then
        SigmaEpsConcrete = EPS * Eb
     Else
        If EPS < epsb1 And EPS > Ecomp(3) Then
           SigmaEpsConcrete = -((1 - Abs(sigmab1) / Rb) * (Abs(EPS) - Abs(epsb1)) / (Abs(Ecomp(3)) - Abs(epsb1)) + Abs(sigmab1) / Rb) * Rb
        Else
           SigmaEpsConcrete = -Rb
        End If
     End If
  Else 'при растяжении
     If tensflag = "нет" Then
        SigmaEpsConcrete = 0#
     Else
        sigmabt1 = 0.6 * Rbt
        epsbt1 = sigmabt1 / Eb
        If EPS <= epsbt1 Then
           SigmaEpsConcrete = EPS * Eb
        Else
           If EPS > epsbt1 And EPS < Etens(3) Then
              SigmaEpsConcrete = ((1 - sigmabt1 / Rbt) * (EPS - epsbt1) / (Etens(3) - epsbt1) + sigmabt1 / Rbt) * Rbt
           Else
              If EPS >= Etens(3) And EPS <= Etens(4) Then
                 SigmaEpsConcrete = Rbt
              Else
                 SigmaEpsConcrete = 0#
              End If
           End If
        End If
     End If
  End If
'***************************************
End Select
End Function
Function SigmaEpsReinf(epsA#, Eult As Range, Es!, Rs!, Rsc!, reinfDiag As Variant)
Dim sigmaS1!, sigmaS1c!, epsS1#, epsS1c#
Select Case reinfDiag
'***************************************
  Case "двухлинейная"
  If epsA <= 0 Then 'при сжатии
     If epsA > -Eult(1) Then
        SigmaEpsReinf = Es * epsA
     Else
        SigmaEpsReinf = -Rsc
     End If
  Else 'при растяжении
     If epsA < Eult(1) Then
        SigmaEpsReinf = Es * epsA
     Else
        SigmaEpsReinf = Rs
     End If
  End If
'***************************************
  Case "трёхлинейная"
  sigmaS1 = 0.9 * Rs
  sigmaS1c = -0.9 * Rsc
  epsS1 = sigmaS1 / Es
  epsS1c = sigmaS1c / Es
  If epsA <= 0 Then 'при сжатии
     If epsA > epsS1c Then
        SigmaEpsReinf = epsA * Es
     Else
        If epsA <= epsS1c And epsA >= -(2 * Eult(1) - Abs(epsS1)) Then
           SigmaEpsReinf = -((1 - Abs(sigmaS1c) / Rsc) * (Abs(epsA) - Abs(epsS1c)) / (Eult(1) - Abs(epsS1c)) + Abs(sigmaS1c) / Rsc) * Rsc
        Else
           SigmaEpsReinf = -1.1 * Rsc
        End If
     End If
  Else 'при растяжении
     If epsA < epsS1 Then
        SigmaEpsReinf = epsA * Es
     Else
        If epsA >= epsS1 And epsA <= 2 * Eult(1) - epsS1 Then
           SigmaEpsReinf = ((1 - sigmaS1 / Rs) * (epsA - epsS1) / (Eult(1) - epsS1) + sigmaS1 / Rs) * Rs
        Else
           SigmaEpsReinf = 1.1 * Rs
        End If
     End If
  End If
'***************************************
End Select
End Function
Function MainIteration(ByVal MaxIteration As Integer, _
            ByVal ShowDialog As Boolean, ByVal StatusBar As Boolean) As Integer
Dim TOL_eps!, TOL_curv!, i%
Dim CRE, Bs, eps0, LD, EP As Variant
Dim sFini As String
Dim Deps, DkapX, DkapY As Double

TOL_eps = ActiveSheet.Cells(46, 44).Value2
TOL_curv = ActiveSheet.Cells(47, 44).Value2

i = 0 'счётчик итераций
eps0 = Range(Cells(3, 16), Cells(5, 16)) 'начальное значение деформаций
LD = Range(Cells(3, 13), Cells(5, 13)) 'вектор нагрузок
EP = eps0
MainIteration = 0

Do
CRE = EP
Cells(3, 25).Resize(UBound(EP), UBound(EP, 2)) = CRE
Calculate

Bs = Range(Cells(3, 21), Cells(5, 23))
i = i + 1
  If StatusBar = True Then
     Application.StatusBar = "Выполняется итерация " & i
  End If
  If i = MaxIteration Then
     If ShowDialog = True Then
        Dim answer%
        Dim sQuestion As String
        sQuestion = "Было выполнено " & i & " итераций, а решение не было найдено. Это может означать, что при введённых усилиях несущая способность сечения не обеспечена и решение уравнений не может быть найдено либо введены очень жёсткие требования к точности вычислений. Остановить расчёт?"
        answer = MsgBox(sQuestion, vbYesNo + vbQuestion, "Вопрос")
        If answer = vbYes Then
           If StatusBar = True Then
              Application.StatusBar = ""
           End If
           Exit Function
        End If
     Else
        MainIteration = 1
        Exit Function
     End If
  End If
EP = WorksheetFunction.MMult(WorksheetFunction.MInverse(Bs), LD) 'решение системы уравнений
  If Abs(EP(1, 1)) > 1# Or Abs(EP(2, 1)) > 1# Or Abs(EP(3, 1)) > 1# Then
     If ShowDialog = True Then
        MsgBox "Задача нелинейных уравнений расходится. Как правило, это означает, что при введённых усилиях несущая способность сечения не обеспечена и решение уравнений не может быть найдено. Необходимо уменьшить величины усилий или увеличить допуски на сходимость расчёта"
        If StatusBar = True Then
           Application.StatusBar = ""
        End If
        Exit Function
     Else
        MainIteration = 2
        Exit Function
     End If
  End If
  Deps = Abs((EP(1, 1) - CRE(1, 1)) / EP(1, 1))
  If LD(2, 1) = 0# Then ' если момент равен 0.0 кривизна тоже
     DkapX = 0#
  Else
     DkapX = Abs((EP(2, 1) - CRE(2, 1)) / EP(2, 1))
  End If
  If LD(1, 1) = 0# Then ' если момент равен 0.0 кривизна тоже
     DkapY = 0#
  Else
     DkapY = Abs((EP(3, 1) - CRE(3, 1)) / EP(3, 1))
  End If
Loop While Deps > TOL_eps Or _
           DkapX > TOL_curv Or _
           DkapY > TOL_curv

'вывод на лист достигнутой точности
ActiveSheet.Cells(3, 26).Value2 = Deps
ActiveSheet.Cells(4, 26).Value2 = DkapX
ActiveSheet.Cells(5, 26).Value2 = DkapY
ActiveSheet.Cells(6, 26).Value2 = i

If ShowDialog = True Then
    sFini = "Расчет окончен, выполнено " & i & " итераций. Требуемые параметры точности достигнуты."
    Call MsgBox(sFini, vbInformation, "Завершено")
End If
If StatusBar = True Then
    Application.StatusBar = ""
End If
End Function

