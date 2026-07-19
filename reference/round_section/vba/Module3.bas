Attribute VB_Name = "Module3"
Option Base 1


Public Function W(a As Range, x As Double, y As Double)
Dim a1, b1, c1, d1, c2, d2, g1, g2 As Double
Dim i, k, t As Integer
Dim M, M1, M2 As Double


    k = 1
    For i = 1 To a.Columns.Count
        k = k + 1
        If a(1, k) >= x Then Exit For
    Next

    t = 1
    For i = 1 To a.Rows.Count
        t = t + 1
        If a(t, 1) >= y Then Exit For
    Next

If x > a(1, a.Columns.Count) Or y > a(a.Rows.Count, 1) Or x < a(1, 2) Or y < a(2, 1) Then
    W = "error"
Else


    If k > 2 And t > 2 Then
             a1 = a(t, 1)
             b1 = a(t - 1, 1)
             c1 = a(t - 1, k - 1)
             d1 = a(t, k - 1)
             c2 = a(t - 1, k)
             d2 = a(t, k)
             g1 = a(1, k - 1)
             g2 = a(1, k)

             M1 = ((c1 - d1) * (y - a1) + d1 * (b1 - a1)) / (b1 - a1)
             M2 = ((c2 - d2) * (y - a1) + d2 * (b1 - a1)) / (b1 - a1)
             M = ((x - g1) * (M2 - M1) + M1 * (g2 - g1)) / (g2 - g1)
             W = M

    End If

    If k = 2 And t = 2 Then
    W = a(k, t)
    End If

    If k > 2 And t = 2 Then
    W = a(t, k) - ((a(t, k) - a(t, k - 1)) * (a(1, k) - x)) / (a(1, k) - a(1, k - 1))
    End If

    If k = 2 And t > 2 Then
    W = a(t - 1, k) - ((a(t - 1, k) - a(t, k)) * (a(t - 1, 1) - y)) / (a(t - 1, 1) - a(t, 1))
    End If

End If

End Function





'Public Function split(B As Range, L As Integer, l2 As Integer, h As Double)
'Dim nii(), x()  As Variant, i, j, k, t, ni, N As Integer
''b - диапазон €чеек с номерами грунтов (в первом столбце) и толщинами грунтовых слоев (во втором столбце)
''nii() - массив с количеством слоев (не более 2-х м)  в каждом грунтовом слое
''x - массив с толщинами разбитых грунтовых толщь
''n - общее количество разбитых слоев
''l = 20 ' количество неразбитых слоев (число строк в выбранном диапазоне)
''l2 = 100 ' количество разбитых слоев
''h - максимальна€ толщина разбитого сло€
'ReDim nii(L)
'ReDim x(l2, 3)
'
'' находим количество слоев в каждом массиве грунта (ni) и общее количество слоев (n)
'For i = 1 To L
'    k = k + 1
'        If Int(B(k, 3) / h) * h = B(k, 3) Then
'            ni = B(k, 3) / h
'        Else
'            ni = Int(B(k, 3) / h) + 1
'        End If
'    N = N + ni
'    nii(k) = ni
'Next
''провер€ем вместимость слоев в наш массив и разбиваем грунтовые слои на слои <=h метров
'If N > l2 Then
'    split = "ошибка 1"
'Else
'k = 0
'    For i = 1 To L
'        k = k + 1
'        If B(k, 3) - h * nii(k) = 0 Then
'            For j = 1 To nii(k)
'                t = t + 1
'                x(t, 3) = h
'                x(t, 2) = B(k, 2)
'                x(t, 1) = B(k, 1)
'            Next
'        End If
'        If B(k, 3) - h * nii(k) < 0 Then
'            For j = 1 To nii(k) - 1
'                t = t + 1
'                x(t, 3) = h
'                x(t, 2) = B(k, 2)
'                x(t, 1) = B(k, 1)
'            Next
'            t = t + 1
'            x(t, 3) = B(k, 3) - h * (nii(k) - 1)
'            x(t, 2) = B(k, 2)
'            x(t, 1) = B(k, 1)
'        End If
'    Next
'    split = x
'End If
'
'End Function
'
'
'Public Function point(a As Range, B As Range, g As Double)
'Dim xp1, yp1, xp2, yp2 As Double
'Dim xv1, yv1, xv2, yv2 As Double
'Dim y, x As Variant
'Dim k, t As Integer ' u
'Dim xpb, xpm, ypb, ypm, xvb, xvm, yvb, yvm  As Double 'pp
''Dim m(50, 10) As Double
'k = 0
'For i = 1 To a.Rows.Count - 1
'k = k + 1
'xp1 = a(k, 2)
'xp2 = a(k + 1, 2)
'yp1 = a(k, 1)
'yp2 = a(k + 1, 1)
'x = "!XX!"
'y = "!XX!"
'If xp1 = xp2 And yp1 = yp2 Then Exit For
'
'    For j = 1 To B.Rows.Count - 1
'    u = u + 1
'        t = t + 1
'        xv1 = B(t, 2)
'        xv2 = B(t + 1, 2)
'        yv1 = B(t, 1)
'        yv2 = B(t + 1, 1)
'        If xv1 = xv2 And yv1 = yv2 Then Exit For
'
'   If xp1 = xp2 Then xp2 = xp1 + 0.0001
'   If xv1 = xv2 Then xv2 = xv1 + 0.0001
'   If yp1 = yp2 Then yp2 = yp1 + 0.0001
'   If yv1 = yv2 Then yv2 = yv1 + 0.0001
'
'    If xp1 > xp2 Then
'    xpb = xp1
'    xpm = xp2
'    Else
'    xpb = xp2
'    xpm = xp1
'    End If
'
'    If xv1 > xv2 Then
'    xvb = xv1
'    xvm = xv2
'    Else
'    xvb = xv2
'    xvm = xv1
'    End If
'
'    If yp1 > yp2 Then
'    ypb = yp1
'    ypm = yp2
'    Else
'    ypb = yp2
'    ypm = yp1
'    End If
'
'    If yv1 > yv2 Then
'    yvb = yv1
'    yvm = yv2
'    Else
'    yvb = yv2
'    yvm = yv1
'    End If
'
'        y = (yp1 * (xp2 - xp1) * (yv2 - yv1) - yv1 * (yp2 - yp1) * (xv2 - xv1) + xv1 * (yp2 - yp1) * (yv2 - yv1) - xp1 * (yp2 - yp1) * (yv2 - yv1)) / ((xp2 - xp1) * (yv2 - yv1) - (yp2 - yp1) * (xv2 - xv1))
'        x = (y - yv1) * (xv2 - xv1) / (yv2 - yv1) + xv1
'
''pp = 0
'
'If x > xpm And x < xpb Then
''pp = 11
'If y > ypm And y < ypb Then
''pp = 22
'If x > xvm And x < xvb Then
''pp = 33
'If y > yvm And y < yvb Then
'Exit For
''pp = 44
'End If
'End If
'End If
'End If
'
'
''m(u, 1) = pp
''m(u, 2) = x
''m(u, 3) = xpb
''m(u, 4) = xpm
''m(u, 5) = ypb
''m(u, 6) = ypm
''m(u, 7) = xvb
''m(u, 8) = xvm
''m(u, 9) = yvb
''m(u, 10) = yvm
'
'
'Next
'
'If x > xpm And x < xpb Then
''pp = 11
'If y > ypm And y < ypb Then
''pp = 22
'If x > xvm And x < xvb Then
''pp = 33
'If y > yvm And y < yvb Then
'Exit For
''pp = 44
'End If
'End If
'End If
'End If
'
'
't = 0
'Next
'        If g = 1 Then point = x
'        If g = 0 Then point = y
''point = m
'
'End Function
'
'Public Function e_cir(N As Double, prop As Range, step As Double, fault As Double)
'Dim Rs, As_tot, Rb, Ab  As Double
'Dim x, pi  As Double
'pi = 3.14159265358979
'
''prop_2 = Rs
''prop_3 = As,tot
''prop_4 = Rb
''prop_5 = Ab
'Rs = prop(1)
'As_tot = prop(2)
'Rb = prop(3)
'Ab = prop(4)
'
'If N <= 0.77 * Rb * Ab + 0.645 * Rs * As_tot Then
'    Do
'    x = x + step
'
'    e_cir = (N + Rs * As_tot + Rb * Ab * Sin(2 * pi * x) / (2 * pi)) / (Rb * Ab + 2.55 * Rs * As_tot)
'
'    Loop While Abs(e_cir - x) >= fault
'Else
' Do
'    x = x + step
'
'    e_cir = (N + Rb * Ab * Sin(2 * pi * x) / (2 * pi)) / (Rb * Ab + Rs * As_tot)
'
'    Loop While Abs(e_cir - x) >= fault
'End If
'
'End Function
'
'
Sub MAX()
Dim N, t, n_max, t_max As Integer
Dim i, j, k, k1, p, p1 As Integer
Dim zapas, zapas2 As Double
zapas = 1000000000
'n - номер сваи
't - номер загружени€
'n_max - всего свай
't_max - всего загружений

Dim r As Range
Set r = Worksheets("расчет").Range("A1")

n_max = 14 'r(63, 65)
t_max = 7 'r(64, 65)

For i = 1 To n_max
j = 0
p = 0
k = k + 1
r(63, 5) = k
    For j = 1 To t_max
    p = p + 1
    r(64, 5) = p
    zapas2 = r(62, 6)

    If zapas2 < zapas Then
    zapas = zapas2
    k1 = k
    p1 = p
    End If
    Next

Next


'r(5, 7) = i2

r(63, 6) = zapas
r(63, 5) = k1
r(64, 5) = p1

End Sub
'
'Public Function alfaAAA(r As Double, Rs As Double, Astot As Double, Rb As Double, Ab As Double, N As Double, shag As Double, tochnost As Double)
'Dim A_s, A_s_сж, Abc  As Double
'Dim Sa1, Sa2, Sb, h, L  As Double
'Dim a, a_rad As Double
'Dim i  As Integer
'Dim pi  As Double
'Dim left, right  As Double
'
'pi = 3.14159265358979
'
''For i = 1 To 2 * pi
'Do
'
'a_rad = a_rad + shag
'h = r * (1 - Cos(a_rad / 2))
'L = r * a_rad
'Sb = 1 / 2 * (r * L - 2 * ((h * (2 * r - h)) ^ (1 / 2)) * (r - h))
'Sa1 = a_rad / 2 * Astot / pi
'Sa2 = Astot * (1 - a_rad / 2 / pi)
'left = N + Rs * Sa2
'right = Rs * Sa1 + Rb * Sb
'Loop Until left - right <= tochnost
''Next
'alfa = a_rad
'End Function
'
'
'Public Function alfa(r As Double, Rs As Double, Astot As Double, Rb As Double, Ab As Double, N As Double, shag As Double, tochnost As Double)
'Dim A_s, A_s_сж, Abc  As Double
'Dim Sa1, Sa2, Sb, h, L  As Double
'Dim e_cir As Double
'Dim i  As Integer
'Dim pi  As Double
'Dim a  As Double 'угол альфа
'Dim left, right  As Double
'
'pi = 3.14159265358979
'
'Do
'
'e_cir = e_cir + shag
'
'left = Rb * Ab * (e_cir - Sin(2 * pi * e_cir) / 2 / pi) + Rs * Astot * e_cir - Rs * Astot * (1 - 1.55 * e_cir)
'
'If e_cir > 0.645 Then left = Rb * Ab * (e_cir - Sin(2 * pi * e_cir) / 2 / pi) + Rs * Astot * e_cir
'
'right = N
'
'Loop Until left - right >= tochnost
'
''Do
''a = a + shag
''Loop Until e_cir - r ^ 2 * (a / 2 - Sin(a / 2) * Cos(a / 2)) / Ab <= tochnost
''
''
''If result = 0 Then alfa = e_cir
''If result = 1 Then alfa = a
'alfa = e_cir
'
'
'End Function
'
'
'
'
''Public Function NDM_beton(har As Range, D As Double, asmin As Double, es1 As Double, es2 As Double, eb1 As Double, eb2 As Double, Xr As dauble)
''Dim X  As Double
''
''Dim Sa1, Sa2, Sb, h, L  As Double
''Dim e_cir As Double
''Dim i  As Integer
''Dim pi  As Double
''Dim a  As Double 'угол альфа
''Dim left, right  As Double
''
''pi = 3.14159265358979
''
''Do
''
''Loop Until left - right >= tochnost
''
''
''End Function
'
'
'
'Sub PROCH()
''Mx#, Mx_pr#, N#, N_pr#
'Dim i, j As Integer
'Dim e0, r  As Double
'Dim shag, tochnost  As Double
'Dim Mx#, Mx_pr#, N#, N_pr#
'
'Set L = Worksheets("расчет").Range("A1")
'
'
'
'shag = 0.000001
'tochnost = 50
'
'For i = 1 To 20
'e0 = e0 + shag
'    For j = 1 To 20
'        r = r + shag / 10
'
'        L(75, 19) = e0
'        L(76, 19) = r
'
'        Mx = L(83, 18)
'        N = L(84, 18)
'        Mx_pr = L(83, 19)
'        N_pr = L(84, 19)
'
'        If Abs(Mx - Mx_pr) <= tochnost And Abs(N - N_pr) <= tochnost Then Exit For
'    Next
'    r = 0
'Next
'
'L(80, 15) = r
'L(81, 15) = e0
'End Sub

