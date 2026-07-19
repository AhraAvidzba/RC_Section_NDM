Attribute VB_Name = "II_gruppa"
Option Base 1
Const pi As Double = 3.14159265358979

'Ab - площадь сжатой бетонной части, м2
'Sb - статический момент инерции сжатой бетонной части, м3
'yc - расстояние от нижней грани сечения до ц.т. сжатого бетонного сечения, м
'alfa- угол в зависимости от высоты сжатой зоны сечения, рад
'tsi - расстояние от i-ой арматуры до нейтральной оси (линия разделения сжатой и растянутой зоны сечения)
'y - расстояние от верхней грани сечения до ц.т. всего приведенноо сечения (без учета растянутого бетона).


Public Function Sb(ByVal x#, ByVal r#, ByVal var#) As Double 'статический момент инерции сжатой бетонной части
Dim Ab#, yc#, alfa
alfa = 2 * Acos((2 * r / 2 - x) / (2 * r / 2))
Ab = (2 * r) ^ 2 / 8 * (alfa - Sin(alfa))
yc = 2 * r / 2 + (2 * r * Sin(alfa / 2) ^ 3) / (3 * (alfa / 2 - Sin(alfa / 2) * Cos(alfa / 2)))
Sb = Ab * (yc - 2 * r + x)
If var = 1 Then Sb = Ab
End Function


Public Function Ssi_c(ByVal r#, ByVal M%, ByVal ls#, ByVal ds#, ByVal x#, ByVal var#) As Double 'статический момент инерции ОДНОГО СЖАТОГО ряда арматуры
Dim i%
For i = 1 To M

tsi = tsi_(i, r, M, ls) - (2 * r - x)

If tsi >= 0 Then

    If var = 0 Then
        Ssi_c = Ssi_c + asi_(ds) * Abs(tsi)
    Else
        Ssi_c = Ssi_c + asi_(ds)
    End If
End If

Next
End Function

Public Function Ssi_p(ByVal r#, ByVal M%, ByVal ls#, ByVal ds#, ByVal x#, ByVal var#) As Double 'статический момент инерции ОДНОГО РАССТЯНУТОГО ряда арматуры
Dim i%

For i = 1 To M
tsi = tsi_(i, r, M, ls) - (2 * r - x)

If tsi < 0 Then

    If var = 0 Then
        Ssi_p = Ssi_p + asi_(ds) * Abs(tsi)
    Else
        Ssi_p = Ssi_p + asi_(ds)
    End If
End If
  'Ssi_p = tsi_(5, r, M, ls) - (2 * r - x)
Next
End Function

Public Function C_gravity(ByVal r#, ByVal M1%, ByVal M2%, ByVal M3%, ByVal ls1#, ByVal ls2#, ByVal ls3#, ByVal ds1#, ByVal ds2#, ByVal ds3#, ByVal x#, ByVal N#) As Double 'расстояние от верхней границы сечения до центра тяжести приведенного сечения
Dim Ssi_c1#, Ssi_c2#, Ssi_c3#
Dim Asi_c1#, Asi_c2#, Asi_c3#
Dim Ssi_p1#, Ssi_p2#, Ssi_p3#
Dim Asi_p1#, Asi_p2#, Asi_p3#

If x <= 0 Then x = 0.000000000001
If x > 2 * r Then x = 2 * r

If M1 > 0 Then
Ssi_c1 = Ssi_c(r, M1, ls1, ds1, x, 0)
Ssi_p1 = Ssi_p(r, M1, ls1, ds1, x, 0)
Asi_c1 = Ssi_c(r, M1, ls1, ds1, x, 1)
Asi_p1 = Ssi_p(r, M1, ls1, ds1, x, 1)
End If

If M2 > 0 Then
Ssi_c2 = Ssi_c(r, M2, ls2, ds2, x, 0)
Ssi_p2 = Ssi_p(r, M2, ls2, ds2, x, 0)
Asi_c2 = Ssi_c(r, M2, ls2, ds2, x, 1)
Asi_p2 = Ssi_p(r, M2, ls2, ds2, x, 1)
End If

If M2 > 0 Then
Ssi_c3 = Ssi_c(r, M3, ls3, ds3, x, 0)
Ssi_p3 = Ssi_p(r, M3, ls3, ds3, x, 0)
Asi_c3 = Ssi_c(r, M3, ls3, ds3, x, 1)
Asi_p3 = Ssi_p(r, M3, ls3, ds3, x, 1)
End If

C_gravity = (Sb(x, r, 0) + (N - 1) * Ssi_c1 + (N - 1) * Ssi_c2 + (N - 1) * Ssi_c3 - N * Ssi_p1 - N * Ssi_p2 - N * Ssi_p3) / (Sb(x, r, 1) + (N - 1) * Asi_c1 + (N - 1) * Asi_c2 + (N - 1) * Asi_c3 + N * Asi_p1 + N * Asi_p2 + N * Asi_p3)
C_gravity = x - C_gravity
End Function


Public Function Ired(ByVal r#, ByVal M1%, ByVal M2%, ByVal M3%, ByVal ls1#, ByVal ls2#, ByVal ls3#, ByVal ds1#, ByVal ds2#, ByVal ds3#, ByVal x#, ByVal N#) As Double 'момент инерции сжатой бетонной части относительно ц.т. приведенного сечения
Dim Ib#, Ab#, yc#, alfa#, y#
Dim Isi_c1#, Isi_c2#, Isi_c3#
Dim Isi_p1#, Isi_p2#, Isi_p3#

y = C_gravity(r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, x, N)

If x <= 0 Then x = 0.000000000001
If x > 2 * r Then x = 2 * r

'момент инерции сжатой бетонной части относительно ц.т. приведенного сечения
alfa = 2 * Acos((2 * r / 2 - x) / (2 * r / 2))
Ab = (2 * r) ^ 2 / 8 * (alfa - Sin(alfa))
yc = 2 * r / 2 + (2 * r * Sin(alfa / 2) ^ 3) / (3 * (alfa / 2 - Sin(alfa / 2) * Cos(alfa / 2)))
Ib = (2 * r) ^ 4 / 128 * (alfa - Sin(alfa) * Cos(alfa)) - Ab * (yc - 2 * r / 2) ^ 2 + Ab * (yc - (2 * r - y)) ^ 2

'момент инерции сжатых рядов арматуры относительно ц.т. приведенного сечения

For i = 1 To M1
asi = asi_(ds1)
tsi = tsi_(i, r, M1, ls1) - (2 * r - x)
If tsi >= 0 Then Isi_c1 = Isi_c1 + asi * Abs(tsi - (x - y)) ^ 2
Next

If M2 > 0 Then
For i = 1 To M2
asi = asi_(ds2)
tsi = tsi_(i, r, M2, ls2) - (2 * r - x)
If tsi >= 0 Then Isi_c2 = Isi_c2 + asi * Abs(tsi - (x - y)) ^ 2
Next
End If

If M3 > 0 Then
For i = 1 To M3
asi = asi_(ds3)
tsi = tsi_(i, r, M3, ls3) - (2 * r - x)
If tsi >= 0 Then Isi_c3 = Isi_c3 + asi * Abs(tsi - (x - y)) ^ 2
Next
End If

'момент инерции растянутых рядов арматуры относительно ц.т. приведенного сечения

For i = 1 To M1
asi = asi_(ds1)
tsi = tsi_(i, r, M1, ls1) - (2 * r - x)
If tsi <= 0 Then Isi_p1 = Isi_p1 + asi * Abs(tsi - (x - y)) ^ 2
Next

If M2 > 0 Then
For i = 1 To M2
asi = asi_(ds2)
tsi = tsi_(i, r, M2, ls2) - (2 * r - x)
If tsi <= 0 Then Isi_p2 = Isi_p2 + asi * Abs(tsi - (x - y)) ^ 2
Next
End If

If M3 > 0 Then
For i = 1 To M3
asi = asi_(ds3)
tsi = tsi_(i, r, M3, ls3) - (2 * r - x)
If tsi <= 0 Then Isi_p3 = Isi_p3 + asi * Abs(tsi - (x - y)) ^ 2
Next
End If

Ired = Ib + (N - 1) * (Isi_c1 + Isi_c2 + Isi_c3) + N * (Isi_p1 + Isi_p2 + Isi_p3)
'Ired = 0.023808163963
End Function

Public Function Ab(ByVal x#, ByVal r#) As Double 'площадь сжатой бетонной части
Dim yc#, alfa#
alfa = 2 * Acos((2 * r / 2 - x) / (2 * r / 2))
Ab = (2 * r) ^ 2 / 8 * (alfa - Sin(alfa))
End Function


Public Function Asi_c(ByVal r#, ByVal M%, ByVal ls#, ByVal ds#, ByVal x#) As Double 'площадь ОДНОГО СЖАТОГО ряда арматуры
Dim i%
For i = 1 To M
asi = asi_(ds)
tsi = tsi_(i, r, M, ls) - (2 * r - x)
If tsi >= 0 Then Asi_c = Asi_c + asi
Next
End Function

Public Function Asi_p(ByVal r#, ByVal M%, ByVal ls#, ByVal ds#, ByVal x#) As Double 'площадь ОДНОГО РАССТЯНУТОГО ряда арматуры
Dim i%
For i = 1 To M
asi = asi_(ds)
tsi = tsi_(i, r, M, ls) - (2 * r - x)
If tsi < 0 Then Asi_p = Asi_p + asi
Next
End Function

Public Function Ared(r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, x, N) As Double 'приведенная площадь ВСЕГО сечения
If x <= 0 Then x = 0.000000000001
If x > 2 * r Then x = 2 * r
Ared = Ab(x, r) + (N - 1) * Asi_c(r, M1, ls1, ds1, x) + N * Asi_p(r, M1, ls1, ds1, x)
If M2 > 0 Then Ared = Ared + (N - 1) * Asi_c(r, M2, ls2, ds2, x) + N * Asi_p(r, M2, ls2, ds2, x)
If M3 > 0 Then Ared = Ared + (N - 1) * Asi_c(r, M3, ls3, ds3, x) + N * Asi_p(r, M3, ls3, ds3, x)
'Ared = 0.568628270299
End Function


Public Function sigma(ByVal r#, ByVal M1%, ByVal M2%, ByVal M3%, ByVal ls1#, ByVal ls2#, ByVal ls3#, ByVal ds1#, ByVal ds2#, ByVal ds3#, ByVal x#, ByVal N#, N_vn, M_vn) 'расчет напряжения в точке находящейся на расстоянии X от верхней границы сечения (ниже от верхней грани сечения (+), выше (-))
M_vn = Abs(M_vn)
Dim Ared_#, Ired_#
Ared_ = Ared(r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, x, N)
Ired_ = Ired(r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, x, N)
y = C_gravity(r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, x, N)

sigma = N_vn / Ared_ + M_vn * (y - x) / Ired_ - N_vn * (r - y) * (y - x) / Ired_
End Function



'Public Function zonaII(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, N, N_vn, M_vn, tochnostII)  'высота сжатой зоны при которой напряжения в нейтральной линии равны нулю
'M_vn = Abs(M_vn)
'Dim shodII, shII, xII#
'Dim a# 'шаг итераций
'Dim ii%
'
'Dim sigma_0#, sigma_d#
'
'sigma_0 = sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, 0, N, N_vn, M_vn) 'напряжения в верхней части сечения (x=0)
'sigma_d = sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, 2 * r, N, N_vn, M_vn) 'напряжения в нижней части сечения (x=d)
'
'If Sgn(sigma_0) = Sgn(sigma_d) Then
'
'    If N_vn > 0 Then xII = 2 * r Else xII = 0
'
'Else
'
'    shII = -1000000
'    a = 1
'
'    For ii = 1 To 10
'
'If sigma_0 <= 1 Then: xII = 0: Exit For
'If sigma_d >= -0.1 Then: xII = 2 * r: Exit For
'
'    If shII < 0 Then a = Abs(a) / 10 Else a = -Abs(a) / 10
'
'        For i = 1 To WorksheetFunction.RoundUp(2 * r / 0.1, 0) + 1
'
'            xII = xII + a
'            shodII = -sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, xII, N, N_vn, M_vn)
'            If Abs(shodII) > Abs(shII) Then Exit For
'            shII = shodII
'            'Debug.Print xII, shII, shodII, a, i, ii, tochnostII
'        Next
'
'            xII = xII - a
'            'Debug.Print xII, shII, shodII, a, i, ii, tochnostII
'            If Abs(shII) <= tochnostII Then Exit For
'
'    Next
'
'End If
'zonaII = xII
'End Function

Public Function zonaII(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, N, N_vn, M_vn, tochnostII)  'высота сжатой зоны при которой напряжения в нейтральной линии равны нулю
M_vn = Abs(M_vn)
Dim Q, xII#
Dim a# 'шаг итераций
Dim ii%

Dim sigma_0#, sigma_d#

sigma_0 = sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, 0, N, N_vn, M_vn) 'напряжения в верхней части сечения (x=0)
sigma_d = sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, 2 * r, N, N_vn, M_vn) 'напряжения в нижней части сечения (x=d)

If Sgn(sigma_0) = Sgn(sigma_d) Then

    If N_vn > 0 Then xII = 2 * r Else xII = 0

Else

    Q = sigma_0
    a = 2 * r

    For ii = 1 To 5

    If Q > 0 Then a = Abs(a) / 10 Else a = -Abs(a) / 10

        For i = 1 To 10
            xII = xII + a
            If Abs(Q) < Abs(sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, xII, N, N_vn, M_vn)) Then
            xII = xII - a
            Exit For
            End If
            Q = sigma(r, M1, M2, M3, ls1, ls2, ls3#, ds1, ds2, ds3, xII, N, N_vn, M_vn)
            If Abs(Q) <= tochnostII Then Exit For
            'Debug.Print xII, Q, a, i, ii, tochnostII
        Next
            'Debug.Print xII, Q, a, i, ii, tochnostII
            If Abs(Q) <= tochnostII Then Exit For

    Next

End If
zonaII = xII
End Function

