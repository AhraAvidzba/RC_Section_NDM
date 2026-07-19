Attribute VB_Name = "NDM"
Option Base 1
Const pi As Double = 3.14159265358979

Public x_end As Double 'принятая к расчету (окончательная) высота сжатой зоны
'глобальные переменные для бетона
Public bet() As Double '(Abk(), hk(), ebk(), Qbk())
'глобальные переменные для арматуры
Public arm1() As Double '(tsi(), asi(), esi(), Qsi())
Public arm2() As Double
Public arm3() As Double




Public Function Acos(ByVal a As Double) As Double 'функция арккосинуса (напрямую к встроенной функции эксель вба обращается долго)
If a = 1 Then
Acos = 0
ElseIf a = -1 Then
Acos = pi
Else
Acos = Atn(-a / Sqr(-a * a + 1)) + 2 * Atn(1)
End If
End Function

'# - double
'% - integer

'k - номер сегмента бетонного круглого сечения
'r - радиус круглого сечения, м
'v - высота одного сегмента круглого бетонного сечения (сечение бьется на несколько равных по высоте сегментов), м
'hk - расстояние от крайней нижней точки сечения (окружности) до k-го сегмента бетонного круглого сечения, м
'es2 - предельные относительные деформации арматуры
'eb2 - предельные относительные деформации бетона
'asmin - расстояние от торца сечения до ближайшей к торцу оси ряда арматуры, м
'Xr - граничная высота сжатой зоны бетона (при которой разрушение происходит одновременно по бетону и арматуре), м
'X - высота сжатой зоны, м
'Qbk - напряжения в k-ом сегменте бетонного сечения, т/м2
'eb1 - относительные деформации бетона при кот. проявляется пластика (при увеличении относительных деформаций больше eb1 напряжения не возрастают)
'Rb - расчетная прочность бетона, т/м2
'Ebred - приведенный модуль деформации бетона, т/м2
'N - количество сегментов по высоте, на которые разбивается бетонное круглое сечение
'N_vneshnee - внешняя продольная сила, т



Function Sk_(ByVal k%, ByVal r#, ByVal v#) 'площадь k-го сегмента бетонного круглого сечения
Dim ak As Double
ak = 2 * Acos((r - k * v) / r)
Sk_ = (2 * r) ^ 2 / 8 * (ak - Sin(ak))
End Function

Function hk_(ByVal k%, ByVal r#, ByVal v#) 'расстояние от крайней нижней точки сечения (окружности) до k-го сегмента бетонного круглого сечения
hk_ = 2 * r - (k - 1) * v - v / 2
End Function

Function ebk_(ByVal hk#, ByVal r#, ByVal es2#, ByVal eb2#, ByVal asmin#, ByVal xr#, ByVal x#) 'относительные деформации в k-ом сегменте бетонного сечения
If x <= xr And x > 0 Then ebk_ = es2 * (hk - (2 * r - x)) / (2 * r - x - asmin)
If x > xr And x < 2 * r Then ebk_ = eb2 * (hk - (2 * r - x)) / x

If x <= 0 Then ebk_ = 0 'es2 * (1 - hk / (2 * r + Abs(x)))
If x >= 2 * r Then ebk_ = eb2 * (x - 2 * r + hk) / x

End Function

Function Qbk_(ByVal ebk#, ByVal eb1#, ByVal Ebred#, ByVal Rb#) 'напряжения в k-ом сегменте бетонного сечения
If ebk <= 0 Then Qbk_ = 0
If ebk > 0 And ebk <= eb1 Then Qbk_ = Ebred * ebk
If ebk > eb1 Then Qbk_ = Rb
End Function

Public Function bet_(r, N, es2, eb2, asmin, xr, x, eb1, Ebred, Rb) 'создание массива из данных по бетону
'Dim Abk(), hk(), ebk(), Qbk() As Double
Dim v# 'высота одного сегмента
Dim k%
Dim bet__() As Double
v = 2 * r / N
ReDim Abk(N)
ReDim hk(N)
ReDim ebk(N)
ReDim Qbk(N)
ReDim bet__(N, 4)

For k = 1 To N
    Abk(k) = Sk_(k, r, v) - Sk_(k - 1, r, v)
    hk(k) = hk_(k, r, v)
    ebk(k) = ebk_(hk(k), r, es2, eb2, asmin, xr, x)
    Qbk(k) = Qbk_(ebk(k), eb1, Ebred, Rb)
    'Debug.Print Abk(k)
    
    bet__(k, 1) = Abk(k)
    bet__(k, 2) = hk(k)
    bet__(k, 3) = ebk(k)
    bet__(k, 4) = Qbk(k)
Next

bet_ = bet__
End Function



'i - номер арматуры в ряду
'asi - площадь отдельной i-ой арматуры в одном ряду, м2
'tsi - расстояние от крайней нижней точки бетонного сечения (окружности) до оси i-го стержня в ряду, м
'ds - диаметр арматуры, м
'm - количество стержней арматуры в одном ряду
'ls - радиус окружности по оси ряда арматуры, м
'esi - относительные деформации в i-ом арматурном стержне
'es2 - предельные относительные деформации арматуры
'es0 - относительные деформации арматуры при кот. проявляется пластика (при увеличении относительных деформаций больше eb1 напряжения не возрастают)
'Es - модуль деформации арматуры, т/м2
'Rs - расчетная прочность арматуры, т/м2


Public Function asi_(ByVal ds#) 'площадь отдельной i-ой арматуры в одном ряду, м2
If ds = 6 / 1000 Then asi_ = 0.283 / 10000
If ds = 8 / 1000 Then asi_ = 0.503 / 10000
If ds = 10 / 1000 Then asi_ = 0.785 / 10000
If ds = 12 / 1000 Then asi_ = 1.131 / 10000
If ds = 14 / 1000 Then asi_ = 1.539 / 10000
If ds = 16 / 1000 Then asi_ = 2.011 / 10000
If ds = 18 / 1000 Then asi_ = 2.545 / 10000
If ds = 20 / 1000 Then asi_ = 3.142 / 10000
If ds = 22 / 1000 Then asi_ = 3.801 / 10000
If ds = 25 / 1000 Then asi_ = 4.909 / 10000
If ds = 28 / 1000 Then asi_ = 6.158 / 10000
If ds = 32 / 1000 Then asi_ = 8.042477 / 10000
If ds = 36 / 1000 Then asi_ = 10.18 / 10000
If ds = 40 / 1000 Then asi_ = 12.56 / 10000
If ds = 45 / 1000 Then asi_ = 15.904 / 10000
If ds = 50 / 1000 Then asi_ = 19.635 / 10000
If ds = 55 / 1000 Then asi_ = 23.76 / 10000
If ds = 60 / 1000 Then asi_ = 28.27 / 10000
End Function

Function tsi_(ByVal i%, ByVal r#, ByVal M%, ByVal ls#) 'расстояние от крайней нижней точки сечения (окружности) до оси i-го стержня в ряду, м
If i <= M Then
tsi_ = r - ls * Sin((360 * (i - 1) / M - 90) * pi / 180)
Else
tsi_ = 0
End If
End Function

Function esi_(ByVal tsi#, ByVal r#, ByVal es2#, ByVal eb2#, ByVal asmin#, ByVal xr#, ByVal x#) 'относительные деформации в i-ом арматурном стержне
If x <= xr And x > 0 Then esi_ = es2 * (tsi - (2 * r - x)) / (2 * r - x - asmin)
If x > xr And x < 2 * r Then esi_ = eb2 * (tsi - (2 * r - x)) / x

If x <= 0 Then esi_ = -es2 * (1 - tsi / (2 * r + Abs(x)))
If x >= 2 * r Then esi_ = eb2 * (x - 2 * r + tsi) / x
End Function

Function Qsi_(ByVal esi#, ByVal es0#, ByVal es2#, ByVal Es#, ByVal Rs#) 'напряжения в i-ом арматурном стержне
If esi <= -es0 Then Qsi_ = -Rs
If esi > -es0 And esi <= es0 Then Qsi_ = Es * esi
If esi > es0 Then Qsi_ = Rs
End Function

Public Function arm_(r, M, ls, ds, es2, eb2, asmin, xr, x, es0, Es, Rs) 'создание массива из данных по арматуре
Dim i%
Dim arm__() As Double
ReDim tsi(M)
ReDim asi(M)
ReDim esi(M)
ReDim Qsi(M)
ReDim arm__(M, 4)

For i = 1 To M
tsi(i) = tsi_(i, r, M, ls)
asi(i) = asi_(ds)
esi(i) = esi_(tsi(i), r, es2, eb2, asmin, xr, x)
Qsi(i) = Qsi_(esi(i), es0, es2, Es, Rs)

arm__(i, 1) = tsi(i)
arm__(i, 2) = asi(i)
arm__(i, 3) = esi(i)
arm__(i, 4) = Qsi(i)
Next

arm_ = arm__
End Function

Public Function shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x)  'расчет сходимости
Dim s%
Dim N_bet, N_s1, N_s2, N_s3#

ReDim arm1(M1, 4)
If M2 > 0 Then ReDim arm2(M2, 4)
If M3 > 0 Then ReDim arm3(M3, 4)

bet = bet_(r, N, es2, eb2, asmin, xr, x, eb1, Ebred, Rb)
arm1 = arm_(r, M1, ls1, ds1, es2, eb2, asmin, xr, x, es0, Es, Rs)
If M2 > 0 Then arm2 = arm_(r, M2, ls2, ds2, es2, eb2, asmin, xr, x, es0, Es, Rs)
If M3 > 0 Then arm3 = arm_(r, M3, ls3, ds3, es2, eb2, asmin, xr, x, es0, Es, Rs)

For s = 1 To N
N_bet = N_bet + bet(s, 4) * bet(s, 1)
Next

For s = 1 To M1
N_s1 = N_s1 + arm1(s, 4) * arm1(s, 2)
Next

If M2 > 0 Then
For s = 1 To M2
N_s2 = N_s2 + arm2(s, 4) * arm2(s, 2)
Next
End If

If M3 > 0 Then
For s = 1 To M3
N_s3 = N_s3 + arm3(s, 4) * arm3(s, 2)
Next
End If

shod_ = (N_bet + N_s1 + N_s2 + N_s3) - N_vneshnee
End Function



'Public Function zona(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, tochnost)  'расчет высоты сжатой зоны
'Dim shod, sh, x#
'Dim a# 'шаг итераций
'Dim ii%
'
'Dim x_min#, x_max#
'Dim sh_x_min#, sh_x_0#, sh_x_d#, sh_x_max#
'
'x_min = -(es0 * (2 * r - asmin) - es2 * asmin) / (es2 - es0)
'If x_min > 0 Then x_min = 0
'x_max = eb2 * 2 * r / (eb2 - eb1)
'
'sh_x_min = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x_min)
''sh_x_0 = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, 0)
''sh_x_d = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, 2 * r)
'sh_x_max = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x_max)
'
'x = x_min
'If N_vneshnee >= 0 And sh_x_max > 0 Or N_vneshnee <= 0 And sh_x_min < 0 Then
'
'    sh = sh_x_min
'    a = 1
'
'    For ii = 1 To 10
'
'    'Debug.Print sh, x, a, ii
'    If sh < 0 Then a = Abs(a) / 10 Else a = -Abs(a) / 10
'
'        For i = 1 To 50
'      'Debug.Print a
'            x = x + a
'            shod = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x)
'            If Abs(shod) >= Abs(sh) Then Exit For
'            sh = shod
'            'Debug.Print sh, x, a, ii, i
'        Next
'
'            x = x - a
''Debug.Print sh, tochnost, x
'            If Abs(sh) <= tochnost Then Exit For
'
'    Next
'
'Else
'x = -10000
'End If
'
''Debug.Print x
'zona = x
'
'End Function

Public Function zona(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, tochnost)  'расчет высоты сжатой зоны
Dim sh#, x# 'shod#,
Dim a# 'шаг итераций
Dim ii%

Dim x_min#, x_max#
Dim sh_x_min#, sh_x_max#

x_min = -(es0 * (2 * r - asmin) - es2 * asmin) / (es2 - es0)
If x_min > 0 Then x_min = 0
x_max = eb2 * 2 * r / (eb2 - eb1)

sh_x_min = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x_min)
sh_x_max = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x_max)

If N_vneshnee >= 0 And sh_x_max >= 0 Or N_vneshnee <= 0 And sh_x_min <= 0 Then
x = x_min
sh = sh_x_min
    a = Abs(x_min) + Abs(x_max)

    For ii = 1 To 5
    'Debug.Print sh, x, a, ii
    If sh < 0 Then a = Abs(a) / 10 Else a = -Abs(a) / 10
    
        For i = 1 To 9
            x = x + a
            If x > x_max Then Exit For
            If Abs(sh) < Abs(shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x)) Then
            x = x - a
            Exit For
            End If
            sh = shod_(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, x)
            'Debug.Print sh, x, a, ii, i
            If Abs(sh) <= tochnost Then Exit For
        Next
    If Abs(sh) <= tochnost Then Exit For
    Next
Else
x = -10000
End If

'Debug.Print x
zona = x

End Function


Public Function result(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, tochnost, rez) 'расчет удерживающего момента

Erase bet
Erase arm1
Erase arm2
Erase arm3
M_bet = 0
M_s1 = 0
M_s2 = 0
M_s3 = 0

x_end = zona(N_vneshnee, r, M1, M2, M3, ls1, ls2, ls3, ds1, ds2, ds3, es2, eb2, asmin, xr, es0, Es, Rs, N, eb1, Ebred, Rb, tochnost)

bet = bet_(r, N, es2, eb2, asmin, xr, x_end, eb1, Ebred, Rb)
arm1 = arm_(r, M1, ls1, ds1, es2, eb2, asmin, xr, x_end, es0, Es, Rs)
If M2 > 0 Then arm2 = arm_(r, M2, ls2, ds2, es2, eb2, asmin, xr, x_end, es0, Es, Rs)
If M3 > 0 Then arm3 = arm_(r, M3, ls3, ds3, es2, eb2, asmin, xr, x_end, es0, Es, Rs)

For s = 1 To N
M_bet = M_bet + bet(s, 4) * bet(s, 1) * (bet(s, 2) - r)
Next

For s = 1 To M1
M_s1 = M_s1 + arm1(s, 4) * arm1(s, 2) * (arm1(s, 1) - r)
Next

If M2 > 0 Then
For s = 1 To M2
M_s2 = M_s2 + arm2(s, 4) * arm2(s, 2) * (arm2(s, 1) - r)
Next
End If

If M3 > 0 Then
For s = 1 To M3
M_s3 = M_s3 + arm3(s, 4) * arm3(s, 2) * (arm3(s, 1) - r)
Next
End If

'If x_end < 2 * r Then
If rez = 0 Then result = M_bet + M_s1 + M_s2 + M_s3
'End If

'If x_end >= 2 * r Then
'If rez = 0 Then result = 0
'End If

If rez = 1 Then result = x_end


End Function

