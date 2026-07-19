VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public r As Double 'радиус окружности (сваи)
Public N As Integer ' количество элементов на которые бьется окружность по высоте


Function Acos(ByVal a As Double) As Double 'функция арккосинуса (напрямую к встроенной функции эксель вба обращается долго)
If a = 1 Then
Acos = 0
ElseIf a = -1 Then
Acos = 3.14159265358979
Else
Acos = Atn(-a / Sqr(-a * a + 1)) + 2 * Atn(1)
End If
End Function


Public Sub Abk_Hk()
Dim ak, Sk, v As Double
Dim k As Integer
r = 0.75
k = 1
v = 7.5 / 1000
ak = 2 * Acos((r - k * v) / r)
End Sub

