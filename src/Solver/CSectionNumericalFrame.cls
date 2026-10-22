VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionNumericalFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Начало координат и масштабы численных уравнений
' ==========================================================================
' Переносит поле деформаций и усилия к подготовленному центру сетки
' и задаёт масштаб длины для устойчивого решения уравнений.
' Исходная геометрия и координаты, доступные пользователю, сохраняются.

Private mSection As CSectionModel ' Модель, для которой подготовлены численные координаты.
Private mRevision As Long ' Версия геометрии при подготовке координат.
Private mOriginX As Double, mOriginY As Double, mLength As Double ' Координаты численного начала и характерная длина для нормирования, мм.
Private mSupportChecked As Boolean, mFullSupport As Boolean, mConcreteArea As Double ' Проверена ли геометрия, независимы ли epsilon0/kappaX/kappaY, и площадь бетона в мм2.

' Задаёт единичный численный масштаб и нулевое начало до подготовки геометрии.
Private Sub Class_Initialize()
    mLength = 1#
End Sub

' Центр берется из общей подготовки бетонной сетки модели; здесь определяются
' только численный масштаб и правила переноса. Геометрический радиус также
' кэшируется моделью, а не пересчитывается для каждой пробной плоскости.
Friend Sub Initialize(ByVal section As CSectionModel)
    If section Is mSection Then
        If section.Revision = mRevision Then Exit Sub
    End If
    Dim sumArea As Double
    section.ConcreteMeshCentroid mOriginX, mOriginY, sumArea
    mLength = section.MaximumElementRadiusFromConcreteCenter
    If mLength <= 0# Then mLength = 1# ' Вырожденная геометрия остается вырожденной; масштаб не добавляет жесткости.
    mConcreteArea = sumArea: mSupportChecked = False
    Set mSection = section: mRevision = section.Revision
End Sub

' Регуляризация материала допустима только при независимой геометрии трех
' параметров плоскости. Она не исправляет сечение из одной точки или линии.
Friend Function HasFullPlaneSupport() As Boolean
    If mSection Is Nothing Then Exit Function
    If Not mSupportChecked Then
        Dim g(1 To 3, 1 To 3) As Double, i As Long, x As Double, y As Double, area As Double ' g — матрица Грама полей 1, y/L и x/L; её ранг проверяет независимость параметров плоскости.
        For i = 1 To mSection.ConcreteCount + mSection.RebarCount
            If i <= mSection.ConcreteCount Then
                mSection.ConcreteElementForSolver i, x, y, area
            Else
                mSection.RebarElementForSolver i - mSection.ConcreteCount, x, y, area
            End If
            x = (x - mOriginX) / mLength: y = (y - mOriginY) / mLength
            area = area / mConcreteArea
            g(1, 1) = g(1, 1) + area
            g(1, 2) = g(1, 2) + area * y: g(1, 3) = g(1, 3) + area * x
            g(2, 2) = g(2, 2) + area * y * y: g(2, 3) = g(2, 3) + area * x * y
            g(3, 3) = g(3, 3) + area * x * x
        Next i
        Dim system As CLinearSystem3x3
        Set system = New CLinearSystem3x3
        mFullSupport = system.Solve(g(1, 1), g(1, 2), g(1, 3), g(1, 2), g(2, 2), g(2, 3), _
            g(1, 3), g(2, 3), g(3, 3), 0#, 0#, 0#)
        mSupportChecked = True
    End If
    HasFullPlaneSupport = mFullSupport
End Function

' Переносит epsilon0 в начало численной системы, сохраняя то же линейное поле деформаций.
Friend Function CenterStrain(ByVal eps0 As Double, ByVal kx As Double, ByVal ky As Double) As Double
    CenterStrain = eps0 + kx * mOriginY + ky * mOriginX
End Function

' Возвращает epsilon0 в исходные координаты сечения без изменения поля деформаций.
Friend Function OriginalStrain(ByVal epsCenter As Double, ByVal kx As Double, ByVal ky As Double) As Double
    OriginalStrain = epsCenter - kx * mOriginY - ky * mOriginX
End Function

' Нагрузка переносится тем же правилом, что и интеграл внутренних усилий.
Friend Function CenterMx(ByVal n As Double, ByVal mx As Double) As Double
    CenterMx = mx - n * mOriginY
End Function

' Переносит My к численному началу с учётом осевой силы N и смещения по X.
Friend Function CenterMy(ByVal n As Double, ByVal my As Double) As Double
    CenterMy = my - n * mOriginX
End Function

' Возвращает Mx к исходному началу с учётом N и смещения по Y.
Friend Function OriginalMx(ByVal n As Double, ByVal mx As Double) As Double
    OriginalMx = mx + n * mOriginY
End Function

' Возвращает My к исходному началу с учётом N и смещения по X.
Friend Function OriginalMy(ByVal n As Double, ByVal my As Double) As Double
    OriginalMy = my + n * mOriginX
End Function

' Координата X численного начала относительно исходной модели, мм.
Friend Property Get OriginX() As Double
    OriginX = mOriginX
End Property

' Координата Y численного начала относительно исходной модели, мм.
Friend Property Get OriginY() As Double
    OriginY = mOriginY
End Property

' Положительный масштаб длины для нормирования кривизн, мм.
Friend Property Get LengthScale() As Double
    LengthScale = mLength
End Property
