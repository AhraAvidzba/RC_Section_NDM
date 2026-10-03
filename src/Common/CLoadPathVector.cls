VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadPathVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Универсальная lambda-траектория нагрузки
' ==========================================================================
' CLoadPathVector хранит только математическую форму нагрузки:
'
'     Target(lambda) = Offset + lambda * Base
'
' Класс не знает о несущей способности, трещинах, пользовательских ключах
' LoadPath и нормативных ветках. Его задача - быть единственным
' местом, где для расчетного ядра собираются N, Mx и My по заданному lambda.
' Более высокие классы сами решают, что физически означает Offset/Base.

Private mNOffset As Double ' Постоянная часть N, Н.
Private mNBase As Double ' Масштабируемая часть N, Н.
Private mMxOffset As Double ' Постоянная часть Mx относительно начала модели, Н*мм.
Private mMxBase As Double ' Масштабируемая часть Mx, Н*мм.
Private mMyOffset As Double ' Постоянная часть My относительно начала модели, Н*мм.
Private mMyBase As Double ' Масштабируемая часть My, Н*мм.

' Задает траекторию напрямую через постоянные и масштабируемые компоненты.
Public Sub Initialize(ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double)
    mNOffset = nOffset
    mNBase = nBase
    mMxOffset = mxOffset
    mMxBase = mxBase
    mMyOffset = myOffset
    mMyBase = myBase
End Sub

' Задает траекторию с постоянной продольной силой и масштабируемым моментом
' относительно произвольной расчетной точки. Solver работает относительно
' начала CSectionModel, поэтому момент от N до этой точки входит в Offset.
Public Sub InitializeFixedNAndMomentAboutPoint(ByVal nValue As Double, _
        ByVal momentMxAboutPoint As Double, ByVal momentMyAboutPoint As Double, _
        ByVal pointX As Double, ByVal pointY As Double)
    Initialize nValue, 0#, _
        nValue * pointY, momentMxAboutPoint, _
        nValue * pointX, momentMyAboutPoint
End Sub

' Возвращает полный расчетный вектор для заданного lambda.
Public Sub Target(ByVal lambdaValue As Double, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double)
    nValue = TargetN(lambdaValue)
    mxValue = TargetMx(lambdaValue)
    myValue = TargetMy(lambdaValue)
End Sub

' Возвращает N(lambda).
Public Function TargetN(ByVal lambdaValue As Double) As Double
    TargetN = mNOffset + lambdaValue * mNBase
End Function

' Возвращает Mx(lambda).
Public Function TargetMx(ByVal lambdaValue As Double) As Double
    TargetMx = mMxOffset + lambdaValue * mMxBase
End Function

' Возвращает My(lambda).
Public Function TargetMy(ByVal lambdaValue As Double) As Double
    TargetMy = mMyOffset + lambdaValue * mMyBase
End Function

Public Property Get NOffset() As Double
    NOffset = mNOffset
End Property

Public Property Get NBase() As Double
    NBase = mNBase
End Property

Public Property Get MxOffset() As Double
    MxOffset = mMxOffset
End Property

Public Property Get MxBase() As Double
    MxBase = mMxBase
End Property

Public Property Get MyOffset() As Double
    MyOffset = mMyOffset
End Property

Public Property Get MyBase() As Double
    MyBase = mMyBase
End Property
