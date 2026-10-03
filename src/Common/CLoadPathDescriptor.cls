VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadPathDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Универсальная lambda-траектория поиска предельной точки
' ==========================================================================
' CLoadPathDescriptor переводит пользовательский выбор Auto, lambda*Mx,
' lambda*My, lambda*Mxy, lambda*N или lambda*NMxy в единую математическую форму
' общего LimitSearch для несущей способности и образования трещины:
'
'     Target(lambda) = Offset + lambda * Base
'
' Класс не ищет равновесие и не знает о диаграммах материалов. Его задача -
' один раз разобрать строку из таблицы сочетаний, отделить постоянную часть
' нагрузки от масштабируемой и сохранить служебные признаки, которые нужны
' batch-слою для выбора численной политики: есть ли вообще масштабируемая
' нагрузка и является ли путь фактически силовым осевым случаем.

Private Const KEY_NONE As String = "None" ' Начальное состояние описателя до назначения нагрузки.
Private Const KEY_AUTO As String = "Auto" ' Общий Search выбирает и перебирает конкретные пути.
Private Const KEY_MX As String = "LambdaMx" ' Масштабируется только пользовательский Mx.
Private Const KEY_MY As String = "LambdaMy" ' Масштабируется только пользовательский My.
Private Const KEY_MXY As String = "LambdaMxy" ' Масштабируется пользовательский вектор Mx/My.
Private Const KEY_N As String = "LambdaN" ' Масштабируется N вместе с моментом от эксцентриситета N.
Private Const KEY_NMXY As String = "LambdaNMxy" ' Масштабируется весь расчетный вектор N/Mx/My.

Private mRawText As String ' Исходная строка из таблицы сочетаний.
Private mKey As String ' Нормализованный внутренний ключ lambda-траектории.
Private mValid As Boolean ' False только при явном неизвестном значении пользователя.
Private mErrorMessage As String ' Человекочитаемая причина InputErr.
Private mHasPath As Boolean ' True, если путь удалось выбрать явно или по умолчанию.
Private mHasScaledLoad As Boolean ' True, если выбранный путь содержит ненулевую масштабируемую компоненту.
Private mForceOnly As Boolean ' True для lambda*N и lambda*NMxy без пользовательских Mx/My.
Private mScalesN As Boolean ' Масштабируется ли продольная сила N.
Private mScalesMx As Boolean ' Масштабируется ли пользовательский Mx.
Private mScalesMy As Boolean ' Масштабируется ли пользовательский My.

Private mPath As CLoadPathVector ' Общий вектор Offset + lambda*Base для передачи в solver.

' Задает безопасное начальное состояние: путь отсутствует, объект вектора уже создан.
Private Sub Class_Initialize()
    mValid = True
    mKey = KEY_NONE
    Set mPath = New CLoadPathVector
End Sub

' Разбирает load-path через общий объект нагрузки. Так перенос от
' Load.ReferenceOffset и фильтр малых моментов остаются в одном классе, а
' LoadPath занимается только разложением Target = Offset + lambda*Base.
Public Sub InitializeFromLoadState(ByVal rawPath As String, ByVal loadState As CSectionLoadState)
    If loadState Is Nothing Then _
        Err.Raise vbObjectError + 3821, "CLoadPathDescriptor", "Не передан CSectionLoadState."

    Clear
    mRawText = Trim$(rawPath)

    Dim normalizedUserMx As Double
    Dim normalizedUserMy As Double
    normalizedUserMx = loadState.LoadPointMx
    normalizedUserMy = loadState.LoadPointMy

    mKey = NormalizeKey(mRawText)

    If Len(mKey) = 0 Then
        If Len(mRawText) > 0 Then
            mValid = False
            mErrorMessage = "LoadPath должен быть Auto, lambda*Mx, lambda*My, lambda*Mxy, lambda*N или lambda*NMxy. Пустое значение означает Auto."
        Else
            mKey = KEY_NONE
        End If
        Exit Sub
    End If

    mHasPath = (mKey <> KEY_NONE)
    If mKey = KEY_AUTO Then mHasPath = (loadState.N <> 0# Or normalizedUserMx <> 0# Or normalizedUserMy <> 0#)
    If Not mHasPath Then Exit Sub

    BuildFlags loadState.N, normalizedUserMx, normalizedUserMy
    BuildOffsetBase loadState, normalizedUserMx, normalizedUserMy
End Sub

' Нормализует общий выбор пути без выбора физической траектории Auto.
' Пустой ввод разрешен и означает Auto; неизвестное значение возвращает
' пустой ключ для адресной ошибки вызывающего Excel-слоя.
Public Function NormalizeKey(ByVal rawPath As String) As String
    Dim value As String
    value = LCase$(Trim$(rawPath))
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW$(&H3BB), "lambda")

    Select Case value
        Case "lambda*mx", "lambdamx", "mx"
            NormalizeKey = KEY_MX
        Case "lambda*my", "lambdamy", "my"
            NormalizeKey = KEY_MY
        Case "lambda*mxy", "lambdamxy", "mxy"
            NormalizeKey = KEY_MXY
        Case "lambda*n", "lambdan", "nload", "n"
            NormalizeKey = KEY_N
        Case "lambda*nmxy", "lambdanmxy", "nmxy", "all"
            NormalizeKey = KEY_NMXY
        Case "auto", vbNullString
            NormalizeKey = KEY_AUTO
    End Select
End Function

' Заполняет логические признаки выбранного пути после общей нормализации.
' Дополнительный порог здесь не применяется: ненулевая компонента остается
' масштабируемой, даже когда для достижения критерия требуется большая lambda.
' Нулевые компоненты не создают отдельную ветку: они просто дают нулевой Base,
' а CCapacitySolver дальше получает ту же универсальную формулу Target(lambda).
Private Sub BuildFlags(ByVal nValue As Double, ByVal userMxValue As Double, ByVal userMyValue As Double)
    Select Case mKey
        Case KEY_MX
            mScalesMx = True
            mHasScaledLoad = (userMxValue <> 0#)
        Case KEY_MY
            mScalesMy = True
            mHasScaledLoad = (userMyValue <> 0#)
        Case KEY_MXY
            mScalesMx = True
            mScalesMy = True
            mHasScaledLoad = HasMomentVector(userMxValue, userMyValue)
        Case KEY_N
            mScalesN = True
            mHasScaledLoad = (nValue <> 0#)
        Case KEY_NMXY
            mScalesN = True
            mScalesMx = True
            mScalesMy = True
            mHasScaledLoad = (nValue <> 0# Or HasMomentVector(userMxValue, userMyValue))
        Case KEY_AUTO
            mHasScaledLoad = (nValue <> 0# Or userMxValue <> 0# Or userMyValue <> 0#)
    End Select

    mForceOnly = (mKey = KEY_N Or mKey = KEY_NMXY) And Not HasMomentVector(userMxValue, userMyValue)
End Sub

' Строит постоянную и масштабируемую части нагрузки.
' Момент от смещения точки приложения N остается связанным с N: если N
' масштабируется, этот момент масштабируется вместе с ней; если масштабируются
' только Mx/My, момент от N остается в Offset.
Private Sub BuildOffsetBase(ByVal loadState As CSectionLoadState, ByVal userMxValue As Double, _
        ByVal userMyValue As Double)
    Dim nValue As Double
    Dim mxFromN As Double
    Dim myFromN As Double
    nValue = loadState.N
    mxFromN = loadState.AxialMxAboutPoint(0#)
    myFromN = loadState.AxialMyAboutPoint(0#)

    Select Case mKey
        Case KEY_MX
            mPath.Initialize nValue, 0#, mxFromN, userMxValue, loadState.InternalMy, 0#
        Case KEY_MY
            mPath.Initialize nValue, 0#, loadState.InternalMx, 0#, myFromN, userMyValue
        Case KEY_MXY
            mPath.Initialize nValue, 0#, mxFromN, userMxValue, myFromN, userMyValue
        Case KEY_N
            mPath.Initialize 0#, nValue, userMxValue, mxFromN, userMyValue, myFromN
        Case KEY_NMXY, KEY_AUTO
            mPath.Initialize 0#, nValue, 0#, loadState.InternalMx, 0#, loadState.InternalMy
    End Select
End Sub

' Проверяет наличие ненулевого пользовательского вектора момента после общего фильтра малых моментов.
Private Function HasMomentVector(ByVal mxValue As Double, ByVal myValue As Double) As Boolean
    HasMomentVector = (mxValue <> 0# Or myValue <> 0#)
End Function

' Сбрасывает разобранный путь перед чтением новой строки сочетания нагрузок.
Private Sub Clear()
    mRawText = vbNullString
    mKey = KEY_NONE
    mValid = True
    mErrorMessage = vbNullString
    mHasPath = False
    mHasScaledLoad = False
    mForceOnly = False
    mScalesN = False
    mScalesMx = False
    mScalesMy = False
    Set mPath = New CLoadPathVector
End Sub

' Возвращает подпись load-path в том же виде, в каком пользователь видит ее
' на листе Results. Метод полезен writer-ам и batch-слою, но не участвует в
' расчетной математике.
Public Function DisplayNameForKey(ByVal loadPathKey As String) As String
    Select Case loadPathKey
        Case KEY_MX
            DisplayNameForKey = ChrW$(&H3BB) & "*Mx"
        Case KEY_MY
            DisplayNameForKey = ChrW$(&H3BB) & "*My"
        Case KEY_MXY
            DisplayNameForKey = ChrW$(&H3BB) & "*Mxy"
        Case KEY_N
            DisplayNameForKey = ChrW$(&H3BB) & "*N"
        Case KEY_NMXY
            DisplayNameForKey = ChrW$(&H3BB) & "*NMxy"
        Case Else
            DisplayNameForKey = loadPathKey
    End Select
End Function

Public Property Get Key() As String
    Key = mKey
End Property

Public Property Get DisplayName() As String
    DisplayName = DisplayNameForKey(mKey)
End Property

Public Property Get Valid() As Boolean
    Valid = mValid
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Public Property Get HasPath() As Boolean
    HasPath = mHasPath
End Property

Public Property Get HasScaledLoad() As Boolean
    HasScaledLoad = mHasScaledLoad
End Property

Public Property Get ForceOnly() As Boolean
    ForceOnly = mForceOnly
End Property

Public Property Get ScalesN() As Boolean
    ScalesN = mScalesN
End Property

Public Property Get ScalesMx() As Boolean
    ScalesMx = mScalesMx
End Property

Public Property Get ScalesMy() As Boolean
    ScalesMy = mScalesMy
End Property

Public Property Get NOffset() As Double
    NOffset = mPath.NOffset
End Property

Public Property Get NBase() As Double
    NBase = mPath.NBase
End Property

Public Property Get MxOffset() As Double
    MxOffset = mPath.MxOffset
End Property

Public Property Get MxBase() As Double
    MxBase = mPath.MxBase
End Property

Public Property Get MyOffset() As Double
    MyOffset = mPath.MyOffset
End Property

Public Property Get MyBase() As Double
    MyBase = mPath.MyBase
End Property

' Возвращает расчетный вектор нагрузки на выбранной lambda без дублирования
' формулы Offset + lambda*Base в batch-слое.
Public Sub Target(ByVal lambdaValue As Double, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double)
    mPath.Target lambdaValue, nValue, mxValue, myValue
End Sub

' Возвращает отдельный математический вектор выбранной траектории.
' Search получает снимок, поэтому его пробные точки не меняют описатель LC.
Public Function Vector() As CLoadPathVector
    Dim result As CLoadPathVector
    Set result = New CLoadPathVector
    result.Initialize NOffset, NBase, MxOffset, MxBase, MyOffset, MyBase
    Set Vector = result
End Function
