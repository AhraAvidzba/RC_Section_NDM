Attribute VB_Name = "modTorsionPostTypes"
Option Explicit

' ==========================================================================
' Данные отдельной приближённой проверки арматуры и хомута на кручение
' ==========================================================================
' Не являются нагрузками решателя НДМ. Размерности: мм, Н, Н*мм, МПа.
' Проверочная сумма напряжений не задаёт новое равновесное НДС.

Public Const TORSION_POST_COLUMNS As Long = 39
Public Const TORSION_POST_DISPLAY_COLUMNS As Long = 43 ' Отдельный блок с ResultComment, нагрузками и тремя проверками.
Public Const TORSION_BAR_COLUMNS As Long = 4 ' ID сочетания, стержень, усилие от T и соответствующая длина оси хомута.
Public Const TORSION_BATCH_STATUS_COLUMN As Long = 16
Public Const TORSION_BATCH_RESERVE_COLUMN As Long = 30
Public Const TORSION_POST_DATA_KEY As String = "TorsionPost.Data"
Public Const TORSION_POST_HEADERS_KEY As String = "TorsionPost.Headers"
Public Const TORSION_BAR_DATA_KEY As String = "TorsionPost.Bars"

Public Type TTorsionPostParameters
    StateName As String
    AxisOffset As Double
    HoopDiameter As Double
    HoopSpacing As Double
    SteelResistance As Double
    HoopResistance As Double
    BarBand As Double
    StrutAngle As Double
    ConcreteMethod As String
    ConcreteInputInvalid As Boolean ' Ошибка только исходных данных проверки бетона; арматура и хомут рассчитываются независимо.
    WallThickness As Double
    ConcreteResistance As Double
    ConcreteEfficiency As Double
    GorskyCoefficient As Double
End Type

' Возвращает внешний статус применимой проверки по общей политике проекта.
Public Function TorsionPostCheckStatus(ByVal reserve As Double) As String
    Dim meta As New CResultMeta, policy As New CResultStatusPolicy
    If reserve >= 1# Then
        meta.SetResult rsSuccess, rcCheckPassed, rkGeneric
    Else
        meta.SetResult rsCheckFailed, rcCheckFailed, rkGeneric
    End If
    TorsionPostCheckStatus = policy.ExternalStatus(meta)
End Function
