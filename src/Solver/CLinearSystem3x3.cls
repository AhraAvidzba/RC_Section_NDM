VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinearSystem3x3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSolved As Boolean
Private mX1 As Double
Private mX2 As Double
Private mX3 As Double
Private mDeterminant As Double
Private mAbsoluteResidual As Double
Private mRelativeResidual As Double
Private mStopReason As String

Public Function Solve(ByVal a11 As Double, ByVal a12 As Double, ByVal a13 As Double, _
        ByVal a21 As Double, ByVal a22 As Double, ByVal a23 As Double, _
        ByVal a31 As Double, ByVal a32 As Double, ByVal a33 As Double, _
        ByVal b1 As Double, ByVal b2 As Double, ByVal b3 As Double) As Boolean

    Dim a(1 To 3, 1 To 4) As Double
    a(1, 1) = a11: a(1, 2) = a12: a(1, 3) = a13: a(1, 4) = b1
    a(2, 1) = a21: a(2, 2) = a22: a(2, 3) = a23: a(2, 4) = b2
    a(3, 1) = a31: a(3, 2) = a32: a(3, 3) = a33: a(3, 4) = b3

    mSolved = False
    mX1 = 0#: mX2 = 0#: mX3 = 0#
    mDeterminant = 1#
    mAbsoluteResidual = 0#
    mRelativeResidual = 0#
    mStopReason = vbNullString

    Dim col As Long
    Dim row As Long
    Dim pivotRow As Long
    Dim pivotAbs As Double
    Dim candidateAbs As Double
    Dim tmp As Double
    Dim factor As Double
    Dim sign As Double
    sign = 1#

    For col = 1 To 3
        pivotRow = col
        pivotAbs = Abs(a(col, col))
        For row = col + 1 To 3
            candidateAbs = Abs(a(row, col))
            If candidateAbs > pivotAbs Then
                pivotAbs = candidateAbs
                pivotRow = row
            End If
        Next row

        If pivotAbs <= 1E-18 Then
            mStopReason = "Singular matrix."
            Solve = False
            Exit Function
        End If

        If pivotRow <> col Then
            Dim j As Long
            For j = col To 4
                tmp = a(col, j)
                a(col, j) = a(pivotRow, j)
                a(pivotRow, j) = tmp
            Next j
            sign = -sign
        End If

        mDeterminant = mDeterminant * a(col, col)

        For row = col + 1 To 3
            factor = a(row, col) / a(col, col)
            a(row, col) = 0#
            For j = col + 1 To 4
                a(row, j) = a(row, j) - factor * a(col, j)
            Next j
        Next row
    Next col

    mDeterminant = mDeterminant * sign

    mX3 = a(3, 4) / a(3, 3)
    mX2 = (a(2, 4) - a(2, 3) * mX3) / a(2, 2)
    mX1 = (a(1, 4) - a(1, 2) * mX2 - a(1, 3) * mX3) / a(1, 1)

    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    r1 = a11 * mX1 + a12 * mX2 + a13 * mX3 - b1
    r2 = a21 * mX1 + a22 * mX2 + a23 * mX3 - b2
    r3 = a31 * mX1 + a32 * mX2 + a33 * mX3 - b3
    mAbsoluteResidual = Abs(r1) + Abs(r2) + Abs(r3)

    Dim bNorm As Double
    bNorm = Abs(b1) + Abs(b2) + Abs(b3)
    If bNorm > 1E-18 Then
        mRelativeResidual = mAbsoluteResidual / bNorm
    Else
        mRelativeResidual = mAbsoluteResidual
    End If

    If mRelativeResidual > 1E-8 Then
        mStopReason = "Residual check failed."
        Solve = False
        Exit Function
    End If

    mSolved = True
    mStopReason = "Solved"
    Solve = True
End Function

Public Property Get Solved() As Boolean
    Solved = mSolved
End Property

Public Property Get X1() As Double
    X1 = mX1
End Property

Public Property Get X2() As Double
    X2 = mX2
End Property

Public Property Get X3() As Double
    X3 = mX3
End Property

Public Property Get Determinant() As Double
    Determinant = mDeterminant
End Property

Public Property Get AbsoluteResidual() As Double
    AbsoluteResidual = mAbsoluteResidual
End Property

Public Property Get RelativeResidual() As Double
    RelativeResidual = mRelativeResidual
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property






