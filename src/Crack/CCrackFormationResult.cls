VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrackFormationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Результат проверки образования нормальной трещины
' ==========================================================================
' CCrackFormationResult отделяет search-этап Ncrc/Mcrc от расчета ширины.
' Здесь хранится только итог проверки образования трещины: была ли достигнута
' расчетная точка, какой путь сработал, какие усилия найдены и какие named-state
' PreCrackState/PostCrackState доступны downstream-слоям.
'
' Класс не запускает solver и не решает, как статус будет выглядеть на листе.
' Внешний OK/FAIL/NumFail/N/A формирует CResultStatusPolicy по ResultMeta.

Private mMeta As CResultMeta ' Типизированный статус search-этапа образования трещины.
Private mCrackFormed As Boolean ' True, если нормальная трещина образована при текущем LC.
Private mFormationMethod As String ' Фактически принятый путь lambda-поиска.
Private mLambdaCrc As Double ' Найденный множитель образования трещины.
Private mNcrc As Double ' Ncrc для центрального растяжения, если применимо.
Private mFormationNcrc As Double ' Полная N в найденной общей точке образования трещины.
Private mMcrc As Double ' Модуль Mxy,crc в принятой точке отсчета.
Private mAred As Double ' Приведенная площадь для ветки центрального растяжения.
Private mRbtSer As Double ' Rbt,ser, использованный для Ncrc.
Private mPreCrackState As CSectionStateResult ' Состояние до выключения растянутого бетона.
Private mPostCrackState As CSectionStateResult ' Состояние после выключения растянутого бетона.
Private mStopReason As String ' Короткая причина остановки formation-этапа.
Private mDiagnosticLog As String ' Подробная диагностика search-этапа.
Private mHasLimitPoint As Boolean ' Точка реально найдена, а не заменена warning-fallback.
Private mFallbackPsi1 As Boolean ' Фиксированный путь допускает ширину с psi=1 без PostState.
Private mCentralTensionBranch As Boolean ' Для ширины применима ветка полного растяжения.
Private mSolverCallCount As Long ' Число solve-вызовов formation-сценария.
Private mSearchResult As CLimitSearchResult ' Независимый численный снимок найденной точки.

' Создает нейтральный not-applicable результат до фактического заполнения.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkCrackFormation
End Sub

' Заполняет результат по calculator-у образования трещины. Calculator отвечает
' за физику поиска, а этот объект фиксирует итоговые величины и typed status для batch.
Public Sub InitializeFromCalculator(ByVal crack As CCrackFormationCalculator, ByVal meta As CResultMeta)
    ResetResult
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, _
            "Meta результата образования трещины не сформирована."
    Else
        Set mMeta = meta.Clone
    End If

    If crack Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, "Не передан расчет образования трещины."
        Exit Sub
    End If
    mHasLimitPoint = crack.HasLimitPoint
    mFallbackPsi1 = crack.FallbackPsi1
    mCentralTensionBranch = crack.CentralTensionBranch
    mSolverCallCount = crack.SolverCallCount
    Set mSearchResult = New CLimitSearchResult
    mSearchResult.InitializeFromCrackFormation crack, crack.CrackFormationSolutionStrategy, crack.Converged
    mCrackFormed = crack.CrackFormed
    mFormationMethod = crack.CrackFormationMethod
    mLambdaCrc = crack.LambdaCrc
    mNcrc = crack.Ncrc
    mFormationNcrc = crack.FormationNcrc
    mMcrc = crack.Mcrc
    mAred = crack.Ared
    mRbtSer = crack.RbtSer
    Set mPreCrackState = crack.PreCrackState
    Set mPostCrackState = crack.PostCrackState
    mStopReason = crack.StopReason
    mDiagnosticLog = crack.DiagnosticLog
End Sub

' Удаляет все данные предыдущего заполнения; неуспешный повтор не оставляет
' старую точку, состояния или диагностический журнал доступными потребителям.
Private Sub ResetResult()
    mMeta.Clear
    mCrackFormed = False
    mFormationMethod = vbNullString
    mLambdaCrc = 0#
    mNcrc = 0#
    mFormationNcrc = 0#
    mMcrc = 0#
    mAred = 0#
    mRbtSer = 0#
    Set mPreCrackState = Nothing
    Set mPostCrackState = Nothing
    mStopReason = vbNullString
    mDiagnosticLog = vbNullString
    mHasLimitPoint = False
    mFallbackPsi1 = False
    mCentralTensionBranch = False
    mSolverCallCount = 0
    Set mSearchResult = Nothing
End Sub

Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

Public Property Get CrackFormed() As Boolean
    CrackFormed = mCrackFormed
End Property

Public Property Get FormationMethod() As String
    FormationMethod = mFormationMethod
End Property

Public Property Get LambdaCrc() As Double
    LambdaCrc = mLambdaCrc
End Property

Public Property Get Ncrc() As Double
    Ncrc = mNcrc
End Property

Public Property Get FormationNcrc() As Double
    FormationNcrc = mFormationNcrc
End Property

Public Property Get Mcrc() As Double
    Mcrc = mMcrc
End Property

Public Property Get Ared() As Double
    Ared = mAred
End Property

Public Property Get RbtSer() As Double
    RbtSer = mRbtSer
End Property

Public Property Get PreCrackState() As CSectionStateResult
    Set PreCrackState = mPreCrackState
End Property

Public Property Get PostCrackState() As CSectionStateResult
    Set PostCrackState = mPostCrackState
End Property

Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Доступность точки и fallback хранятся отдельно от статуса процедуры.
Public Property Get HasLimitPoint() As Boolean
    HasLimitPoint = mHasLimitPoint
End Property

Public Property Get FallbackPsi1() As Boolean
    FallbackPsi1 = mFallbackPsi1
End Property

Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property

Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

Public Property Get SearchResult() As CLimitSearchResult
    Set SearchResult = mSearchResult
End Property

' Успех formation включает корректный warning-fallback без фиктивной точки.
Public Property Get Converged() As Boolean
    Converged = (mMeta.InternalStatus = rsSuccess Or mMeta.InternalStatus = rsSuccessWithWarning)
End Property
