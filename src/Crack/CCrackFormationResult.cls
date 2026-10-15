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
' Результат определения нагрузки образования нормальной трещины
' ==========================================================================
' Хранит достигнутую нагрузку образования трещины, фактический способ её поиска
' и доступные состояния PreCrackState/PostCrackState для расчёта ширины раскрытия.
' Статус результата передаётся через ResultMeta; его видимое обозначение
' формирует CResultStatusPolicy.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mMeta As CResultMeta ' Статус и код причины результата определения нагрузки образования трещины.
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
Private mStopReason As String ' Краткая причина остановки определения нагрузки образования трещины.
Private mDiagnosticLog As String ' Подробный журнал поиска нагрузки образования трещины.
Private mHasLimitPoint As Boolean ' Признак подтверждённой предельной точки; предупреждение без найденной точки этот признак не устанавливает.
Private mFallbackPsi1 As Boolean ' Фиксированный путь допускает ширину с psi=1 без PostState.
Private mCentralTensionBranch As Boolean ' Для ширины применима ветка полного растяжения.
Private mSolverCallCount As Long ' Число служебных решений НДС при проверке образования трещины.
Private mSearchResult As CLimitSearchResult ' Независимый численный снимок найденной точки.

' Создает нейтральный not-applicable результат до фактического заполнения.
Private Sub Class_Initialize()
    Set mMeta = New CResultMeta
    mMeta.SetNotApplicable rkCrackFormation
End Sub

' Копирует готовый итог проверки образования трещины: нагрузки, доступные состояния,
' статус и пояснение. Физическую допустимость точки определяет сам расчёт.
Public Sub InitializeFromCalculator(ByVal crack As CCrackFormationCalculator, ByVal meta As CResultMeta)
    AssertWritable
    ResetResult
    If meta Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, _
            "Meta результата образования трещины не сформирована.", vbNullString, True, False
    Else
        Set mMeta = meta.Clone
    End If

    If crack Is Nothing Then
        mMeta.SetResult rsInternalError, rcInternalError, rkCrackFormation, "Не передан расчет образования трещины.", _
            vbNullString, True, False
        Exit Sub
    End If
    mHasLimitPoint = crack.HasLimitPoint
    mFallbackPsi1 = crack.FallbackPsi1
    mCentralTensionBranch = crack.CentralTensionBranch
    mSolverCallCount = crack.SolverCallCount
    Set mSearchResult = crack.BuildSearchSnapshot(crack.CrackFormationSolutionStrategy)
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

' Сохраняет причину, по которой поиск не запускался: отключённая проверка,
' ошибка ввода или отсутствие обязательных данных. Прежние численные данные
' удаляются; вымышленные пробные нагрузки и НДС не создаются.
Friend Sub InitializeUncalculated(ByVal meta As CResultMeta)
    AssertWritable
    ResetResult
    If meta Is Nothing Then Err.Raise vbObjectError + 3944, "CCrackFormationResult", "Не передана причина отсутствия расчета образования трещины."
    If meta.Calculated Then Err.Raise vbObjectError + 3944, "CCrackFormationResult", "Незапущенный поиск не может иметь признак выполненного расчета."
    Set mMeta = meta.Clone
    mStopReason = meta.ResultComment
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

' Возвращает независимую копию статуса и пояснения образования трещины.
' Изменение копии не меняет сохранённый результат.
Public Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' True, если нормальная трещина образована при текущем LC.
Public Property Get CrackFormed() As Boolean
    CrackFormed = mCrackFormed
End Property

' Подтверждает факт образования трещины по результату её проверки.
' Найденный порог сохраняется при неудаче НДС после трещины.
' Трещина может быть подтверждена и в постоянной нагрузке без отдельной предельной точки.
Public Property Get ConfirmedCracked() As Boolean
    ConfirmedCracked = mCrackFormed And (mHasLimitPoint Or mFallbackPsi1)
End Property

' Различает подтверждённое отсутствие трещины и неудачный поиск порога.
' Найденный порог выше текущей нагрузки требует подтверждённой точки;
' аналитическое недостижение критерия допустимо без состояний до и после трещины.
' Ошибка или граница поиска отсутствия трещины не доказывают: остаётся расчёт с psi_s = 1.
Public Property Get ConfirmedNotCracked() As Boolean
    If mCrackFormed Then Exit Property
    If mMeta.InternalStatus <> rsSuccess And mMeta.InternalStatus <> rsSuccessWithWarning Then Exit Property
    Select Case mMeta.ResultCode
        Case rcCriterionNotReached, rcCrackNotFormed
            ConfirmedNotCracked = True
        Case rcCheckPassed
            ConfirmedNotCracked = mHasLimitPoint
    End Select
End Property

' Фактически принятый путь lambda-поиска.
Public Property Get FormationMethod() As String
    FormationMethod = mFormationMethod
End Property

' Найденный множитель образования трещины.
Public Property Get LambdaCrc() As Double
    LambdaCrc = mLambdaCrc
End Property

' Ncrc для центрального растяжения, если применимо.
Public Property Get Ncrc() As Double
    Ncrc = mNcrc
End Property

' Полная N в найденной общей точке образования трещины.
Public Property Get FormationNcrc() As Double
    FormationNcrc = mFormationNcrc
End Property

' Модуль Mxy,crc в принятой точке отсчета.
Public Property Get Mcrc() As Double
    Mcrc = mMcrc
End Property

' Приведенная площадь для ветки центрального растяжения.
Public Property Get Ared() As Double
    Ared = mAred
End Property

' Rbt,ser, использованный для Ncrc.
Public Property Get RbtSer() As Double
    RbtSer = mRbtSer
End Property

' Состояние до выключения растянутого бетона.
Public Property Get PreCrackState() As CSectionStateResult
    Set PreCrackState = mPreCrackState
End Property

' Состояние после выключения растянутого бетона.
Public Property Get PostCrackState() As CSectionStateResult
    Set PostCrackState = mPostCrackState
End Property

' Возвращает краткую причину остановки определения нагрузки образования трещины.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Возвращает подробный журнал поиска нагрузки образования трещины.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Показывает доступность найденной точки независимо от успешного завершения процедуры.
Public Property Get HasLimitPoint() As Boolean
    HasLimitPoint = mHasLimitPoint
End Property

' Фиксированный путь допускает ширину с psi=1 без PostState.
Public Property Get FallbackPsi1() As Boolean
    FallbackPsi1 = mFallbackPsi1
End Property

' Для ширины применима ветка полного растяжения.
Public Property Get CentralTensionBranch() As Boolean
    CentralTensionBranch = mCentralTensionBranch
End Property

' Возвращает число служебных решений НДС при проверке образования трещины.
Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

' Возвращает отдельный объект результата поиска, в том числе при аналитической Ncrc
' или раннем выходе. SearchExecuted показывает, был ли численный поиск.
' Nothing означает отсутствие заполненного результата; конечные НДС не копируются по элементам.
Public Property Get SearchResult() As CLimitSearchResult
    If Not mSearchResult Is Nothing Then Set SearchResult = mSearchResult.Clone
End Property

' Показывает успешность процедуры по её статусу. Это не подтверждает само по себе
' наличие предельной точки или пригодность состояний до и после трещины.
Public Property Get Converged() As Boolean
    Converged = (mMeta.InternalStatus = rsSuccess Or mMeta.InternalStatus = rsSuccessWithWarning)
End Property

' Завершает заполнение результата и найденных состояний до и после трещины.
' После этого их нельзя менять. Отсутствующие НДС не создаются; повторный вызов безопасен.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    If Not mPreCrackState Is Nothing Then mPreCrackState.Freeze
    If Not mPostCrackState Is Nothing Then mPostCrackState.Freeze
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CCrackFormationResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub
