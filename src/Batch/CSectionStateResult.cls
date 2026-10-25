VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionStateResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Сохранённое напряжённо-деформированное состояние сечения
' ==========================================================================
' Для одного сочетания сохраняются состояния по прочности, несущей способности
' и трещинам, включая состояния до и после образования трещины.
' Объект хранит поле деформаций, усилия, невязки, диагностику решения равновесия
' и фактически использованную модель материалов для последующей записи в Results.

Private mFrozen As Boolean ' После передачи готового результата его заполнение и очистка запрещены; сохранённые числа доступны для чтения.
Private mStateType As ESectionStateType ' Какое именованное состояние рассчитано.
Private mMaterialModelRole As ECalculationPurpose ' Назначение модели материалов: прочность, образование трещины или НДС с трещинами.
Private mMaterialSpec As CMaterialModelSpec ' Фактически выбранные параметры модели материалов расчётного профиля.
Private mEpsilon0 As Double ' Деформация в начале координат расчетной модели.
Private mKappaX As Double ' Кривизна kappaX, 1/мм.
Private mKappaY As Double ' Кривизна kappaY, 1/мм.
Private mTargetN As Double ' Продольная сила, для которой требовалось найти НДС, Н.
Private mTargetMx As Double ' Момент Mx, для которого требовалось найти НДС, Н*мм.
Private mTargetMy As Double ' Момент My, для которого требовалось найти НДС, Н*мм.
Private mNint As Double ' Внутреннее N, полученное интегрированием напряжений, Н.
Private mMxint As Double ' Внутренний Mx, полученный интегрированием напряжений, Н*мм.
Private mMyint As Double ' Внутренний My, полученный интегрированием напряжений, Н*мм.
Private mResidualN As Double ' Невязка Nint - Ntarget, Н.
Private mResidualMx As Double ' Невязка Mxint - Mxtarget, Н*мм.
Private mResidualMy As Double ' Невязка Myint - Mytarget, Н*мм.
Private mRelativeResidualN As Double ' Относительная невязка по N для диагностики качества состояния.
Private mRelativeResidualMx As Double ' Относительная невязка по Mx для диагностики качества состояния.
Private mRelativeResidualMy As Double ' Относительная невязка по My для диагностики качества состояния.
Private mLastResidualNorm As Double ' Общая мера невязок равновесия последней попытки решения.
Private mExtensionUsed As Boolean ' Показывает, вышло ли найденное НДС на вспомогательное продолжение диаграмм за физическими пределами.
Private mWithinPhysicalRange As Boolean ' Показывает, находится ли найденное НДС в физических пределах диаграмм материалов.
Private mConverged As Boolean ' Показывает, найдено ли равновесие в попытке, из которой сохранено поле деформаций.
Private mStopReason As String ' Причина остановки попытки решения равновесия.
Private mResultMeta As CResultMeta ' Статус и код причины сохранённого расчётного состояния.
Private mSolverCallCount As Long ' Число вызовов CSectionSolver.Solve при получении этого НДС.
Private mIterationCount As Long ' Число итераций последней попытки, из которой получено это НДС.
Private mDiagnosticLog As String ' Сохранённый журнал поиска равновесия для последующего вывода результатов.
Private mMinConcreteStrain As Double ' Минимальная деформация бетона в сохранённом НДС.
Private mMaxConcreteStrain As Double ' Максимальная деформация бетона в сохранённом НДС.
Private mMinSteelStrain As Double ' Минимальная деформация арматуры в сохранённом НДС.
Private mMaxSteelStrain As Double ' Максимальная деформация арматуры в сохранённом НДС.
Private mMinConcreteStress As Double ' Наиболее сжимающее напряжение бетона со знаком.
Private mMaxConcreteStress As Double ' Наиболее растягивающее напряжение бетона со знаком.
Private mMinSteelStress As Double ' Наиболее сжимающее напряжение арматуры со знаком.
Private mMaxSteelStress As Double ' Наиболее растягивающее напряжение арматуры со знаком.
Private mHasConcreteCompression As Boolean ' True, если в состоянии реально есть сжатый бетон.
Private mHasConcreteTension As Boolean ' True, если бетон реально работает на растяжение.
Private mHasSteelCompression As Boolean ' True, если есть сжатая арматура.
Private mHasSteelTension As Boolean ' True, если есть растянутая арматура.
Private mSnapshotSection As CSectionModel ' Модель сечения, по которой найдено НДС; проверка версии выявляет последующее изменение геометрии.
Private mSnapshotRevision As Long ' Ревизия геометрии при фиксации состояния.
Private mSnapshotConcrete As CMaterialDiagram ' Неизменяемая фактическая диаграмма бетона.
Private mSnapshotSteel As CMaterialDiagram ' Неизменяемая фактическая диаграмма арматуры.
Private Const ELEMENT_ZERO_TOLERANCE As Double = 0.000000001 ' Порог деформации для определения физического состояния элемента при выводе результатов.

' Сохраняет расчётное НДС по результату CSectionSolver.
Public Sub InitializeFromSolver(ByVal stateType As ESectionStateType, _
        ByVal materialModelRole As ECalculationPurpose, ByVal materialSpec As CMaterialModelSpec, _
        ByVal solver As CSectionSolver, Optional ByVal extensionUsed As Boolean = False, _
        Optional ByVal withinPhysicalRange As Boolean = True, _
        Optional ByVal solverCallCount As Long = 0, _
        Optional ByVal diagnosticLog As String = vbNullString)
    AssertWritable
    If materialSpec Is Nothing Then Err.Raise vbObjectError + 3970, "CSectionStateResult", "Не передано описание выбранных диаграмм и характеристик материалов (MaterialModelSpec)."
    If solver Is Nothing Then Err.Raise vbObjectError + 3971, "CSectionStateResult", "Решатель для именованного состояния не передан."

    ClearSnapshot
    mStateType = stateType
    mMaterialModelRole = materialModelRole
    Set mMaterialSpec = materialSpec.Clone
    mEpsilon0 = solver.Epsilon0
    mKappaX = solver.KappaX
    mKappaY = solver.KappaY
    CaptureEquilibriumData solver
    mExtensionUsed = extensionUsed
    mWithinPhysicalRange = (withinPhysicalRange And solver.Converged)
    mConverged = solver.Converged
    mStopReason = solver.StopReason
    Set mResultMeta = BuildMetaFromSolver(stateType, solver, extensionUsed, withinPhysicalRange, mStopReason)
    mSolverCallCount = solverCallCount
    mIterationCount = solver.Iterations
    mDiagnosticLog = diagnosticLog
    mMinConcreteStrain = solver.MinConcreteStrain
    mMaxConcreteStrain = solver.MaxConcreteStrain
    mMinSteelStrain = solver.MinSteelStrain
    mMaxSteelStrain = solver.MaxSteelStrain
    mMinConcreteStress = solver.MinConcreteStress
    mMaxConcreteStress = solver.MaxConcreteStress
    mMinSteelStress = solver.MinSteelStress
    mMaxSteelStress = solver.MaxSteelStress
    mHasConcreteCompression = solver.HasConcreteCompression
    mHasConcreteTension = solver.HasConcreteTension
    mHasSteelCompression = solver.HasSteelCompression
    mHasSteelTension = solver.HasSteelTension
    solver.CopyMaterialContext mSnapshotSection, mSnapshotConcrete, mSnapshotSteel
    If Not mSnapshotSection Is Nothing Then mSnapshotRevision = mSnapshotSection.Revision
End Sub

' Сохраняет причину, по которой решение не удалось начать: обязательные программные
' данные не переданы. Такой результат не используется как ранее найденное НДС.
Public Sub InitializeFailure(ByVal request As CStateRequest, ByVal meta As CResultMeta)
    AssertWritable
    If request Is Nothing Then Err.Raise vbObjectError + 3972, "CSectionStateResult", "Запрос диагностического состояния не передан."
    If meta Is Nothing Then Err.Raise vbObjectError + 3973, "CSectionStateResult", "Причина отсутствия состояния не передана."
    ClearSnapshot
    mStateType = request.StateType
    mMaterialModelRole = request.MaterialRole
    Set mMaterialSpec = request.MaterialSpec
    mTargetN = request.TargetN
    mTargetMx = request.TargetMx
    mTargetMy = request.TargetMy
    Set mResultMeta = meta.Clone
    mStopReason = meta.ResultComment
End Sub

' Очищает физические величины и диагностику перед заполнением результата.
' Если новый расчёт завершился ошибкой, числа прежнего НДС не должны оставаться в нём.
Private Sub ClearSnapshot()
    Set mSnapshotSection = Nothing
    Set mSnapshotConcrete = Nothing
    Set mSnapshotSteel = Nothing
    mSnapshotRevision = 0
    Set mMaterialSpec = Nothing
    Set mResultMeta = Nothing
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mTargetN = 0#: mTargetMx = 0#: mTargetMy = 0#
    mNint = 0#: mMxint = 0#: mMyint = 0#
    mResidualN = 0#: mResidualMx = 0#: mResidualMy = 0#
    mRelativeResidualN = 0#: mRelativeResidualMx = 0#: mRelativeResidualMy = 0#
    mLastResidualNorm = 0#
    mExtensionUsed = False: mWithinPhysicalRange = False: mConverged = False
    mStopReason = vbNullString: mDiagnosticLog = vbNullString
    mSolverCallCount = 0: mIterationCount = 0
    mMinConcreteStrain = 0#: mMaxConcreteStrain = 0#
    mMinSteelStrain = 0#: mMaxSteelStrain = 0#
    mMinConcreteStress = 0#: mMaxConcreteStress = 0#
    mMinSteelStress = 0#: mMaxSteelStress = 0#
    mHasConcreteCompression = False: mHasConcreteTension = False
    mHasSteelCompression = False: mHasSteelTension = False
End Sub

' По сохранённому полю деформаций и диаграммам готовит данные каждого элемента:
' имя, материал, деформацию, напряжение и состояние работы. Равновесие не решается
' повторно. Используется прежняя сетка; изменённая геометрия вызывает ошибку.
Public Function PrepareElementSnapshot(ByVal section As CSectionModel) As Variant
    If Not mConverged Then Exit Function
    ValidateElementSnapshotSource section
    Dim data() As Variant
    ReDim data(1 To section.ConcreteCount + section.RebarCount, 1 To 5)
    FillElementSnapshot section, data, 1, 1
    PrepareElementSnapshot = data
End Function

' Заполняет пять полей в готовом массиве потребителя без промежуточной матрицы.
' Значения остаются во внутренних единицах; Excel и преобразования вывода здесь не используются.
Public Sub CopyElementSnapshotTo(ByVal section As CSectionModel, ByRef data() As Variant, ByVal firstRow As Long, ByVal firstColumn As Long)
    If Not mConverged Then Exit Sub
    ValidateElementSnapshotSource section
    If firstRow < LBound(data, 1) Or firstColumn < LBound(data, 2) Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Неверное начало массива снимка элементов."
    If UBound(data, 1) - firstRow + 1 < section.ConcreteCount + section.RebarCount Or UBound(data, 2) - firstColumn < 4 Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Недостаточный размер массива снимка элементов."
    FillElementSnapshot section, data, firstRow, firstColumn
End Sub

' Проверяет принадлежность готового НДС неизменённой модели и наличие фактических диаграмм.
Private Sub ValidateElementSnapshotSource(ByVal section As CSectionModel)
    If section Is Nothing Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Модель для снимка элементов не передана."
    If Not section Is mSnapshotSection Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Снимок НДС относится к другой модели сечения."
    If section.Revision <> mSnapshotRevision Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "После расчета НДС геометрия изменилась; выполните расчет для текущего сечения."
    If Not HasMaterialSnapshot Then Err.Raise vbObjectError + 3974, "CSectionStateResult", "Фактические диаграммы состояния недоступны для записи элементов."
End Sub

' Вычисляет прежние поля элементов тем же порядком операций в заданном окне массива.
' Общая реализация сохраняет точное совпадение старого пятиколоночного и прямого путей вывода.
Private Sub FillElementSnapshot(ByVal section As CSectionModel, ByRef data() As Variant, ByVal firstRow As Long, ByVal firstColumn As Long)
    Dim i As Long, row As Long, strain As Double, stress As Double
    row = firstRow
    For i = 1 To section.ConcreteCount
        strain = mEpsilon0 + mKappaX * section.ConcreteY(i) + mKappaY * section.ConcreteX(i)
        stress = mSnapshotConcrete.GetStress(strain)
        data(row, firstColumn) = section.ConcreteID(i): data(row, firstColumn + 1) = "Concrete"
        data(row, firstColumn + 2) = strain: data(row, firstColumn + 3) = stress: data(row, firstColumn + 4) = ElementPhysicalState("Concrete", strain, stress)
        row = row + 1
    Next i
    For i = 1 To section.RebarCount
        strain = mEpsilon0 + mKappaX * section.RebarY(i) + mKappaY * section.RebarX(i)
        stress = mSnapshotSteel.GetStress(strain)
        data(row, firstColumn) = section.RebarID(i): data(row, firstColumn + 1) = "Rebar"
        data(row, firstColumn + 2) = strain: data(row, firstColumn + 3) = stress: data(row, firstColumn + 4) = ElementPhysicalState("Rebar", strain, stress)
        row = row + 1
    Next i
End Sub

' Определяет работу элемента по его деформации и напряжению.
' Это состояние материала, а не итоговый статус сочетания нагрузок.
Private Function ElementPhysicalState(ByVal materialType As String, ByVal strain As Double, ByVal stress As Double) As String
    If Abs(stress) <= ELEMENT_ZERO_TOLERANCE Then
        If materialType = "Concrete" And strain > ELEMENT_ZERO_TOLERANCE Then
            ElementPhysicalState = "InactiveTensionConcrete"
        Else
            ElementPhysicalState = "NearZero"
        End If
    ElseIf stress < 0# Then
        ElementPhysicalState = "Compression"
    Else
        ElementPhysicalState = "Tension"
    End If
End Function

' Показывает, сохранены ли фактически использованные диаграммы материалов.
' Эти диаграммы защищены от последующих изменений.
Public Property Get HasMaterialSnapshot() As Boolean
    HasMaterialSnapshot = Not mSnapshotConcrete Is Nothing And Not mSnapshotSteel Is Nothing
End Property
' Неизменяемая фактическая диаграмма бетона.
Public Property Get ConcreteDiagramSnapshot() As CMaterialDiagram
    Set ConcreteDiagramSnapshot = mSnapshotConcrete
End Property
' Неизменяемая фактическая диаграмма арматуры.
Public Property Get SteelDiagramSnapshot() As CMaterialDiagram
    Set SteelDiagramSnapshot = mSnapshotSteel
End Property

' Копирует усилия и невязки из решателя, чтобы сохранить качество равновесия.
' Целевая нагрузка получается как внутреннее усилие минус невязка: именно
' такой знак невязок принят в CSectionSolver.
Private Sub CaptureEquilibriumData(ByVal solver As CSectionSolver)
    mNint = solver.Nint
    mMxint = solver.Mxint
    mMyint = solver.Myint
    mResidualN = solver.ResidualN
    mResidualMx = solver.ResidualMx
    mResidualMy = solver.ResidualMy
    mTargetN = mNint - mResidualN
    mTargetMx = mMxint - mResidualMx
    mTargetMy = mMyint - mResidualMy
    mRelativeResidualN = solver.RelativeResidualN
    mRelativeResidualMx = solver.RelativeResidualMx
    mRelativeResidualMy = solver.RelativeResidualMy
    mLastResidualNorm = solver.LastResidualNorm
End Sub

' Какое именованное состояние рассчитано.
Public Property Get StateType() As ESectionStateType
    StateType = mStateType
End Property

' Какое именованное состояние рассчитано.
Public Property Get StateTypeText() As String
    StateTypeText = SectionStateTypeToText(mStateType)
End Property

' Возвращает назначение модели материалов: прочность, образование трещины или НДС с трещинами.
Public Property Get MaterialModelRole() As ECalculationPurpose
    MaterialModelRole = mMaterialModelRole
End Property

' Возвращает текстовое название назначения модели материалов.
Public Property Get MaterialModelRoleText() As String
    MaterialModelRoleText = MaterialModelRoleToText(mMaterialModelRole)
End Property

' Выдает независимую небольшую спецификацию, не копируя расчетную сетку.
' Внешний потребитель не может изменить модель уже сохраненного состояния.
Public Property Get MaterialSpec() As CMaterialModelSpec
    If Not mMaterialSpec Is Nothing Then Set MaterialSpec = mMaterialSpec.Clone
End Property

' Деформация в начале координат расчетной модели.
Public Property Get Epsilon0() As Double
    Epsilon0 = mEpsilon0
End Property

' Кривизна kappaX, 1/мм.
Public Property Get KappaX() As Double
    KappaX = mKappaX
End Property

' Кривизна kappaY, 1/мм.
Public Property Get KappaY() As Double
    KappaY = mKappaY
End Property

' Показывает, использовано ли в конечном НДС продолжение диаграмм за физическими пределами.
Public Property Get ExtensionUsed() As Boolean
    ExtensionUsed = mExtensionUsed
End Property

' Показывает, находится ли конечное НДС в физических пределах диаграмм материалов.
Public Property Get WithinPhysicalRange() As Boolean
    WithinPhysicalRange = mWithinPhysicalRange
End Property

' Показывает, найдено ли равновесие для сохранённого поля деформаций.
Public Property Get Converged() As Boolean
    Converged = mConverged
End Property

' Возвращает сохранённую причину остановки решения равновесия.
Public Property Get StopReason() As String
    StopReason = mStopReason
End Property

' Возвращает обозначение статуса НДС для пользователя по общим правилам проекта.
' Отдельная независимо изменяемая строка статуса не хранится.
Public Property Get Status() As String
    Status = ExternalStatus(ResultMeta)
End Property

' Возвращает копию статуса и пояснения расчётного состояния.
' Если состояние ещё не заполнено, сообщает о внутренней ошибке вместо успешного результата.
Public Property Get ResultMeta() As CResultMeta
    If mResultMeta Is Nothing Then
        Dim missing As CResultMeta
        Set missing = New CResultMeta
        missing.SetResult rsInternalError, rcInternalError, ResultKindFromStateType(mStateType), _
            "Расчетное состояние не заполнено.", vbNullString, True, False
        Set ResultMeta = missing
    Else
        Set ResultMeta = mResultMeta.Clone
    End If
End Property

' Внутренний статус результата; не является пользовательским подтверждением нормативной
' пригодности.
Public Property Get InternalStatus() As EResultInternalStatus
    InternalStatus = ResultMeta.InternalStatus
End Property

' Возвращает машинный код причины, сохранённый расчётом; внешнее обозначение статуса формируется отдельно.
Public Property Get ResultCode() As EResultCode
    ResultCode = ResultMeta.ResultCode
End Property

' Пояснение результата с фактической причиной и доступной привязкой к вводным данным.
Public Property Get ResultComment() As String
    ResultComment = ResultMeta.ResultComment
End Property

' Заданная продольная сила сохранённого НДС, Н.
Public Property Get TargetN() As Double
    TargetN = mTargetN
End Property

' Заданный момент Mx сохранённого НДС, Н*мм.
Public Property Get TargetMx() As Double
    TargetMx = mTargetMx
End Property

' Заданный момент My сохранённого НДС, Н*мм.
Public Property Get TargetMy() As Double
    TargetMy = mTargetMy
End Property

' Внутреннее N, полученное интегрированием напряжений, Н.
Public Property Get Nint() As Double
    Nint = mNint
End Property

' Внутренний Mx, полученный интегрированием напряжений, Н*мм.
Public Property Get Mxint() As Double
    Mxint = mMxint
End Property

' Внутренний My, полученный интегрированием напряжений, Н*мм.
Public Property Get Myint() As Double
    Myint = mMyint
End Property

' Невязка Nint - Ntarget, Н.
Public Property Get ResidualN() As Double
    ResidualN = mResidualN
End Property

' Невязка Mxint - Mxtarget, Н*мм.
Public Property Get ResidualMx() As Double
    ResidualMx = mResidualMx
End Property

' Невязка Myint - Mytarget, Н*мм.
Public Property Get ResidualMy() As Double
    ResidualMy = mResidualMy
End Property

' Относительная невязка по N для диагностики качества состояния.
Public Property Get RelativeResidualN() As Double
    RelativeResidualN = mRelativeResidualN
End Property

' Относительная невязка по Mx для диагностики качества состояния.
Public Property Get RelativeResidualMx() As Double
    RelativeResidualMx = mRelativeResidualMx
End Property

' Относительная невязка по My для диагностики качества состояния.
Public Property Get RelativeResidualMy() As Double
    RelativeResidualMy = mRelativeResidualMy
End Property

' Возвращает общую меру невязок равновесия последней попытки решения.
Public Property Get LastResidualNorm() As Double
    LastResidualNorm = mLastResidualNorm
End Property

' Возвращает число вызовов CSectionSolver.Solve для получения этого НДС.
Public Property Get SolverCallCount() As Long
    SolverCallCount = mSolverCallCount
End Property

' Возвращает число итераций последней попытки, из которой получено НДС.
Public Property Get IterationCount() As Long
    IterationCount = mIterationCount
End Property

' Возвращает сохранённый журнал поиска равновесия для последующего вывода.
Public Property Get DiagnosticLog() As String
    DiagnosticLog = mDiagnosticLog
End Property

' Возвращает минимальную деформацию бетона в сохранённом НДС.
Public Property Get MinConcreteStrain() As Double
    MinConcreteStrain = mMinConcreteStrain
End Property

' Возвращает максимальную деформацию бетона в сохранённом НДС.
Public Property Get MaxConcreteStrain() As Double
    MaxConcreteStrain = mMaxConcreteStrain
End Property

' Возвращает минимальную деформацию арматуры в сохранённом НДС.
Public Property Get MinSteelStrain() As Double
    MinSteelStrain = mMinSteelStrain
End Property

' Возвращает максимальную деформацию арматуры в сохранённом НДС.
Public Property Get MaxSteelStrain() As Double
    MaxSteelStrain = mMaxSteelStrain
End Property

' Наиболее сжимающее напряжение бетона со знаком.
Public Property Get MinConcreteStress() As Double
    MinConcreteStress = mMinConcreteStress
End Property

' Наиболее растягивающее напряжение бетона со знаком.
Public Property Get MaxConcreteStress() As Double
    MaxConcreteStress = mMaxConcreteStress
End Property

' Наиболее сжимающее напряжение арматуры со знаком.
Public Property Get MinSteelStress() As Double
    MinSteelStress = mMinSteelStress
End Property

' Наиболее растягивающее напряжение арматуры со знаком.
Public Property Get MaxSteelStress() As Double
    MaxSteelStress = mMaxSteelStress
End Property

' True, если в состоянии реально есть сжатый бетон.
Public Property Get HasConcreteCompression() As Boolean
    HasConcreteCompression = mHasConcreteCompression
End Property

' True, если бетон реально работает на растяжение.
Public Property Get HasConcreteTension() As Boolean
    HasConcreteTension = mHasConcreteTension
End Property

' True, если есть сжатая арматура.
Public Property Get HasSteelCompression() As Boolean
    HasSteelCompression = mHasSteelCompression
End Property

' True, если есть растянутая арматура.
Public Property Get HasSteelTension() As Boolean
    HasSteelTension = mHasSteelTension
End Property

' Создаёт статус сохранённого НДС по результату CSectionSolver.
' Использует признаки решения и физическую допустимость диаграмм,
' не определяя успех по тексту причины остановки.
Private Function BuildMetaFromSolver(ByVal stateType As ESectionStateType, ByVal solver As CSectionSolver, _
        ByVal extensionUsed As Boolean, ByVal withinPhysicalRange As Boolean, _
        ByVal stopReason As String) As CResultMeta
    If solver Is Nothing Then
        Dim missing As CResultMeta
        Set missing = New CResultMeta
        missing.SetResult rsInternalError, rcSolverDidNotReturnState, ResultKindFromStateType(stateType), _
            "Решатель не вернул обязательное расчетное состояние.", vbNullString, True, False
        Set BuildMetaFromSolver = missing
    ElseIf Not solver.Converged Then
        Set BuildMetaFromSolver = BuildFailureMeta(stateType, solver, stopReason)
    ElseIf extensionUsed Or Not withinPhysicalRange Then
        Set BuildMetaFromSolver = BuildPhysicalRangeMeta(stateType, extensionUsed, stopReason)
    Else
        Set BuildMetaFromSolver = BuildSuccessMeta(stateType, stopReason)
    End If
End Function

' Создаёт успешный статус для полученного физически допустимого поля деформаций.
Private Function BuildSuccessMeta(ByVal stateType As ESectionStateType, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetResult rsSuccess, rcCheckPassed, ResultKindFromStateType(stateType), resultComment
    Set BuildSuccessMeta = meta
End Function

' Сохраняет фактическую причину неудачи решателя. Ошибка ввода до итераций не означает
' выполненный расчёт; вырожденная система и исчерпание числа итераций различаются.
Private Function BuildFailureMeta(ByVal stateType As ESectionStateType, ByVal solver As CSectionSolver, _
        ByVal resultComment As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    meta.SetSolverFailure solver.FailureCode, ResultKindFromStateType(stateType), _
        resultComment, (solver.Iterations > 0), solver.FailureSettingKeys
    meta.DiagnosticDetails = solver.DiagnosticLog
    Set BuildFailureMeta = meta
End Function

' Сохраняет статус состояния, для которого равновесие найдено за физическими пределами
' диаграмм. Это превышение физического предела FAIL, а не численная несходимость.
Private Function BuildPhysicalRangeMeta(ByVal stateType As ESectionStateType, _
        ByVal extensionUsed As Boolean, ByVal diagnosticDetails As String) As CResultMeta
    Dim meta As CResultMeta
    Set meta = New CResultMeta
    Dim comment As String
    If extensionUsed Then
        comment = "Равновесие найдено с расширением диаграмм, но деформации материала вышли за физические пределы."
    Else
        comment = "Равновесие найдено, но деформации материала вышли за физические пределы диаграмм."
    End If
    comment = comment & " Проверка по физической диаграмме не проходит."
    meta.SetResult rsCheckFailed, rcStateOutsidePhysicalRange, ResultKindFromStateType(stateType), comment, diagnosticDetails
    Set BuildPhysicalRangeMeta = meta
End Function

' Возвращает внешний статус для плоских полей Results.
Private Function ExternalStatus(ByVal meta As CResultMeta) As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    ExternalStatus = policy.ExternalStatus(meta)
End Function

' Определяет вид инженерной проверки по типу сохранённого НДС.
Private Function ResultKindFromStateType(ByVal stateType As ESectionStateType) As EResultKind
    Select Case stateType
        Case sstCapacityState
            ResultKindFromStateType = rkCapacity
        Case sstPreCrackState, sstPostCrackState
            ResultKindFromStateType = rkCrackFormation
        Case sstCrackedState
            ResultKindFromStateType = rkDirectState
        Case Else
            ResultKindFromStateType = rkDirectState
    End Select
End Function

' Завершает заполнение НДС. Повторный вызов безопасен; поле деформаций, усилия,
' признаки решения, статус и пояснение после этого нельзя изменить.
Friend Sub Freeze()
    If mFrozen Then Exit Sub
    mFrozen = True
End Sub

' Запрещает изменение результата после передачи потребителям.
' Для повторного расчёта создаётся новый объект, а прежний результат сохраняется.
Private Sub AssertWritable()
    If mFrozen Then Err.Raise vbObjectError + 4220, "CSectionStateResult", _
        "Опубликованный результат нельзя изменить; для нового расчета нужен новый объект."
End Sub


' Возвращает геометрический смысл сохранённого поля деформаций без нового решения НДС.
' Не меняет даже frozen-состояние; при нулевой кривизне прежний h отсутствует, x определён и равен нулю.
' При Nothing сечения обе величины не определены; writer получает готовые флаги записи, а не выбирает критерии.
Public Sub GetGeometryDepths(ByVal section As CSectionModel, ByRef depthH As Double, _
        ByRef compressedDepth As Double, ByRef hasH As Boolean, ByRef hasX As Boolean)
    depthH = 0#: compressedDepth = 0#: hasH = False: hasX = False
    If section Is Nothing Then Exit Sub
    CalculateDepthsPerpendicularToNeutral section, mEpsilon0, mKappaX, mKappaY, depthH, compressedDepth
    hasH = (depthH > 0#): hasX = (compressedDepth >= 0#)
End Sub

' Определяет габарит h и глубину сжатой зоны x вдоль нормали к нейтральной линии.
Private Sub CalculateDepthsPerpendicularToNeutral(ByVal section As CSectionModel, ByVal eps0 As Double, _
        ByVal kappaX As Double, ByVal kappaY As Double, ByRef depthH As Double, ByRef compressedDepth As Double)
    Dim normValue As Double
    normValue = Sqr(kappaX * kappaX + kappaY * kappaY)
    If normValue <= 0.000000000000001 Then Exit Sub

    Dim nx As Double
    Dim ny As Double
    nx = kappaY / normValue ' Нормаль поля epsilon = eps0 + kappaX*y + kappaY*x направлена к росту деформации.
    ny = kappaX / normValue

    Dim minS As Double
    Dim maxS As Double
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateProjection section, nx, ny, False, minS, maxS

    depthH = maxS - minS
    If depthH <= 0# Then Exit Sub

    Dim neutralS As Double
    neutralS = -eps0 / normValue ' Координата нулевой деформации по единичной нормали, мм.

    If eps0 + normValue * minS < 0# And eps0 + normValue * maxS < 0# Then
        compressedDepth = depthH
    ElseIf eps0 + normValue * minS >= 0# And eps0 + normValue * maxS >= 0# Then
        compressedDepth = 0#
    Else
        compressedDepth = MaxDouble(0#, MinDouble(maxS, neutralS) - minS)
    End If
End Sub

' Возвращает максимум двух чисел для геометрических величин таблицы.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxDouble = a Else MaxDouble = b
End Function

' Возвращает минимум двух чисел для геометрических величин таблицы.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a <= b Then MinDouble = a Else MinDouble = b
End Function
