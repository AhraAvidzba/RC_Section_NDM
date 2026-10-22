VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialModelProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подготовка расчётных диаграмм бетона и арматуры
' ==========================================================================
' По параметрам CSystemSettingsReader строит CMaterialDiagram для прочности,
' образования трещины и НДС с трещинами. Здесь выбираются I/II ГПС,
' сопротивления R/Rser, TwoLine/ThreeLine и учёт растянутого бетона.
' При разрешённом численном продолжении добавляет вспомогательную ветвь
' за пределом диаграммы для решения равновесия. Исходные физические пределы
' деформаций и критерии инженерных проверок сохраняются.

Private Const SLOT_STRENGTH As Long = 1 ' Массивный слот физической диаграммы I ГПС.
Private Const SLOT_MCRC As Long = 2 ' Массивный слот физической диаграммы образования трещины.
Private Const SLOT_CRACKED_NDS As Long = 3 ' Массивный слот физической диаграммы раскрытой трещины.
Private Const STATE_EXTENSION_E_FACTOR As Double = 0.01 ' Жесткость extension: 1% от начального модуля E.
Private Const EXTENSION_TOLERANCE As Double = 0.000000000001 ' Допуск, чтобы не добавлять почти совпадающую extension-точку.

Private mConcreteParameters As CConcreteMaterialParameters ' Исходные параметры бетона из Config.
Private mSteelParameters As CSteelMaterialParameters ' Исходные параметры арматуры из Config.
Private mConcreteDiagram(1 To 3) As CMaterialDiagram ' Диаграммы бетона в физических пределах, без вспомогательного численного продолжения.
Private mSteelDiagram(1 To 3) As CMaterialDiagram ' Диаграммы арматуры в физических пределах, без вспомогательного численного продолжения.
Private mPurposeName(1 To 3) As String ' Принятое имя назначения материалов для записи в Results и диагностику.
Private mLimitStateGroup(1 To 3) As String ' I или II ГПС, фактически использованная provider-ом.
Private mConcreteDiagramType(1 To 3) As String ' TwoLine или ThreeLine для бетона.
Private mConcreteTensionBehavior(1 To 3) As String ' Ignore или UseDiagram для растянутого бетона.
Private mSteelDiagramType(1 To 3) As String ' TwoLine или ThreeLine для арматуры.
Private mConcreteCompressionLimit(1 To 3) As Double ' Физический предел сжатой деформации бетона.
Private mConcreteTensionLimit(1 To 3) As Double ' Физический предел растянутой деформации бетона, если ветвь включена.
Private mConcreteTensionLimitEnabled(1 To 3) As Boolean ' True, если растянутый бетон имеет физический eps-предел.
Private mSteelCompressionLimit(1 To 3) As Double ' Физический предел сжатой деформации арматуры.
Private mSteelTensionLimit(1 To 3) As Double ' Физический предел растянутой деформации арматуры.
Private mConcreteCompressionResistance(1 To 3) As Double ' Rb/Rb,ser выбранной ГПС, МПа.
Private mConcreteTensionResistance(1 To 3) As Double ' Rbt/Rbt,ser выбранной ГПС, МПа.
Private mSteelCompressionResistance(1 To 3) As Double ' Rsc/Rsc,ser выбранной ГПС, МПа.
Private mSteelTensionResistance(1 To 3) As Double ' Rs/Rs,ser выбранной ГПС, МПа.
Private mConcreteElasticModulus(1 To 3) As Double ' Eb выбранной бетонной модели, МПа.
Private mSteelElasticModulus(1 To 3) As Double ' Es выбранной модели арматуры, МПа.
Private mDiagramExtensionEnabled As Boolean ' Разрешение вспомогательного продолжения диаграмм при поиске равновесия.
Private mPlateauSlopeFactor As Double ' Et = alpha * E; default Config 0.001 дает наклон плато 0.001 * E (0.1% E); ноль — горизонтальное плато.
Private mInputLocations As Object ' Независимый снимок строк фактических адресов Concrete/Steel; не хранит изменяемый reader или объекты Excel.
Private mInitialized As Boolean ' Защита от использования provider-а до чтения параметров.
Private mDiagramCache As Object ' Кэш CMaterialDiagram по MaterialModelSpec; диаграммы после Initialize неизменяемы.
Private mRevision As Long ' Версия набора материалов и effective Extension после переинициализации.

' Читает Config, создает параметры материалов и строит физические диаграммы.
Public Sub Initialize(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    Set mInputLocations = Nothing
    If settings Is Nothing Then Err.Raise vbObjectError + 3240, "CMaterialModelProvider", "Настройки материалов не переданы."
    SnapshotMaterialInputLocations settings
    mInitialized = False
    mRevision = mRevision + 1
    Set mDiagramCache = CreateObject("Scripting.Dictionary")

    Set mConcreteParameters = New CConcreteMaterialParameters
    mConcreteParameters.LoadFromSettings settings, units

    Set mSteelParameters = New CSteelMaterialParameters
    mSteelParameters.LoadFromSettings settings, units

    mDiagramExtensionEnabled = ReadDiagramExtensionSetting(settings)
    mPlateauSlopeFactor = settings.GetRequiredDouble("General.PlateauSlopeFactor")
    If mPlateauSlopeFactor < 0# Then Err.Raise vbObjectError + 3247, "CMaterialModelProvider", _
        settings.InputErrorMessage("General.PlateauSlopeFactor", "Коэффициент наклона плато отрицателен.", _
            "Введите число не меньше нуля: 0 — горизонтальное плато; положительное число — коэффициент к E.")

    BuildPurpose cpStrength, "I", _
        settings.GetString("Diagram.Strength.Concrete", "TwoLine"), _
        settings.GetString("Diagram.Strength.ConcreteTension", "Ignore"), _
        settings.GetString("Diagram.Strength.Steel", "TwoLine")

    BuildPurpose cpMcrc, "II", _
        settings.GetString("Diagram.Mcrc.Concrete", "ThreeLine"), _
        "UseDiagram", _
        settings.GetString("Diagram.Mcrc.Steel", "TwoLine")

    BuildPurpose cpCrackedNDS, "II", _
        settings.GetString("Diagram.CrackedNDS.Concrete", "TwoLine"), _
        "Ignore", _
        settings.GetString("Diagram.CrackedNDS.Steel", "TwoLine")

    mInitialized = True
End Sub

' Сохраняет только неизменяемые строки мест ввода материалов один раз.
' FromSpec может построить другую активную диаграмму после Initialize; ее
' ошибка также получит фактический адрес без чтения листов и хранения reader.
Private Sub SnapshotMaterialInputLocations(ByVal settings As CSystemSettingsReader)
    Set mInputLocations = CreateObject("Scripting.Dictionary")
    mInputLocations.CompareMode = vbTextCompare
    Dim i As Long, key As String
    For i = 1 To settings.KeyCount
        key = settings.KeyAt(i)
        If StrComp(Left$(key, 9), "Concrete.", vbTextCompare) = 0 Or _
                StrComp(Left$(key, 6), "Steel.", vbTextCompare) = 0 Then
            If Not mInputLocations.Exists(key) Then mInputLocations.Add key, settings.InputLocation(key)
        End If
    Next i
End Sub

' Читает обязательную текущую настройку численного продолжения диаграмм.
Private Function ReadDiagramExtensionSetting(ByVal settings As CSystemSettingsReader) As Boolean
    ReadDiagramExtensionSetting = settings.GetRequiredBoolean("General.DiagramExtension")
End Function

' Строит полный физический материальный набор одной расчетной цели.
Private Sub BuildPurpose(ByVal purpose As ECalculationPurpose, ByVal limitStateGroup As String, _
        ByVal concreteDiagramType As String, ByVal concreteTensionBehavior As String, _
        ByVal steelDiagramType As String)
    Dim slot As Long
    slot = PhysicalPurposeSlot(purpose)

    mPurposeName(slot) = PurposeToText(purpose)
    mLimitStateGroup(slot) = NormalizeLimitStateGroup(limitStateGroup)
    mConcreteDiagramType(slot) = NormalizeDiagramType(concreteDiagramType)
    mConcreteTensionBehavior(slot) = NormalizeTensionBehavior(concreteTensionBehavior)
    mSteelDiagramType(slot) = NormalizeDiagramType(steelDiagramType)

    ConcreteStrengths mLimitStateGroup(slot), mConcreteCompressionResistance(slot), mConcreteTensionResistance(slot)
    SteelStrengths mLimitStateGroup(slot), mSteelCompressionResistance(slot), mSteelTensionResistance(slot)
    mConcreteElasticModulus(slot) = mConcreteParameters.Eb
    mSteelElasticModulus(slot) = mSteelParameters.Es

    Set mConcreteDiagram(slot) = BuildConcreteDiagram(mLimitStateGroup(slot), _
        mConcreteDiagramType(slot), mConcreteTensionBehavior(slot))
    Set mSteelDiagram(slot) = BuildSteelDiagram(mLimitStateGroup(slot), mSteelDiagramType(slot))

    mConcreteCompressionLimit(slot) = mConcreteDiagram(slot).UltimateCompressionStrain
    mConcreteTensionLimit(slot) = mConcreteDiagram(slot).UltimateTensionStrain
    mConcreteTensionLimitEnabled(slot) = (mConcreteDiagram(slot).HasPhysicalTensionLimit And _
        StrComp(mConcreteTensionBehavior(slot), "UseDiagram", vbTextCompare) = 0)
    mSteelCompressionLimit(slot) = mSteelDiagram(slot).UltimateCompressionStrain
    mSteelTensionLimit(slot) = mSteelDiagram(slot).UltimateTensionStrain
End Sub

' Выбирает формулу бетонной диаграммы.
Private Function BuildConcreteDiagram(ByVal limitStateGroup As String, ByVal diagramType As String, _
        ByVal tensionBehavior As String) As CMaterialDiagram
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        Set BuildConcreteDiagram = BuildConcreteThreeLine(limitStateGroup, tensionBehavior)
    Else
        Set BuildConcreteDiagram = BuildConcreteTwoLine(limitStateGroup, tensionBehavior)
    End If
End Function

' Строит двухлинейную бетонную диаграмму. Параметры eps задает Config,
' provider добавляет знак, сопротивление выбранной ГПС и опциональную
' растянутую ветвь.
Private Function BuildConcreteTwoLine(ByVal limitStateGroup As String, ByVal tensionBehavior As String) As CMaterialDiagram
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths limitStateGroup, rb, rbt

    Dim count As Long
    If StrComp(tensionBehavior, "UseDiagram", vbTextCompare) = 0 Then
        count = 5
    Else
        count = 3
    End If

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To count)
    ReDim sig(1 To count)

    eps(1) = -mConcreteParameters.Eb2: sig(1) = -rb
    eps(2) = -mConcreteParameters.Eb1Red: sig(2) = -rb
    eps(3) = 0#: sig(3) = 0#
    If count = 5 Then
        eps(4) = mConcreteParameters.Ebt1Red: sig(4) = rbt
        eps(5) = mConcreteParameters.Ebt2: sig(5) = rbt
    End If

    ValidateStrainPointOrder eps, count, Array("Concrete.Eb2", "Concrete.Eb1Red", vbNullString, _
        "Concrete.Ebt1Red", "Concrete.Ebt2"), "бетона TwoLine"
    Set BuildConcreteTwoLine = DiagramFromArrays(eps, sig, count, mConcreteParameters.Eb, mConcreteParameters.Ebt)
End Function

' Строит трехлинейную бетонную диаграмму. Средняя точка 0.6*R/E
' вычисляется здесь, чтобы расчетное ядро не знало формулы построения.
Private Function BuildConcreteThreeLine(ByVal limitStateGroup As String, ByVal tensionBehavior As String) As CMaterialDiagram
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths limitStateGroup, rb, rbt

    Dim count As Long
    If StrComp(tensionBehavior, "UseDiagram", vbTextCompare) = 0 Then
        count = 7
    Else
        count = 4
    End If

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To count)
    ReDim sig(1 To count)

    eps(1) = -mConcreteParameters.Eb2: sig(1) = -rb
    eps(2) = -mConcreteParameters.Eb0: sig(2) = -rb
    eps(3) = -0.6 * rb / mConcreteParameters.Eb: sig(3) = -0.6 * rb
    eps(4) = 0#: sig(4) = 0#
    If count = 7 Then
        eps(5) = 0.6 * rbt / mConcreteParameters.Ebt: sig(5) = 0.6 * rbt
        eps(6) = mConcreteParameters.Ebt0: sig(6) = rbt
        eps(7) = mConcreteParameters.Ebt2: sig(7) = rbt
    End If

    Dim suffix As String
    suffix = MaterialStrengthKeySuffix(limitStateGroup)
    ValidateStrainPointOrder eps, count, Array("Concrete.Eb2", "Concrete.Eb0", _
        "Concrete.Rb." & suffix & ";Concrete.Eb", vbNullString, _
        "Concrete.Rbt." & suffix & ";Concrete.Ebt", "Concrete.Ebt0", "Concrete.Ebt2"), "бетона ThreeLine"
    Set BuildConcreteThreeLine = DiagramFromArrays(eps, sig, count, mConcreteParameters.Eb, mConcreteParameters.Ebt)
End Function

' Выбирает формулу диаграммы арматуры.
Private Function BuildSteelDiagram(ByVal limitStateGroup As String, ByVal diagramType As String) As CMaterialDiagram
    If StrComp(diagramType, "ThreeLine", vbTextCompare) = 0 Then
        Set BuildSteelDiagram = BuildSteelThreeLine(limitStateGroup)
    Else
        Set BuildSteelDiagram = BuildSteelTwoLine(limitStateGroup)
    End If
End Function

' Строит двухлинейную диаграмму арматуры.
Private Function BuildSteelTwoLine(ByVal limitStateGroup As String) As CMaterialDiagram
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths limitStateGroup, rsc, rs

    Dim eps(1 To 5) As Double
    Dim sig(1 To 5) As Double
    eps(1) = -mSteelParameters.TwoLineEsc2: sig(1) = -rsc
    eps(2) = -rsc / mSteelParameters.Esc: sig(2) = -rsc
    eps(3) = 0#: sig(3) = 0#
    eps(4) = rs / mSteelParameters.Es: sig(4) = rs
    eps(5) = mSteelParameters.TwoLineEs2: sig(5) = rs

    Dim suffix As String
    suffix = MaterialStrengthKeySuffix(limitStateGroup)
    ValidateStrainPointOrder eps, 5, Array("Steel.TwoLine.Esc2", "Steel.Rsc." & suffix & ";Steel.Esc", _
        vbNullString, "Steel.Rs." & suffix & ";Steel.Es", "Steel.TwoLine.Es2"), "арматуры TwoLine"
    Set BuildSteelTwoLine = DiagramFromArrays(eps, sig, 5, mSteelParameters.Esc, mSteelParameters.Es)
End Function

' Строит трехлинейную диаграмму арматуры с расчетными точками перелома.
Private Function BuildSteelThreeLine(ByVal limitStateGroup As String) As CMaterialDiagram
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths limitStateGroup, rsc, rs

    Dim epsSc0 As Double
    Dim epsSc1 As Double
    Dim epsScPl As Double
    Dim epsS0 As Double
    Dim epsS1 As Double
    Dim epsSPl As Double

    epsSc0 = rsc / mSteelParameters.Esc + 0.002
    epsSc1 = 0.9 * rsc / mSteelParameters.Esc
    epsScPl = 2# * epsSc0 - epsSc1 ' Абсцисса достижения 1,1 Rsc при продолжении наклона между точками 0,9 Rsc и Rsc.
    epsS0 = rs / mSteelParameters.Es + 0.002
    epsS1 = 0.9 * rs / mSteelParameters.Es
    epsSPl = 2# * epsS0 - epsS1 ' Аналогичная точка 1,1 Rs для растянутой ветви; характеристики сжатия и растяжения независимы.

    Dim eps(1 To 9) As Double
    Dim sig(1 To 9) As Double
    eps(1) = -mSteelParameters.ThreeLineEsc2: sig(1) = -1.1 * rsc
    eps(2) = -epsScPl: sig(2) = -1.1 * rsc
    eps(3) = -epsSc0: sig(3) = -rsc
    eps(4) = -epsSc1: sig(4) = -0.9 * rsc
    eps(5) = 0#: sig(5) = 0#
    eps(6) = epsS1: sig(6) = 0.9 * rs
    eps(7) = epsS0: sig(7) = rs
    eps(8) = epsSPl: sig(8) = 1.1 * rs
    eps(9) = mSteelParameters.ThreeLineEs2: sig(9) = 1.1 * rs

    Dim suffix As String, compressionKeys As String, tensionKeys As String
    suffix = MaterialStrengthKeySuffix(limitStateGroup)
    compressionKeys = "Steel.Rsc." & suffix & ";Steel.Esc"
    tensionKeys = "Steel.Rs." & suffix & ";Steel.Es"
    ValidateStrainPointOrder eps, 9, Array("Steel.ThreeLine.Esc2", compressionKeys, compressionKeys, _
        compressionKeys, vbNullString, tensionKeys, tensionKeys, tensionKeys, "Steel.ThreeLine.Es2"), "арматуры ThreeLine"
    Set BuildSteelThreeLine = DiagramFromArrays(eps, sig, 9, mSteelParameters.Esc, mSteelParameters.Es)
End Function

' Выбирает только ключ набора сопротивлений для сообщения; формулы точек и
' действующие сопротивления остаются в ConcreteStrengths/SteelStrengths.
Private Function MaterialStrengthKeySuffix(ByVal limitStateGroup As String) As String
    If StrComp(limitStateGroup, "II", vbTextCompare) = 0 Then MaterialStrengthKeySuffix = "SLS" Else MaterialStrengthKeySuffix = "ULS"
End Function

' Проверяет порядок фактически используемых точек до изменения напряжений плато.
' Сортировка готовой CMaterialDiagram не должна превращать ошибочный eps_ult в
' другую точку и незаметно менять физический предел. Неактивные ветви не проверяются.
Private Sub ValidateStrainPointOrder(ByRef strains() As Double, ByVal count As Long, _
        ByVal sourceKeys As Variant, ByVal materialLabel As String)
    Dim i As Long
    For i = 2 To count
        If strains(i) <= strains(i - 1) Then
            Dim message As String, correction As String, keys As String, key As Variant, seen As String
            correction = "Проверьте указанные параметры: модуль предельной деформации должен быть больше деформации соседнего перелома, а промежуточные точки должны следовать по порядку."
            message = "Противоречивые параметры диаграммы " & materialLabel & ": соседние точки имеют деформации " & _
                CStr(strains(i - 1)) & " и " & CStr(strains(i)) & ", которые не возрастают. " & correction
            keys = CStr(sourceKeys(i - 2)) & ";" & CStr(sourceKeys(i - 1))
            For Each key In Split(keys, ";")
                If Len(CStr(key)) > 0 And InStr(1, seen, ";" & CStr(key) & ";", vbBinaryCompare) = 0 Then
                    seen = seen & ";" & CStr(key) & ";"
                    message = message & vbCrLf & "Настройка " & CStr(key)
                    If Not mInputLocations Is Nothing Then
                        If mInputLocations.Exists(CStr(key)) Then message = message & " (" & CStr(mInputLocations(CStr(key))) & ")"
                    End If
                    message = message & "."
                End If
            Next key
            Err.Raise vbObjectError + 3260, "CMaterialModelProvider", message
        End If
    Next i
End Sub

' Создает CMaterialDiagram из массивов физических точек.
Private Function DiagramFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long, _
        ByVal compressionE As Double, ByVal tensionE As Double) As CMaterialDiagram
    Dim diagram As CMaterialDiagram
    Set diagram = New CMaterialDiagram
    ApplyPlateauSlope strains, stresses, pointCount, compressionE, tensionE
    diagram.InitializeFromArrays strains, stresses, pointCount
    diagram.Freeze
    Set DiagramFromArrays = diagram
End Function

' Изменяем точки последовательно от начала координат, сохраняя начало плато
' и непрерывность напряжений. Напряжение и касательный модуль используют те же точки.
' Отключённая нулевая ветвь, физические пределы и отдельное продолжение диаграммы не меняются.
Private Sub ApplyPlateauSlope(ByRef strains() As Double, ByRef stresses() As Double, ByVal count As Long, _
        ByVal compressionE As Double, ByVal tensionE As Double)
    Dim zero As Long, i As Long, previous As Double, original As Double, offset As Double ' previous — исходное напряжение соседнего узла; offset накапливает поправку наклона для непрерывности следующих участков.
    For i = 1 To count
        If strains(i) = 0# And stresses(i) = 0# Then zero = i: Exit For
    Next i
    If zero = 0 Then Err.Raise vbObjectError + 3246, "CMaterialModelProvider", "В диаграмме отсутствует начало координат."
    For i = zero - 1 To 1 Step -1
        original = stresses(i)
        If original <> 0# And original = previous Then _
            offset = offset + mPlateauSlopeFactor * compressionE * (strains(i) - strains(i + 1))
        stresses(i) = original + offset
        previous = original
    Next i
    previous = 0#: offset = 0#
    For i = zero + 1 To count
        original = stresses(i)
        If original <> 0# And original = previous Then _
            offset = offset + mPlateauSlopeFactor * tensionE * (strains(i) - strains(i - 1))
        stresses(i) = original + offset
        previous = original
    Next i
End Sub

' Создает копию физической диаграммы с дополнительными техническими точками
' для равновесного solve. Физические eps_ult передаются в CMaterialDiagram
' отдельно и поэтому не меняются.
Private Function DiagramWithExtension(ByVal baseDiagram As CMaterialDiagram, _
        ByVal compressionE As Double, ByVal tensionE As Double, _
        ByVal extendCompression As Boolean, ByVal extendTension As Boolean) As CMaterialDiagram
    On Error GoTo InvalidExtension
    Dim compressionStart As Double
    Dim tensionStart As Double
    compressionStart = baseDiagram.PhysicalCompressionStrain
    tensionStart = baseDiagram.PhysicalTensionStrain

    Dim extraCompression As Boolean
    Dim extraTension As Boolean
    extraCompression = extendCompression
    extraTension = extendTension And baseDiagram.HasPhysicalTensionLimit

    Dim eps() As Double
    Dim sig() As Double
    ReDim eps(1 To baseDiagram.PointCount + 2)
    ReDim sig(1 To baseDiagram.PointCount + 2)

    Dim i As Long
    Dim targetIndex As Long
    targetIndex = 1
    If extraCompression Then
        eps(targetIndex) = -ExtensionReferenceStrain(compressionStart)
        sig(targetIndex) = baseDiagram.GetStress(compressionStart) + _
            (eps(targetIndex) - compressionStart) * STATE_EXTENSION_E_FACTOR * compressionE
        targetIndex = targetIndex + 1
    End If

    For i = 1 To baseDiagram.PointCount
        eps(targetIndex) = baseDiagram.PointStrain(i)
        sig(targetIndex) = baseDiagram.PointStress(i)
        targetIndex = targetIndex + 1
    Next i

    If extraTension Then
        eps(targetIndex) = ExtensionReferenceStrain(tensionStart)
        sig(targetIndex) = baseDiagram.GetStress(tensionStart) + _
            (eps(targetIndex) - tensionStart) * STATE_EXTENSION_E_FACTOR * tensionE
        targetIndex = targetIndex + 1
    End If

    Dim count As Long
    count = targetIndex - 1

    Dim diagram As CMaterialDiagram
    Set diagram = New CMaterialDiagram
    diagram.InitializeFromArrays eps, sig, count, _
        baseDiagram.PhysicalCompressionStrain, baseDiagram.PhysicalTensionStrain, _
        extraCompression, extraTension, compressionStart, tensionStart
    diagram.Freeze
    Set DiagramWithExtension = diagram
    Exit Function

InvalidExtension:
    Err.Raise vbObjectError + 3245, "CMaterialModelProvider", _
        "Численное продолжение диаграммы невозможно для заданных деформаций и модулей: " & Err.Description
End Function

' Выбирает опорную точку, задающую наклон продолжения. Она находится за
' физическим пределом и не ограничивает деформацию: CMaterialDiagram
' продолжает эту прямую дальше. Непредставимые числа отклоняются построителем.
Private Function ExtensionReferenceStrain(ByVal physicalLimit As Double) As Double
    Dim referenceSpan As Double
    referenceSpan = Abs(physicalLimit)
    If referenceSpan < 1# Then referenceSpan = 1#
    ExtensionReferenceStrain = Abs(physicalLimit) + referenceSpan
End Function

' Возвращает бетонные сопротивления нужной ГПС.
Private Sub ConcreteStrengths(ByVal limitStateGroup As String, ByRef rb As Double, ByRef rbt As Double)
    If StrComp(limitStateGroup, "II", vbTextCompare) = 0 Then
        rb = mConcreteParameters.RbSLS
        rbt = mConcreteParameters.RbtSLS
    Else
        rb = mConcreteParameters.RbULS
        rbt = mConcreteParameters.RbtULS
    End If
End Sub

' Возвращает сопротивления арматуры нужной ГПС.
Private Sub SteelStrengths(ByVal limitStateGroup As String, ByRef rsc As Double, ByRef rs As Double)
    If StrComp(limitStateGroup, "II", vbTextCompare) = 0 Then
        rsc = mSteelParameters.RscSLS
        rs = mSteelParameters.RsSLS
    Else
        rsc = mSteelParameters.RscULS
        rs = mSteelParameters.RsULS
    End If
End Sub

' Возвращает индекс внутренних массивов для физической цели.
Private Function PhysicalPurposeSlot(ByVal purpose As ECalculationPurpose) As Long
    Select Case purpose
        Case cpStrength
            PhysicalPurposeSlot = SLOT_STRENGTH
        Case cpMcrc
            PhysicalPurposeSlot = SLOT_MCRC
        Case cpCrackedNDS
            PhysicalPurposeSlot = SLOT_CRACKED_NDS
        Case Else
            Err.Raise vbObjectError + 3241, "CMaterialModelProvider", _
                "Для поставщика материалов требуется физическая расчетная цель, а не " & PurposeToText(purpose) & "."
    End Select
End Function

' Проверяет, что StateSolution питается допустимой физической базой.
Private Sub ValidateStateBasePurpose(ByVal basePurpose As ECalculationPurpose)
    If Not IsStateBasePurpose(basePurpose) Then
        Err.Raise vbObjectError + 3246, "CMaterialModelProvider", _
            "StateSolution может использовать только Strength или CrackedNDS как физическую базу."
    End If
End Sub

' Нормализует группу предельных состояний.
Private Function NormalizeLimitStateGroup(ByVal limitStateGroup As String) As String
    Select Case LCase$(Trim$(limitStateGroup))
        Case "ii", "2", "sls", "group2"
            NormalizeLimitStateGroup = "II"
        Case "i", "1", "uls", "group1"
            NormalizeLimitStateGroup = "I"
        Case Else
            Err.Raise vbObjectError + 3245, "CMaterialModelProvider", "Группа предельных состояний для поставщика материалов должна быть I или II."
    End Select
End Function

' Нормализует тип диаграммы.
Private Function NormalizeDiagramType(ByVal diagramType As String) As String
    Select Case LCase$(Trim$(diagramType))
        Case "threeline", "three"
            NormalizeDiagramType = "ThreeLine"
        Case Else
            NormalizeDiagramType = "TwoLine"
    End Select
End Function

' Нормализует режим растянутого бетона.
Private Function NormalizeTensionBehavior(ByVal tensionBehavior As String) As String
    Select Case LCase$(Trim$(tensionBehavior))
        Case "usediagram", "use"
            NormalizeTensionBehavior = "UseDiagram"
        Case Else
            NormalizeTensionBehavior = "Ignore"
    End Select
End Function

' Возвращает физическую бетонную диаграмму без extension.
Public Function ConcreteMaterial(ByVal purpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    Set ConcreteMaterial = mConcreteDiagram(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую диаграмму арматуры без extension.
Public Function SteelMaterial(ByVal purpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    Set SteelMaterial = mSteelDiagram(PhysicalPurposeSlot(purpose))
End Function

' Возвращает бетонную диаграмму для решения равновесия на базе заданной физической цели.
Public Function ConcreteMaterialForEquilibrium(ByVal basePurpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    ValidateStateBasePurpose basePurpose
    If mDiagramExtensionEnabled Then
        Set ConcreteMaterialForEquilibrium = DiagramWithExtension(mConcreteDiagram(PhysicalPurposeSlot(basePurpose)), _
            mConcreteParameters.Eb, mConcreteParameters.Ebt, True, True)
    Else
        Set ConcreteMaterialForEquilibrium = mConcreteDiagram(PhysicalPurposeSlot(basePurpose))
    End If
End Function

' Возвращает диаграмму арматуры для решения равновесия на базе заданной физической цели.
Public Function SteelMaterialForEquilibrium(ByVal basePurpose As ECalculationPurpose) As CMaterialDiagram
    EnsureInitialized
    ValidateStateBasePurpose basePurpose
    If mDiagramExtensionEnabled Then
        Set SteelMaterialForEquilibrium = DiagramWithExtension(mSteelDiagram(PhysicalPurposeSlot(basePurpose)), _
            mSteelParameters.Esc, mSteelParameters.Es, True, True)
    Else
        Set SteelMaterialForEquilibrium = mSteelDiagram(PhysicalPurposeSlot(basePurpose))
    End If
End Function

' Возвращает физическую диаграмму бетона по параметрам профиля: набору
' характеристик, TwoLine/ThreeLine и учёту растяжения.
' Численный решатель получает уже готовую CMaterialDiagram.
Public Function ConcreteMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec

    Dim cacheKey As String
    cacheKey = MaterialDiagramCacheKey("ConcretePhysical", spec)
    Dim cached As CMaterialDiagram
    Set cached = CachedDiagram(cacheKey)
    If Not cached Is Nothing Then
        Set ConcreteMaterialFromSpec = cached
        Exit Function
    End If

    Dim diagram As CMaterialDiagram
    Set diagram = BuildConcreteDiagram(LimitStateGroupFromSpec(spec), _
        spec.ConcreteDiagramText, spec.ConcreteTensionText)
    StoreCachedDiagram cacheKey, diagram
    Set ConcreteMaterialFromSpec = diagram
End Function

' Возвращает физическую диаграмму арматуры по спецификации профиля.
Public Function SteelMaterialFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec

    Dim cacheKey As String
    cacheKey = MaterialDiagramCacheKey("SteelPhysical", spec)
    Dim cached As CMaterialDiagram
    Set cached = CachedDiagram(cacheKey)
    If Not cached Is Nothing Then
        Set SteelMaterialFromSpec = cached
        Exit Function
    End If

    Dim diagram As CMaterialDiagram
    Set diagram = BuildSteelDiagram(LimitStateGroupFromSpec(spec), spec.SteelDiagramText)
    StoreCachedDiagram cacheKey, diagram
    Set SteelMaterialFromSpec = diagram
End Function

' Возвращает бетонную диаграмму по спецификации профиля для задач равновесия.
' Если General.DiagramExtension включен, добавляется только техническое
' продолжение активных ветвей; physical limits самого материала не меняются.
Public Function ConcreteMaterialForEquilibriumFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec

    Dim cacheKey As String
    cacheKey = MaterialDiagramCacheKey("ConcreteState|" & CStr(mDiagramExtensionEnabled), spec)
    Dim cached As CMaterialDiagram
    Set cached = CachedDiagram(cacheKey)
    If Not cached Is Nothing Then
        Set ConcreteMaterialForEquilibriumFromSpec = cached
        Exit Function
    End If

    Dim physical As CMaterialDiagram
    Set physical = ConcreteMaterialFromSpec(spec)
    Dim diagram As CMaterialDiagram
    If mDiagramExtensionEnabled Then
        Set diagram = DiagramWithExtension(physical, _
            mConcreteParameters.Eb, mConcreteParameters.Ebt, True, True)
    Else
        Set diagram = physical
    End If
    StoreCachedDiagram cacheKey, diagram
    Set ConcreteMaterialForEquilibriumFromSpec = diagram
End Function

' Возвращает диаграмму арматуры по спецификации профиля для задач равновесия.
Public Function SteelMaterialForEquilibriumFromSpec(ByVal spec As CMaterialModelSpec) As CMaterialDiagram
    EnsureInitialized
    ValidateSpec spec

    Dim cacheKey As String
    cacheKey = MaterialDiagramCacheKey("SteelState|" & CStr(mDiagramExtensionEnabled), spec)
    Dim cached As CMaterialDiagram
    Set cached = CachedDiagram(cacheKey)
    If Not cached Is Nothing Then
        Set SteelMaterialForEquilibriumFromSpec = cached
        Exit Function
    End If

    Dim physical As CMaterialDiagram
    Set physical = SteelMaterialFromSpec(spec)
    Dim diagram As CMaterialDiagram
    If mDiagramExtensionEnabled Then
        Set diagram = DiagramWithExtension(physical, _
            mSteelParameters.Esc, mSteelParameters.Es, True, True)
    Else
        Set diagram = physical
    End If
    StoreCachedDiagram cacheKey, diagram
    Set SteelMaterialForEquilibriumFromSpec = diagram
End Function

' Возвращает предельную сжатую деформацию бетона для заданной профильной спецификации.
Public Function ConcreteCompressionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ConcreteCompressionLimitFromSpec = ConcreteMaterialFromSpec(spec).UltimateCompressionStrain
End Function

' Возвращает предельную растянутую деформацию бетона для заданной профильной спецификации.
Public Function ConcreteTensionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ConcreteTensionLimitFromSpec = ConcreteMaterialFromSpec(spec).UltimateTensionStrain
End Function

' Возвращает eps_bt0 для критерия образования трещины по СП 63 п. 8.1.30, формула (8.54).
' Использует те же параметры профиля, что диаграмма образования трещины,
' чтобы расчёт получал согласованные материалы без отдельного чтения Config.
Public Function ConcreteTensionEbt0FromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteTensionEbt0FromSpec = mConcreteParameters.Ebt0
End Function

' Показывает, активен ли физический предел растянутого бетона в данной спецификации.
Public Function ConcreteTensionLimitEnabledFromSpec(ByVal spec As CMaterialModelSpec) As Boolean
    Dim diagram As CMaterialDiagram
    Set diagram = ConcreteMaterialFromSpec(spec)
    ConcreteTensionLimitEnabledFromSpec = (diagram.HasPhysicalTensionLimit And spec.ConcreteTension = ctkUseDiagram)
End Function

' Возвращает предельную сжатую деформацию арматуры для заданной профильной спецификации.
Public Function SteelCompressionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    SteelCompressionLimitFromSpec = SteelMaterialFromSpec(spec).UltimateCompressionStrain
End Function

' Возвращает предельную растянутую деформацию арматуры для заданной профильной спецификации.
Public Function SteelTensionLimitFromSpec(ByVal spec As CMaterialModelSpec) As Double
    SteelTensionLimitFromSpec = SteelMaterialFromSpec(spec).UltimateTensionStrain
End Function

' Возвращает расчетное/нормативное сопротивление бетона сжатию по ValueSet спецификации.
Public Function ConcreteCompressionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteCompressionResistanceFromSpec = ConcreteCompressionResistanceFromValueSet(spec.ValueSet)
End Function

' Возвращает расчетное/нормативное сопротивление бетона растяжению по ValueSet спецификации.
Public Function ConcreteTensionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteTensionResistanceFromSpec = ConcreteTensionResistanceFromValueSet(spec.ValueSet)
End Function

' Возвращает Rb/Rb,ser по выбранному набору характеристик без привязки к
' диаграмме. Это нужно для проверок, где СП использует расчетные параметры
' материала, но не требует строить sigma-epsilon диаграмму.
Public Function ConcreteCompressionResistanceFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths LimitStateGroupFromValueSet(valueSet), rb, rbt
    ConcreteCompressionResistanceFromValueSet = rb
End Function

' Возвращает Rbt/Rbt,ser по выбранному набору характеристик.
Public Function ConcreteTensionResistanceFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    Dim rb As Double
    Dim rbt As Double
    ConcreteStrengths LimitStateGroupFromValueSet(valueSet), rb, rbt
    ConcreteTensionResistanceFromValueSet = rbt
End Function

' Возвращает модуль бетона для приведенных характеристик, связанных с указанной спецификацией.
Public Function ConcreteElasticModulusFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    ConcreteElasticModulusFromSpec = ConcreteElasticModulusFromValueSet(spec.ValueSet)
End Function

' Начальный модуль бетона хранится отдельно от R/Rser; valueSet оставлен в
' сигнатуре, чтобы вызывающий код явно передавал выбранную нормативную базу.
Public Function ConcreteElasticModulusFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    ConcreteElasticModulusFromValueSet = mConcreteParameters.Eb
End Function

' Возвращает модуль арматуры для приведенных характеристик, связанных с указанной спецификацией.
Public Function SteelElasticModulusFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    SteelElasticModulusFromSpec = SteelElasticModulusFromValueSet(spec.ValueSet)
End Function

' Возвращает модуль арматуры для выбранного набора характеристик.
' Сейчас Es общий, но этот метод позволяет получить его без создания лишней диаграммы.
Public Function SteelElasticModulusFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    SteelElasticModulusFromValueSet = mSteelParameters.Es
End Function

' Возвращает сопротивление арматуры сжатию по ValueSet спецификации.
Public Function SteelCompressionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    SteelCompressionResistanceFromSpec = SteelCompressionResistanceFromValueSet(spec.ValueSet)
End Function

' Возвращает сопротивление арматуры растяжению по ValueSet спецификации.
Public Function SteelTensionResistanceFromSpec(ByVal spec As CMaterialModelSpec) As Double
    ValidateSpec spec
    SteelTensionResistanceFromSpec = SteelTensionResistanceFromValueSet(spec.ValueSet)
End Function

' Возвращает Rsc/Rsc,ser по выбранному набору характеристик без построения
' диаграммы арматуры.
Public Function SteelCompressionResistanceFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths LimitStateGroupFromValueSet(valueSet), rsc, rs
    SteelCompressionResistanceFromValueSet = rsc
End Function

' Возвращает Rs/Rs,ser по выбранному набору характеристик.
Public Function SteelTensionResistanceFromValueSet(ByVal valueSet As EMaterialValueSet) As Double
    EnsureInitialized
    Dim rsc As Double
    Dim rs As Double
    SteelStrengths LimitStateGroupFromValueSet(valueSet), rsc, rs
    SteelTensionResistanceFromValueSet = rs
End Function

' ГПС определяется спецификацией профиля, а не расчетным классом.
Private Function LimitStateGroupFromSpec(ByVal spec As CMaterialModelSpec) As String
    ValidateSpec spec
    LimitStateGroupFromSpec = LimitStateGroupFromValueSet(spec.ValueSet)
End Function

' Преобразует ValueSet в номер ГПС для общих helpers сопротивлений.
Private Function LimitStateGroupFromValueSet(ByVal valueSet As EMaterialValueSet) As String
    If valueSet = mvsSLS Then
        LimitStateGroupFromValueSet = "II"
    Else
        LimitStateGroupFromValueSet = "I"
    End If
End Function

' Проверяет, заданы ли допустимые параметры модели материалов выбранного профиля.
Private Sub ValidateSpec(ByVal spec As CMaterialModelSpec)
    If spec Is Nothing Then Err.Raise vbObjectError + 3250, "CMaterialModelProvider", "Не передано описание выбранных диаграмм и характеристик материалов (MaterialModelSpec)."
    If Not spec.IsComplete Then Err.Raise vbObjectError + 3251, "CMaterialModelProvider", "В описании материалов указаны не все обязательные характеристики и диаграммы (MaterialModelSpec)."
End Sub

' Возвращает диаграмму из кэша, если такой MaterialModelSpec уже запрашивался.
Private Function CachedDiagram(ByVal cacheKey As String) As CMaterialDiagram
    If mDiagramCache Is Nothing Then Set mDiagramCache = CreateObject("Scripting.Dictionary")
    If mDiagramCache.Exists(cacheKey) Then Set CachedDiagram = mDiagramCache.Item(cacheKey)
End Function

' Сохраняет готовую диаграмму в кэш provider-а. Диаграммы используются как
' read-only объекты, поэтому повторная выдача одного экземпляра безопасна.
Private Sub StoreCachedDiagram(ByVal cacheKey As String, ByVal diagram As CMaterialDiagram)
    If diagram Is Nothing Then Exit Sub
    If mDiagramCache Is Nothing Then Set mDiagramCache = CreateObject("Scripting.Dictionary")
    If mDiagramCache.Exists(cacheKey) Then mDiagramCache.Remove cacheKey
    mDiagramCache.Add cacheKey, diagram
End Sub

' Формирует ключ кэша: роль диаграммы + полный профиль material spec.
Private Function MaterialDiagramCacheKey(ByVal roleKey As String, ByVal spec As CMaterialModelSpec) As String
    MaterialDiagramCacheKey = roleKey & "|" & spec.SpecKey
End Function

' Возвращает группу предельных состояний диаграмм, подготовленных для расчетной роли.
Public Function LimitStateGroup(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    LimitStateGroup = mLimitStateGroup(PhysicalPurposeSlot(purpose))
End Function

' Возвращает тип бетонной диаграммы, подготовленной для расчетной роли.
Public Function ConcreteDiagramType(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    ConcreteDiagramType = mConcreteDiagramType(PhysicalPurposeSlot(purpose))
End Function

' Возвращает режим растянутого бетона в диаграмме выбранной расчетной роли.
Public Function ConcreteTensionBehavior(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    ConcreteTensionBehavior = mConcreteTensionBehavior(PhysicalPurposeSlot(purpose))
End Function

' Возвращает тип диаграммы арматуры, подготовленной для расчетной роли.
Public Function SteelDiagramType(ByVal purpose As ECalculationPurpose) As String
    EnsureInitialized
    SteelDiagramType = mSteelDiagramType(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую предельную деформацию бетона при сжатии для расчетной роли.
Public Function ConcreteCompressionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteCompressionLimit = mConcreteCompressionLimit(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую предельную деформацию растянутого бетона для расчетной роли.
Public Function ConcreteTensionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteTensionLimit = mConcreteTensionLimit(PhysicalPurposeSlot(purpose))
End Function

' Показывает, действует ли физический предел растянутого бетона для расчетной роли.
Public Function ConcreteTensionLimitEnabled(ByVal purpose As ECalculationPurpose) As Boolean
    EnsureInitialized
    ConcreteTensionLimitEnabled = mConcreteTensionLimitEnabled(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую предельную деформацию арматуры при сжатии для расчетной роли.
Public Function SteelCompressionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelCompressionLimit = mSteelCompressionLimit(PhysicalPurposeSlot(purpose))
End Function

' Возвращает физическую предельную деформацию арматуры при растяжении для расчетной роли.
Public Function SteelTensionLimit(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelTensionLimit = mSteelTensionLimit(PhysicalPurposeSlot(purpose))
End Function

' Возвращает сопротивление бетона сжатию для диаграмм расчетной роли, МПа.
Public Function ConcreteCompressionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteCompressionResistance = mConcreteCompressionResistance(PhysicalPurposeSlot(purpose))
End Function

' Возвращает сопротивление бетона растяжению для диаграмм расчетной роли, МПа.
Public Function ConcreteTensionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteTensionResistance = mConcreteTensionResistance(PhysicalPurposeSlot(purpose))
End Function

' Возвращает модуль бетона для диаграмм расчетной роли, МПа.
Public Function ConcreteElasticModulus(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    ConcreteElasticModulus = mConcreteElasticModulus(PhysicalPurposeSlot(purpose))
End Function

' Возвращает сопротивление арматуры сжатию для диаграмм расчетной роли, МПа.
Public Function SteelCompressionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelCompressionResistance = mSteelCompressionResistance(PhysicalPurposeSlot(purpose))
End Function

' Возвращает сопротивление арматуры растяжению для диаграмм расчетной роли, МПа.
Public Function SteelTensionResistance(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelTensionResistance = mSteelTensionResistance(PhysicalPurposeSlot(purpose))
End Function

' Возвращает модуль арматуры для диаграмм расчетной роли, МПа.
Public Function SteelElasticModulus(ByVal purpose As ECalculationPurpose) As Double
    EnsureInitialized
    SteelElasticModulus = mSteelElasticModulus(PhysicalPurposeSlot(purpose))
End Function

' Показывает, разрешено ли вспомогательное продолжение диаграмм при поиске равновесия.
Public Property Get DiagramExtensionEnabled() As Boolean
    EnsureInitialized
    DiagramExtensionEnabled = mDiagramExtensionEnabled
End Property

' Коэффициент к модулю соответствующей ветви для ненулевых физических плато.
Public Property Get PlateauSlopeFactor() As Double
    EnsureInitialized
    PlateauSlopeFactor = mPlateauSlopeFactor
End Property

' Возвращает версию подготовленного набора материалов. Вместе с самим объектом
' подготовки она позволяет проверить пригодность ранее найденного НДС;
' одинакового имени профиля для этого недостаточно.
Public Property Get Revision() As Long
    EnsureInitialized
    Revision = mRevision
End Property

' Возвращает отдельные копии исходных параметров. Их изменение не затрагивает
' внутренние характеристики и уже подготовленные диаграммы этого объекта.
Public Property Get ConcreteParameters() As CConcreteMaterialParameters
    EnsureInitialized
    Set ConcreteParameters = mConcreteParameters.Clone
End Property

' Исходные параметры арматуры из Config.
Public Property Get SteelParameters() As CSteelMaterialParameters
    EnsureInitialized
    Set SteelParameters = mSteelParameters.Clone
End Property

' Проверяет, что provider уже построил диаграммы.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3242, "CMaterialModelProvider", "Поставщик материалов не инициализирован."
End Sub

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' позволяет задать материалы программно для проверок диаграмм и расчётных случаев
' без чтения Config. Обычный запуск книги использует Initialize(settings, units).
Public Sub TestInitializeFromParameters(ByVal concreteParameters As CConcreteMaterialParameters, _
        ByVal steelParameters As CSteelMaterialParameters, _
        Optional ByVal strengthConcreteDiagram As String = "TwoLine", _
        Optional ByVal strengthConcreteTension As String = "Ignore", _
        Optional ByVal strengthSteelDiagram As String = "TwoLine", _
        Optional ByVal mcrcConcreteDiagram As String = "ThreeLine", _
        Optional ByVal mcrcSteelDiagram As String = "TwoLine", _
        Optional ByVal crackedConcreteDiagram As String = "TwoLine", _
        Optional ByVal crackedSteelDiagram As String = "TwoLine", _
        Optional ByVal diagramExtensionEnabled As Boolean = True, _
        Optional ByVal plateauSlopeFactor As Double = 0#)
    Set mInputLocations = Nothing
    If concreteParameters Is Nothing Then Err.Raise vbObjectError + 3243, "CMaterialModelProvider", "Параметры бетона не переданы."
    If steelParameters Is Nothing Then Err.Raise vbObjectError + 3244, "CMaterialModelProvider", "Параметры арматуры не переданы."
    If plateauSlopeFactor < 0# Then Err.Raise vbObjectError + 3247, "CMaterialModelProvider", "Коэффициент наклона плато должен быть не меньше нуля."

    mInitialized = False
    mRevision = mRevision + 1
    Set mConcreteParameters = concreteParameters.Clone
    Set mSteelParameters = steelParameters.Clone
    Set mDiagramCache = CreateObject("Scripting.Dictionary")
    mDiagramExtensionEnabled = diagramExtensionEnabled
    mPlateauSlopeFactor = plateauSlopeFactor
    BuildPurpose cpStrength, "I", strengthConcreteDiagram, strengthConcreteTension, strengthSteelDiagram
    BuildPurpose cpMcrc, "II", mcrcConcreteDiagram, "UseDiagram", mcrcSteelDiagram
    BuildPurpose cpCrackedNDS, "II", crackedConcreteDiagram, "Ignore", crackedSteelDiagram
    mInitialized = True
End Sub
