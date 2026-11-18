# Каталог параметров кручения используется сборщиком и проверкой покрытия справки.
function Get-TorsionSettingsCatalog {
    return @(
        @('Torsion.StateType','StrengthState','-','Сохранённое НДС для суммы напряжений NMT; не запускает получение другого состояния.'),
        @('Torsion.HoopAxisOffset',29,'мм','Расстояние от поверхности бетона до оси наружного хомута; не меньше половины его диаметра.'),
        @('Torsion.HoopDiameter',8,'мм','Диаметр одной ветви замкнутого хомута; положительное число.'),
        @('Torsion.HoopSpacing',100,'мм','Шаг замкнутых хомутов вдоль элемента; положительное число.'),
        @('Torsion.Rs','=INDEX(rngSteelMaterialParameters,2,3)','МПа','Допустимое растягивающее напряжение продольной арматуры в отдельной проверке; больше нуля.'),
        @('Torsion.Rsw',300,'МПа','Принятое сопротивление ветви хомута растяжению; больше нуля.'),
        @('Torsion.BarBand',40,'мм','Полоса отбора стержней внутрь от оси наружного хомута, >0; габаритом сечения не ограничена. Арматура у отверстия исключается.'),
        @('Torsion.StrutAngle',45,'deg','Угол бетонных раскосов к продольной оси, строго между 0 и 90°. Базовая ферма — 45°.'),
        @('Torsion.ConcreteMethod','Strut','-','Метод бетона: Strut, GorskyRectangle или None. Ошибка выбора блокирует только бетон; арматура и хомут рассчитываются.'),
        @('Torsion.WallThickness',0,'мм','Толщина бетона для раскосов Strut: 0 — авто, >0 — заданная толщина. Отбор стержней задаёт Torsion.BarBand.'),
        @('Torsion.Rb','=INDEX(rngConcreteMaterialParameters,2,2)','МПа','Сопротивление бетона для T, >0; в GorskyRectangle применяется как Rи.'),
        @('Torsion.ConcreteEfficiency',0.5,'-','Для Strut: доля сопротивления бетона, больше 0 и не больше 1. Инженерное допущение.'),
        @('Torsion.GorskyCoefficient',0.07,'-','Для сплошного прямоугольника: коэффициент c >0 в T = c Rи b²h (GorskyRectangle).')
    )
}

# Создаёт две входные таблицы и два самостоятельных якоря T / NMT.
# Размеры и положение задаёт WorkbookLayout; существующие таблицы NM не перемещаются.
function Add-TorsionPostprocessTables {
    param([object]$Workbook,[object]$Layout)
    foreach($name in @('rngTorsionCombinations','rngTorsionSettings','rngTorsionSummaryAnchor','rngTorsionBarForces')) {
        $item=@($Layout.Names|Where-Object Name -eq $name)
        if($item.Count -ne 1){throw ('Не задано положение '+$name)}
        $sheet=$Workbook.Worksheets.Item($item[0].Sheet)
        Add-WorkbookName $Workbook $name $sheet $item[0].Address
    }
    $loads=$Workbook.Names.Item('rngTorsionCombinations').RefersToRange
    $sheet=$loads.Worksheet
    $loads.Font.Name='Arial';$loads.Font.Size=9;$loads.HorizontalAlignment=-4108;$loads.VerticalAlignment=-4108;$loads.WrapText=$false
    $loads.Borders.LineStyle=1;$loads.Borders.Weight=2;$loads.Borders.Color=0
    $loads.Rows.Item(1).Interior.Color=14277081;$loads.Rows.Item(1).Font.Bold=$true
    $loads.Cells.Item(1,1).Value2='ID'
    $loads.Cells.Item(1,2).Formula='="T, "&INDEX(rngUnitSettings,MATCH("Moment",INDEX(rngUnitSettings,,1),0),2)'
    for($row=2;$row -le $loads.Rows.Count;$row++){
        $loads.Cells.Item($row,1).Formula='=IF(INDEX(rngLoadCombinations,'+$row+',1)="","",INDEX(rngLoadCombinations,'+$row+',1))'
        $loads.Cells.Item($row,2).Interior.Color=15135487 # RGB(255,242,230): ввод T; вычисляемый ID этим цветом не выделяется.
    }
    $title=$loads.Cells.Item(1,1).Offset(-1,0).Resize(1,2);$title.Merge();$title.Value2='T (кручение)'
    $title.Font.Name='Arial';$title.Font.Size=9;$title.Font.Bold=$true;$title.Font.Color=0;$title.Interior.Color=15921906

    $settings=$Workbook.Names.Item('rngTorsionSettings').RefersToRange
    $settings.Font.Name='Arial';$settings.Font.Size=9;$settings.HorizontalAlignment=-4108;$settings.VerticalAlignment=-4108;$settings.WrapText=$false
    $settings.Borders.LineStyle=1;$settings.Borders.Weight=2;$settings.Borders.Color=0
    $headers=@('Параметр','Значение','Ед.','Комментарий','Справка')
    for($column=1;$column -le 5;$column++){$settings.Cells.Item(1,$column).Value2=$headers[$column-1]}
    $settings.Rows.Item(1).Interior.Color=14277081;$settings.Rows.Item(1).Font.Bold=$true
    $rows=@(Get-TorsionSettingsCatalog)
    if($settings.Rows.Count -ne $rows.Count+1){throw 'Размер rngTorsionSettings не совпадает с каталогом.'}
    for($index=0;$index -lt $rows.Count;$index++){
        $row=$index+2;$entry=$rows[$index]
        $settings.Cells.Item($row,1).Value2=[string]$entry[0]
        $value=$settings.Cells.Item($row,2)
        if($entry[1] -is [string] -and $entry[1].StartsWith('=')){$value.Formula=$entry[1]}
        elseif($entry[1] -is [string]){$value.Value2=[string]$entry[1]}
        else{$value.Value2=[double]$entry[1]}
        $value.Interior.Color=15135487
        Set-InputUnitCell $settings.Cells.Item($row,3) ([string]$entry[2])
        $settings.Cells.Item($row,4).Value2=[string]$entry[3]
        $settings.Cells.Item($row,4).HorizontalAlignment=-4131
        if($entry[0] -in @('Torsion.StateType','Torsion.ConcreteMethod')){
            $items=if($entry[0] -eq 'Torsion.StateType'){@('StrengthState','CapacityState','PreCrackState','PostCrackState','CrackedState')}else{@('Strut','GorskyRectangle','None')}
            $separator=[string]$Workbook.Application.International(5)
            $value.Validation.Delete();$value.Validation.Add(3,1,1,($items -join $separator))
            $value.Validation.ShowError=$true;$value.Validation.InCellDropdown=$true
            $value.Validation.ErrorTitle='Параметр кручения';$value.Validation.ErrorMessage='Выберите один пункт из списка.'
        }
    }
    $title=$settings.Cells.Item(1,1).Offset(-1,0).Resize(1,5);$title.Merge();$title.Value2='Постпроверка T (кручение) и NMT'
    $title.Font.Name='Arial';$title.Font.Size=9;$title.Font.Bold=$true;$title.Font.Color=0;$title.Interior.Color=15921906
    $bars=$Workbook.Names.Item('rngTorsionBarForces').RefersToRange
    $title=$bars.Cells.Item(1,1).Offset(-1,0).Resize(1,4);$title.Merge();$title.Value2='Усилия от T (кручение)'
    $title.Font.Name='Arial';$title.Font.Size=9;$title.Font.Bold=$true;$title.Font.Color=0;$title.Interior.Color=14277081
}

# Методика входит в общий порядок разделов справки; ссылки отдельных настроек
# ведут к их собственным описаниям в общем каталоге Config.
function Add-TorsionMethodologyGuide {
    param([object]$sheet,[int]$row)
    $row=Add-GuideTitle $sheet $row 'Исходные данные и модель' 12
    $row=Add-GuideParagraph $sheet $row 'Задайте T в отдельной таблице по ID основного сочетания. Пустой ID, пустой T или нулевой T отключают строку. Используется однокамерная пространственная ферма с одной наружной осью замкнутого хомута. Расчёт НДМ N+Mx+My и его собственные результаты сохраняют самостоятельный смысл.'
    $row=Add-GuideParagraph $sheet $row 'Размеры вводятся в INPUT Length, сопротивления — в INPUT Stress, крутящий момент — в INPUT Moment. Подписи единиц являются формулами по rngUnitSettings. При смене единиц числовые значения ввода должны соответствовать выбранной системе. Угол раскосов задаётся в градусах; коэффициенты безразмерны.'
    $row=Add-GuideParagraph $sheet $row 'Уравнения пространственной фермы приведены в ACI 445.1R-12, Report on Torsion in Structural Concrete, глава 9. Для угла 45° зависимости арматуры сопоставляются с книгой В. В. Горского «Проектирование железобетонных конструкций, подверженных кручению», §4. Формулы источников записаны с обозначениями книги: сохраняется их структура, а сопоставление параметров поясняется сразу под формулой. Приписка «Модель книги» обозначает расчётное преобразование, определение или инженерное правило без собственного номера формулы внешнего источника. Проверки не заменяют полный набор требований ACI 318, EC2 или СП.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' '§9.1–9.2.2, с.9–11' -PdfPage 10
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' '§4, формулы (17)–(19), с.27'
    $row=Add-GuideParagraph $sheet $row 'Выберите сохранённое состояние НДМ, по умолчанию StrengthState. Используются его готовые напряжения; НДС от T не уравновешивается. Сошедшееся состояние доступно и при выходе за физические диаграммы: его собственные отказы сохраняются в блоках НДМ. При отсутствии сохранённых напряжений рассчитываются проверки от T, а сумма NMT получает N/A.'
    $row=Add-GuideParagraph $sheet $row 'Ось наружного хомута строится по Torsion.HoopAxisOffset — расстоянию от поверхности бетона до оси ветви. Диаметр ветви к этому расстоянию не добавляется. Площадь A_o определяется внутри полученной оси, периметр u_o — вдоль неё. Дополнительные коэффициенты площади и понижающие коэффициенты прочности из других методов автоматически не переносятся.'
    $row=Add-GuideTitle $sheet $row 'Продольная арматура и хомут' 12
    $row=Add-GuideParagraph $sheet $row 'Torsion.BarBand задаёт ширину отбора центров продольных стержней внутрь от оси наружного хомута по кратчайшему расстоянию. Ширина положительна и не ограничена габаритом сечения: включаются фактически попавшие в полосу пригодные стержни. Площадь пустой части полосы к площади арматуры не прибавляется.'
    $row=Add-GuideParagraph $sheet $row 'В однокамерной модели выбирается наружный ряд. Стержни, ближайшие к границе отверстия, и равноудалённые от неё и наружной границы исключаются при любой ширине полосы. Нужны минимум три участвующих стержня и достаточное покрытие периметра их проекциями.'
    $row=Add-GuideFractionFormula $sheet $row 'q = ' '|T|' '2 A_o' '     По ACI 445.1R-12, (9.2.1a), с.9'
    $row=Add-GuideFractionFormula $sheet $row 'A_s,T,req = ' '|T| u_o cot θ' '2 A_o R_s' '     По ACI 445.1R-12, (9.2.2f), с.11'
    $row=Add-GuideFormula $sheet $row 'F_T = q u_o cot θ      Модель книги: продольная тяга'
    $row=Add-GuideParagraph $sheet $row 'Касательный поток q показывает усилие на единицу длины замкнутого пути по оси наружного хомута. Для распределения продольной тяги каждый участвующий стержень проецируется на эту ось. Точки проекций называются станциями; они следуют друг за другом по периметру.'
    $row=Add-GuideParagraph $sheet $row 'При отдельных однозначных станциях каждому стержню назначается прилегающий участок периметра. Его границы находятся посередине пути до предыдущей и следующей станций. Расстояние измеряется вдоль оси хомута, включая повороты и дуги. Положение этой оси задаётся Torsion.HoopAxisOffset; линия центров продольных стержней может располагаться иначе. Длина назначенного участка равна половине суммы двух соседних расстояний по оси хомута.'
    $row=Add-GuideFractionFormula $sheet $row 'l_i = ' 's_{i-1,i} + s_{i,i+1}' '2' '     Модель книги: прилегающая длина'
    $row=Add-GuideFormula $sheet $row 'ΔF_T,i = q l_i cot θ      Модель книги: тяга на участке'
    $row=Add-GuideParagraph $sheet $row 'Участок углового стержня может включать части двух соседних граней. Если он длиннее участков других стержней, на него приходится большее усилие. Сам угол дополнительного множителя не даёт. Если стержня в углу нет, общая граница участков ближайших к углу стержней всё равно находится посередине пути между их станциями и может лежать на любой из двух граней.'
    $row=Add-GuideParagraph $sheet $row 'Соседние участки стыкуются по общей границе: каждый отрезок или дуга периметра относится к одному участку, без пропусков и повторного учёта. Если несколько стержней имеют одну станцию, участок у них общий, а усилие делится пропорционально площадям стержней. Обычные равные проекции углового стержня на две соседние грани объединяются в одну станцию общего угла.'
    $row=Add-GuideParagraph $sheet $row 'Если одинаково близки удалённые участки оси, используются все равноправные проекции. Физическая площадь стержня делится поровну между отдельными точками и связными дуговыми участками. Полученные усилия складываются под одним ID стержня. Ни площадь, ни отдельный участок периметра не учитываются повторно; выбор не зависит от нумерации и направления обхода контура.'
    $row=Add-GuideParagraph $sheet $row 'В центре круговой дуги ближайшие проекции образуют непрерывный участок. Его доля площади равномерно распределяется вдоль дуги. Для точечных групп площадь относится к их прилегающей длине. Тяга делится пропорционально этим плотностям площади с точным суммированием по границам участков. Это инженерное правило распределения; деление одной дуги на несколько рёбер не меняет результат.'
    $row=Add-GuideFractionFormula $sheet $row 'Δσ_T,i = ' 'ΔF_T,i' 'A_s,i' '     Модель книги: распределение тяги'
    $row=Add-GuideFormula $sheet $row 'Σ_i ΔF_T,i = F_T      Модель книги: равновесие тяги'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Продольная арматура: (9.2.2f); поперечная арматура: (9.2.2g), с.11' -PdfPage 12
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' 'Хомут: (17); продольные стержни: (18), (19), с.27'
    $row=Add-GuideCompositeFormula $sheet $row @(
        @{Kind='Fraction';Numerator='A_sw';Denominator='s'},
        @{Kind='Text';Text=' = '},
        @{Kind='Fraction';Numerator='|T|';Denominator='2 A_o R_sw cot θ'}
    ) 'По ACI 445.1R-12, (9.2.2g), с.11' '(A_sw) / (s) = (|T|) / (2 A_o R_sw cot θ)'
    $row=Add-GuideFractionFormula $sheet $row 'σ_sw = ' 'q s' 'A_sw cot θ' '≤ R_sw      Модель книги: напряжение хомута'
    $row=Add-GuideParagraph $sheet $row 'Для равномерных продольных стержней на отдельных станциях при угле 45°, равных отношениях площади к соответствующему шагу и одинаковых сопротивлениях напряжения и запасы от T совпадают. Продольный шаг измеряется по периметру хомута, шаг хомутов — вдоль элемента. Сумма NMT дополнительно учитывает сохранённые напряжения от N и M.'
    $row=Add-GuideTitle $sheet $row 'Метод проверки бетона: Strut' 12
    $row=Add-GuideParagraph $sheet $row 'Strut проверяет сжатие диагональных бетонных раскосов. Torsion.WallThickness задаёт полную эффективную толщину бетонной полосы; отбор арматуры задаётся Torsion.BarBand. При одном потоке большая толщина даёт меньшее напряжение раскоса. Толщина не меняет суммарную продольную тягу или напряжение хомута.'
    $row=Add-GuideFormula $sheet $row 't_eff,auto = 2 a_min      Модель книги: автоматическая толщина'
    $row=Add-GuideFormula $sheet $row 'R_diag = ν R_b      Модель книги: сопротивление раскоса'
    $row=Add-GuideParagraph $sheet $row 'Ноль в Torsion.WallThickness включает автоматическую толщину. Положительное значение задаёт её явно и должно помещаться в располагаемой бетонной полосе вокруг оси. Автоматическая толщина и коэффициент ν по умолчанию 0,5 являются инженерными параметрами модели. В ACI §9.2.2 описана эффективная стенка EC2 со своими условиями; она не подменяет это правило.'
    $row=Add-GuideFormula $sheet $row 'T_lim = 2 R_diag A_o t_eff sin θ cos θ      По ACI 445.1R-12, (9.2.2b), с.11'
    $row=Add-GuideFractionFormula $sheet $row 'σ_diag = ' 'q (cot θ + tan θ)' 't_eff' '≤ R_diag      Модель книги: напряжение раскоса'
    $row=Add-GuideFractionFormula $sheet $row 'k_b = ' 'R_diag' 'σ_diag' '     Модель книги: коэффициент запаса'
    $row=Add-GuideParagraph $sheet $row 'Предельный момент и запас выражают одно условие сопротивления диагональных раскосов. Нормальные напряжения бетона НДМ к диагональным напряжениям не прибавляются.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Пространственная ферма: §9.1; бетонные раскосы: §9.2.2, (9.2.2b), с.9–11' -PdfPage 12
    $row=Add-GuideTitle $sheet $row 'Метод проверки бетона: GorskyRectangle и None' 12
    $row=Add-GuideParagraph $sheet $row 'GorskyRectangle определяет предел бетона одного сплошного прямоугольника с прямыми углами по условию Горского. Поворот прямоугольника допускается; отверстия, фаски, дуги и произвольный контур не соответствуют формуле. В Torsion.Rb задаётся принятое сопротивление R_и, в Torsion.GorskyCoefficient — коэффициент c, по умолчанию 0,07. Соответствие сопротивления условиям источника проверяется отдельно.'
    $row=Add-GuideFormula $sheet $row '|T| ≤ c R_и b² h      Горский, §3, (13), с.23'
    $row=Add-GuideFormula $sheet $row 'T_lim = c R_и b² h      Модель книги: предел бетона'
    $row=Add-GuideFractionFormula $sheet $row 'k_b = ' 'T_lim' '|T|' '     Модель книги: коэффициент запаса'
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' '§3, формула (13), с.23'
    $row=Add-GuideParagraph $sheet $row 'GorskyRectangle не читает толщину Torsion.WallThickness и коэффициент Torsion.ConcreteEfficiency. None явно отключает только бетон: статус равен N/A, запас и предельный момент отсутствуют; неиспользуемые поля бетона не читаются. Продольная арматура и хомут рассчитываются общей пространственной фермой.'
    $row=Add-GuideParagraph $sheet $row 'Неизвестный метод, ошибка его формулы или непригодная для него геометрия дают понятную причину и InputErr только в бетонной проверке. Другие доступные проверки T и сумма NMT выполняются. Общий статус при неподтверждённой из-за ошибки проверке бетона не становится OK.'
    $row=Add-GuideTitle $sheet $row 'Сумма напряжений NMT и вывод' 12
    $row=Add-GuideFormula $sheet $row 'σ_NMT,i = σ_NM,i + Δσ_T,i;  σ_NMT,i ≤ R_s      Модель книги: сумма напряжений'
    $row=Add-GuideFractionFormula $sheet $row 'k_s,NMT = ' 'R_s' 'σ_NMT,i' '     Модель книги: положительная сумма'
    $row=Add-GuideParagraph $sheet $row 'Используется алгебраическая сумма сохранённого напряжения НДМ и растягивающей добавки от T. Она не восстанавливает совместное равновесие и перераспределение НДС и не определяет совместный предельный M или N. Горский на с.24 обсуждает приближённость раздельного расчёта изгиба и кручения с суммированием требуемой арматуры. Ошибка метода бетона не меняет саму сумму напряжений, но оставляет бетонную проверку неподтверждённой.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Раздельный расчёт арматуры: §9.1, с.9' -PdfPage 10
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' 'Ограничения раздельного суммирования: с.24'
    $row=Add-GuideParagraph $sheet $row 'На Results после устойчивости выводятся проверки арматуры, хомута и бетона от T, затем сумма NMT. Поэлементная таблица участвующих стержней содержит ID сочетания, Rebar, усилие F_T и назначенную длину l_i в OUTPUT Length. Длина соответствует доле периметра физического стержня, включая общие и множественные проекции; сумма длин равна u_o. Сводка содержит три статуса и запаса: бетон, продольная арматура NMT и хомут. Причины НДМ остаются в его блоках; ошибки T содержат собственные причины и цепочки адресов.'
    return $row
}
