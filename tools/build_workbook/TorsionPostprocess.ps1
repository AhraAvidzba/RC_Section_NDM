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
    $title=$bars.Cells.Item(1,1).Offset(-1,0).Resize(1,3);$title.Merge();$title.Value2='Усилия от T (кручение)'
    $title.Font.Name='Arial';$title.Font.Size=9;$title.Font.Bold=$true;$title.Font.Color=0;$title.Interior.Color=14277081
}

# Методика входит в общий порядок разделов справки; ссылки отдельных настроек
# ведут к их собственным описаниям в общем каталоге Config.
function Add-TorsionMethodologyGuide {
    param([object]$sheet,[int]$row)
    $row=Add-GuideTitle $sheet $row 'Исходные данные и модель' 12
    $row=Add-GuideParagraph $sheet $row 'Задайте T в отдельной таблице по ID основного сочетания. Пустой ID, пустой T или нулевой T отключают строку. Метод представляет отдельную проверку однокамерной пространственной фермы с одной наружной осью замкнутого хомута. Расчёт НДМ N+Mx+My и его собственные результаты сохраняют самостоятельный смысл.'
    $row=Add-GuideParagraph $sheet $row 'Все размеры вводятся в INPUT Length, сопротивления — в INPUT Stress, крутящий момент — в INPUT Moment. Подписи единиц являются формулами по rngUnitSettings. При смене единиц числовые значения ввода должны соответствовать выбранной системе. Угол раскосов задаётся в градусах; коэффициенты безразмерны.'
    $row=Add-GuideParagraph $sheet $row 'Уравнения равновесия пространственной фермы рассматриваются в ACI 445.1R-12, Report on Torsion in Structural Concrete, глава 9. Для угла 45° зависимости для арматуры сопоставляются с книгой В. В. Горского «Проектирование железобетонных конструкций, подверженных кручению», §4. Параметры модели книги и её проверки не заменяют полный набор требований ACI 318, EC2 или СП.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' '§9.1–9.2.2, стр.9–11; стр.10–12 PDF издателя' -PdfPage 10
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' '§4, формулы (17)–(19), стр.27; стр.25 предоставленного PDF'
    $row=Add-GuideParagraph $sheet $row 'Выберите сохранённое состояние НДМ, по умолчанию StrengthState. Используются его готовые напряжения; новое НДС от T не уравновешивается. Сошедшееся состояние доступно и при выходе за физические диаграммы: его собственные отказы сохраняются в блоках НДМ. При отсутствии сохранённых напряжений отдельно рассчитываются проверки от T, а сумма NMT получает N/A.'
    $row=Add-GuideParagraph $sheet $row 'Ось наружного хомута строится по Torsion.HoopAxisOffset — расстоянию от поверхности бетона до оси ветви. Диаметр ветви к этому расстоянию не добавляется. Площадь A_o определяется внутри полученной оси, периметр u_o — вдоль неё. Значение A_o является геометрической характеристикой этой оси; дополнительные коэффициенты площади и понижающие коэффициенты прочности из других методов автоматически не переносятся.'
    $row=Add-GuideTitle $sheet $row 'Продольная арматура и хомут' 12
    $row=Add-GuideParagraph $sheet $row 'Torsion.BarBand задаёт ширину отбора центров продольных стержней внутрь от оси наружного хомута. Пригодный стержень должен лежать внутри этой оси, а его кратчайшее расстояние до неё не должно превышать заданную ширину. Ширина положительна и не ограничена максимальным размером сечения: большая полоса включает фактически попавшие в неё пригодные стержни. Площадь пустой части полосы в расчёт арматуры не входит.'
    $row=Add-GuideParagraph $sheet $row 'В однокамерной модели участвует наружный ряд: стержни, ближайшая бетонная граница которых относится к отверстию, и равноудалённые от наружной границы и отверстия исключаются при любой ширине полосы. Нужны минимум три различные станции по периметру. Неоднозначная проекция стержня на разные участки оси требует уточнения геометрии. Эти правила сохраняют принятую схему одного наружного хомута.'
    $row=Add-GuideFractionFormula $sheet $row 'q = ' '|T|' '2 A_o'
    $row=Add-GuideFormula $sheet $row 'F_T = q u_o cot θ;  A_s,T,req = F_T / R_s'
    $row=Add-GuideParagraph $sheet $row 'Тяга F_T распределяется по прилегающим участкам периметра: каждому стержню относится половина расстояния до предыдущей и следующей станций. Стержни на одной станции делят её тягу пропорционально площади. Напряжение от T равно выделенной стержню тяге, делённой на его площадь; суммарные усилия стержней равны F_T.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Формула (9.2.1a); §9.2.2, шаги расчёта арматуры' -PdfPage 10
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' '§4, формулы (17)–(19), стр.27; PDF: стр.25'
    $row=Add-GuideFractionFormula $sheet $row 'σ_sw = ' 'q s' 'A_sw cot θ'
    $row=Add-GuideParagraph $sheet $row 'Для равномерных продольных стержней на отдельных станциях при θ = 45°, равных отношениях площади к соответствующему шагу и одинаковых сопротивлениях напряжения и запасы от T совпадают. Для равномерных продольных стержней шаг измеряется по периметру хомута, для хомутов — вдоль элемента. Достаточный частный случай: одинаковые диаметры, равные указанные шаги и R_s = R_sw. При разных сопротивлениях равенство напряжений само по себе не означает равенства запасов; сумма NMT также учитывает напряжения от N и M.'
    $row=Add-GuideTitle $sheet $row 'Метод проверки бетона: Strut' 12
    $row=Add-GuideParagraph $sheet $row 'Strut проверяет сжатие диагональных бетонных раскосов пространственной фермы. Torsion.WallThickness задаёт полную эффективную толщину бетонной полосы, работающей с касательным потоком, а не расстояние отбора арматуры. При одном потоке большая принятая толщина даёт меньшее напряжение раскоса. Эта толщина используется только в бетонной проверке и не меняет выделенную продольную тягу или напряжение хомута.'
    $row=Add-GuideFormula $sheet $row 't_eff,auto = 2 a_min;  R_diag = ν R_b'
    $row=Add-GuideParagraph $sheet $row 'Значение Torsion.WallThickness = 0 включает указанное автоматическое правило. Положительное значение задаёт толщину явно; оно должно помещаться в располагаемую бетонную полосу вокруг оси. Параметр ширины отбора стержней Torsion.BarBand остаётся независимым и может превышать габарит сечения. Автоматическая толщина 2a_min и значение ν = 0,5 по умолчанию являются явными инженерными допущениями данной модели.'
    $row=Add-GuideFractionFormula $sheet $row 'σ_diag = ' 'q (cot θ + tan θ)' 't_eff'
    $row=Add-GuideParagraph $sheet $row 'Запас бетона равен R_diag/σ_diag. Равнозначная предельная величина момента получается из того же условия: T_lim = 2A_o t_eff R_diag/(cot θ + tan θ), где A_o — площадь внутри оси хомута, t_eff — эффективная толщина, R_diag — допустимое сжатие раскоса, θ — угол раскоса. К нормальным напряжениям бетона НДМ эта величина не прибавляется. В обзоре ACI §9.2.2 приведён расчёт эффективной стенки по EC2 с собственными условиями; он не является автоматическим правилом толщины этой книги.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Пространственная ферма и эффективная стенка: §9.1–9.2.2' -PdfPage 11
    $row=Add-GuideTitle $sheet $row 'Метод проверки бетона: GorskyRectangle и None' 12
    $row=Add-GuideParagraph $sheet $row 'GorskyRectangle определяет предел бетона сплошного прямоугольника с прямыми углами по эмпирическому условию Горского. b — меньшая сторона прямоугольника, h — большая. Поворот прямоугольника допускается; отверстия, фаски, дуги и произвольный контур не соответствуют этой формуле. Torsion.Rb здесь используется как принятое сопротивление R_и, Torsion.GorskyCoefficient — как коэффициент c, по умолчанию 0,07. Соответствие принятого R_и условиям применения формулы проверяется отдельно.'
    $row=Add-GuideFormula $sheet $row 'T_lim = c R_и b² h'
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' '§3, формула (13), стр.23; стр.21 предоставленного PDF'
    $row=Add-GuideParagraph $sheet $row 'Запас этого метода равен T_lim/|T|, где T_lim — момент по формуле Горского, |T| — модуль заданного момента. Толщина Torsion.WallThickness и коэффициент Torsion.ConcreteEfficiency этим методом не читаются. Закон продольной тяги и проверка хомута сохраняются; метод не объявляет произвольное сечение эквивалентным прямоугольнику.'
    $row=Add-GuideParagraph $sheet $row 'None явно отключает только проверку бетона. Её статус равен N/A, запаса и предельного момента нет; неиспользуемые поля бетона не читаются. Неизвестный метод, ошибка его формулы или непригодная для выбранного метода геометрия дают понятную причину и InputErr в бетонной проверке. Продольная арматура, хомут и доступная сумма NMT рассчитываются независимо. Общий статус при ошибке бетона не может быть OK.'
    $row=Add-GuideTitle $sheet $row 'Сумма напряжений NMT и вывод' 12
    $row=Add-GuideFormula $sheet $row 'σ_NMT,i = σ_NM,i + Δσ_T,i;  σ_NMT,i ≤ R_s'
    $row=Add-GuideParagraph $sheet $row 'Это арифметическая постпроверка сохранённых напряжений НДМ и растягивающей добавки от T. Она не восстанавливает совместное равновесие и перераспределение НДС и не даёт нового предельного M или N. Горский на стр.24 книги отдельно оговаривает приближённость раздельного расчёта изгиба и кручения с суммированием требуемой арматуры. Ошибка метода бетона не меняет саму арифметическую сумму напряжений, но оставляет бетонную проверку неподтверждённой.'
    $row=Add-GuideSourceReference $sheet $row 'ACI445' 'Раздельное суммирование арматуры: §9.1, стр.9; PDF: стр.10' -PdfPage 10
    $row=Add-GuideSourceReference $sheet $row 'Gorsky' 'Ограничения раздельного суммирования: стр.24; PDF: стр.22'
    $row=Add-GuideParagraph $sheet $row 'Подробная таблица T / NMT после устойчивости на Results сначала показывает только T: отбор продольной арматуры, хомут и бетон, затем сумму NMT для определяющего стержня. Внизу сохраняются ID сочетания, Rebar и дополнительное усилие от T. Общая сводка выводит три независимых статуса и три запаса: бетон, продольная арматура NMT и хомут. Причины НДМ остаются в его собственных блоках. Ошибки кручения имеют причины и цепочки адресов соответствующих настроек.'
    return $row
}
