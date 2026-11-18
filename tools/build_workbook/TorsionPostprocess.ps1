# Каталог параметров кручения используется сборщиком и проверкой покрытия справки.
function Get-TorsionSettingsCatalog {
    return @(
        @('Torsion.StateType','StrengthState','-','Сохранённое НДС для суммы напряжений NMT; не запускает получение другого состояния.'),
        @('Torsion.HoopAxisOffset',29,'mm','Расстояние от поверхности бетона до оси наружного хомута; не меньше половины его диаметра.'),
        @('Torsion.HoopDiameter',8,'mm','Диаметр одной ветви замкнутого хомута; положительное число.'),
        @('Torsion.HoopSpacing',100,'mm','Шаг замкнутых хомутов вдоль элемента; положительное число.'),
        @('Torsion.Rs','=INDEX(rngSteelMaterialParameters,2,3)','МПа','Допустимое растягивающее напряжение продольной арматуры в отдельной проверке; больше нуля.'),
        @('Torsion.Rsw',300,'МПа','Принятое сопротивление ветви хомута растяжению; больше нуля.'),
        @('Torsion.BarBand',40,'mm','Ширина полосы от оси наружного хомута, >0. Стержни, ближайшие к отверстию, исключаются.'),
        @('Torsion.StrutAngle',45,'deg','Угол бетонных раскосов к продольной оси, строго между 0 и 90°. Базовая ферма — 45°.'),
        @('Torsion.ConcreteMethod','Strut','-','Strut — раскосы; GorskyRectangle — сплошной прямоугольник; None — без проверки бетона.'),
        @('Torsion.WallThickness',0,'mm','Для Strut: толщина полосы; 0 — удвоенный минимальный зазор от оси хомута до бетона; не больше него.'),
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
    $row=Add-GuideParagraph $sheet $row 'Задайте T в отдельной таблице по ID основного сочетания. Пустой ID, пустой T или нулевой T отключают строку. Расчёт НДМ N+Mx+My выполняется независимо: таблицы T не меняют нагрузки, материалы, подбор предельной точки, трещины, устойчивость и вывод NM.'
    $row=Add-GuideParagraph $sheet $row 'Выберите сохранённое состояние НДМ; по умолчанию StrengthState. Постпроверка читает его готовые напряжения без нового решения. Сошедшееся состояние допускается и при выходе за физические диаграммы: его собственные отказы остаются только в блоках НДМ. При отсутствии сохранённых напряжений сумма NMT недоступна, но T проверяется отдельно.'
    $row=Add-GuideParagraph $sheet $row 'Ферма использует одну наружную ось замкнутого хомута, построенную на заданном расстоянии от бетонной поверхности до оси хомута. Нужен достоверный наружный контур; несколько отдельных контуров и многокамерное сечение этой моделью не рассчитываются. Размеры задаются в INPUT, вывод — в OUTPUT; арматурные усилия блока NMT положительны при растяжении.'
    $row=Add-GuideFractionFormula $sheet $row 'q = ' '|T|' '2 A_o'
    $row=Add-GuideFormula $sheet $row 'F_T = q u_o cot θ;  A_s,T,req = F_T / R_s'
    $row=Add-GuideParagraph $sheet $row 'Отбираются центры стержней со стороны бетона, в пределах заданной полосы от оси наружного хомута. У отверстия используются только стержни, ближайшая бетонная граница которых — наружная; равноудалённые и внутренние стержни исключаются. Тяга делится по половинам расстояний вдоль контура до соседних станций; стержни на общей станции делят её пропорционально площади. Нужны минимум три различные станции наружного контура. Вся тяга от T распределяется только на наружные стержни; арматура у отверстия в её восприятии не участвует.'
    $row=Add-GuideFormula $sheet $row 'σ_NMT,i = σ_NM,i + Δσ_T,i;  σ_NMT,i ≤ R_s'
    $row=Add-GuideParagraph $sheet $row 'Сложение является отдельной приближённой проверкой продольной арматуры, а не совместным нелинейным решением. Оно не восстанавливает равновесие и перераспределение НДС и не даёт нового предельного M или N. Горский отдельно оговаривает приближённость раздельного суммирования площадей, стр.24 книги (стр.22 PDF).'
    $row=Add-GuideFractionFormula $sheet $row 'σ_sw = ' 'q s' 'A_sw cot θ'
    $row=Add-GuideFractionFormula $sheet $row 'σ_diag = ' 'q (cot θ + tan θ)' 't_eff'
    $row=Add-GuideFormula $sheet $row 'T_lim = c R_и b² h'
    $row=Add-GuideParagraph $sheet $row 'Подробная таблица T / NMT после устойчивости на Results сначала показывает только T: отбор продольной арматуры, хомут и бетон, затем сумму NMT для определяющего стержня. Внизу отдельно сохраняются только ID сочетания, ID стержня и дополнительное усилие от T. В общей сводке отдельно выводятся статусы и минимальные запасы бетона, продольной арматуры NMT и хомута. Это не множитель предельной нагрузки совместного НДМ. Статусы и причины отказа NM туда не переносятся.'
    $row=Add-GuideParagraph $sheet $row 'Ошибки и предупреждения T показываются в комментарии блока NMT с исходными адресами и цепочками простых ссылок. InputErr означает неверный ввод T; FAIL — превышение проверяемого предела; N/A — не запрошенную проверку или отсутствие сохранённых напряжений для суммы. Ошибка T не останавливает НДМ. При File дополнительные матрицы хранятся в снимке и восстанавливаются вместе с ним; обновление схемы не выполняет расчёт T.'
    return $row
}
