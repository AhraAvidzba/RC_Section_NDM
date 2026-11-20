# Подготавливает утверждённое оформление к снятию автономного шаблона сборки.
# Рабочие расчёты этот файл не вызывают и формат входных листов не меняют.

# Находит ссылки внутри листа, исключая текст формулы и адреса других листов.
function Get-CalculationLocalReferences([string]$Formula) {
    $text=[regex]::Replace($Formula,'(?:''[^'']+''|[\w]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?','')
    $text=[regex]::Replace($text,'"(?:[^"]|"")*"','').Replace('$','')
    return @([regex]::Matches($text,'(?<![\w.])[A-Z]{1,3}\d+\b') | ForEach-Object Value | Sort-Object -Unique)
}

# Заменяет только отсутствие выбранного сочетания прочерком; числовой ноль сохраняется.
# Непредвиденные #REF!/#VALUE! не скрываются общим IFERROR.
function Set-CalculationMissingCombinationFormulas([object]$Sheet) {
    $formulas=@{};$roots=@{};$selectors=@{};$changed=@()
    foreach($cell in $Sheet.Range('A1:BQ59').SpecialCells(-4123)){
        $address=$cell.Address($false,$false);$formula=[string]$cell.Formula
        $formulas[$address]=$formula
        if($formula -match '^=(?:Results!\$?[G-R]\$?12|IF\(Results!\$?[G-R]\$?12=)'){
            $source=[regex]::Match($formula,'Results!\$?[G-R]\$?12').Value.Replace('$','')
            $roots[$address]=@($address);$selectors[$address]=$source
        }
    }
    for($pass=0;$pass -lt $formulas.Count;$pass++){
        $updated=$false
        foreach($address in @($formulas.Keys)){
            if($selectors.ContainsKey($address)){continue}
            $found=@(foreach($reference in (Get-CalculationLocalReferences $formulas[$address])){
                if($roots.ContainsKey($reference)){$roots[$reference]}
            }) | Sort-Object -Unique
            if(($found -join ',') -ne (@($roots[$address]) -join ',')){$roots[$address]=@($found);$updated=$true}
        }
        if(-not $updated){break}
    }
    foreach($address in $formulas.Keys){
        $formula=$formulas[$address];$replacement=$null
        if($selectors.ContainsKey($address)){
            $source=$selectors[$address]
            $replacement='=IF(OR('+ $source+'="",'+$source+'="-"),"-",'+$source+')'
        }elseif($roots.ContainsKey($address) -and @($roots[$address]).Count){
            $sources=@($roots[$address] | ForEach-Object {$selectors[$_]} | Sort-Object -Unique)
            if($sources.Count -eq 1){
                $reference=[string]@($roots[$address])[0]
                $reference=[regex]::Replace($reference,'^([A-Z]+)(\d+)$','$$$1$$$2')
                $body=$formula.Substring(1)
                $prefix='=IF(OR('+$reference+'="",'+$reference+'="-"),"-",'
                if(-not $formula.StartsWith($prefix)){$replacement=$prefix+'IFNA('+$body+',"-"))'}
            }
        }
        if($replacement -and $replacement -cne $formula){$Sheet.Range($address).Formula=$replacement;$changed+=$address}
    }
    return $changed
}

# Условная заливка зависит от INPUT-знака N и обновляется без событий VBA.
function Set-CalculationAxialTensionFormat([object]$Workbook) {
    $sheet=$Workbook.Worksheets.Item('Расчет');$rules=@()
    $sign='INDEX(rngSignConventionSettings,MATCH("+N",INDEX(rngSignConventionSettings,,1),0),2)'
    foreach($column in @('Q','T','Y')){
        $range=$sheet.Range($column+'44:'+$column+'58');$cell=$column+'44'
        $range.FormatConditions.Delete()
        $formula='=AND(ISNUMBER('+$cell+'),'+$cell+'*IF('+$sign+'="Tension",1,-1)>0)'
        $scratch=$sheet.Range('BQ60');$previous=$scratch.Formula
        try{
            $scratch.Formula=$formula
            $local=[string]$scratch.FormulaLocal # Внешний COM принимает формулу условного форматирования в языке и разделителях Excel.
            $condition=$range.FormatConditions.Add(2,0,$local)
        }finally{$scratch.Formula=$previous}
        $condition.Interior.Color=6737151 # RGB(255,217,102): всегда растяжение, независимо от принятого знака N.
        $condition.StopIfTrue=$false;$rules+=,[ordered]@{Range=$range.Address();Formula=$formula}
    }
    return $rules
}

# Меняет цвет только существующих границ в области печати, сохраняя их тип и толщину.
function Set-CalculationPrintBorderColor([object]$Sheet) {
    $changed=0
    $printNames=@($Sheet.Names | Where-Object {$_.Name -match 'Print_Area$|Область_печати$'})
    if($printNames.Count -ne 1){throw 'Не определена единственная область печати листа «Расчет».'}
    $printRange=$printNames[0].RefersToRange
    foreach($cell in $printRange){
        foreach($index in @(5,6,7,8,9,10)){
            $border=$cell.Borders.Item($index)
            if($border.LineStyle -ne -4142 -and [long]$border.Color -ne 8421504){$border.Color=8421504;$changed++}
        }
    }
    return $changed
}

# Прочерк означает отсутствие проверки и сохраняет зелёную базовую заливку статуса.
function Set-CalculationStatusDashFormat([object]$Sheet) {
    $changed=@()
    foreach($condition in @($Sheet.Range('A1:BD59').FormatConditions)){
        $text=([string]$condition.Formula1).TrimStart('=')
        $statusRule=($condition.Type -eq 2 -and $text.Contains('"OK"') -and $text.Contains('"-"'))
        if($condition.Type -eq 1){$statusRule=($condition.Operator -eq 4 -and $text -eq '"OK"')}
        if(-not $statusRule){continue}
        $cell=$condition.AppliesTo.Cells.Item(1,1).Address($true,$true) # Modify трактует относительные ссылки от активной ячейки, поэтому одиночный статус закреплён абсолютно.
        $formula='=AND('+$cell+'<>"OK",'+$cell+'<>"-")'
        $scratch=$Sheet.Range('BQ60');$previous=$scratch.Formula
        try{
            $scratch.Formula=$formula
            $condition.Modify(2,0,[string]$scratch.FormulaLocal)
        }finally{$scratch.Formula=$previous}
        $changed+=$cell
    }
    return $changed
}

# Округляет лишь представление нагрузок; ID, деформации и напряжения сохраняют собственную точность.
function Set-CalculationLoadPrecision([object]$Sheet) {
    foreach($address in @('Q44:V58','Y44:AA58','AS8:AS22','AI7:AK7','AH27:AJ27','AH39:AJ39','AY18:BB18','AH48:AJ49','AH53:AJ54')){
        $Sheet.Range($address).NumberFormat='0.0;-0.0;0.0;@'
    }
}

# Настраивает сохранённую книгу перед захватом шаблона; данные пользователя не округляет.
function Set-CalculationDisplay([object]$Workbook) {
    $sheet=$Workbook.Worksheets.Item('Расчет')
    $formulas=@(Set-CalculationMissingCombinationFormulas $sheet)
    Set-CalculationLoadPrecision $sheet
    $rules=@(Set-CalculationAxialTensionFormat $Workbook)
    $statuses=@(Set-CalculationStatusDashFormat $sheet)
    $borders=Set-CalculationPrintBorderColor $sheet
    $Workbook.Worksheets.Item('Results').Range('AL8').Value2="Нагрузки с учётом продольного`nизгиба (если применимо)"
    return [ordered]@{ChangedFormulas=$formulas;TensionRules=$rules;StatusRules=$statuses;ChangedBorderEdges=$borders;BorderRGB=@(128,128,128);NumberDecimals=1}
}
