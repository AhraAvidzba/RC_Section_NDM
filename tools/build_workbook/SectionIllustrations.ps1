# Справочные иллюстрации и независимый выбор на листе «Расчет»; только для сборки.

# Читает явное перечисление доступных иллюстраций; вспомогательные рисунки в список не включаются.
function Read-SectionIllustrationCatalog {
    Get-Content -LiteralPath (Join-Path $PSScriptRoot 'SectionIllustrations/Illustrations.json') -Raw -Encoding UTF8 | ConvertFrom-Json
}

# Объединяет выбираемые и общие PNG только для сборки; список K2 читает исключительно Names.
function Read-SectionIllustrationAssets {
    $catalog=Read-SectionIllustrationCatalog
    foreach($name in $catalog.Names){[pscustomobject]@{Name=[string]$name;File=([string]$name+'.png')}}
    foreach($asset in $catalog.Additional){$asset}
}

# Встраивает PNG перечисленных типов; новые имена не требуют ветвей в VBA.
function Add-SectionIllustrationPictures {
    param([object]$Workbook, [object]$Layout)
    Add-Type -AssemblyName System.Drawing
    $catalog=Read-SectionIllustrationCatalog
    $source=$Workbook.Worksheets.Item('Сечения');$target=$Workbook.Worksheets.Item('Расчет')
    $sourceLayout=@($Layout.Sheets | Where-Object Name -eq 'Сечения')[0]
    $targetLayout=@($Layout.Sheets | Where-Object Name -eq 'Расчет')[0]
    $assets=@(Read-SectionIllustrationAssets)
    $known=@($sourceLayout.Objects | Where-Object {$_.Name -in $assets.Name})
    $nextTop=($known | ForEach-Object {$_.Top+$_.Height} | Measure-Object -Maximum).Maximum+24
    foreach($asset in $assets){
        $name=[string]$asset.Name
        $position=@($known | Where-Object Name -eq $name)
        if($position.Count){$item=$position[0]}else{
            $item=[pscustomobject]@{Left=$known[0].Left;Top=$nextTop;Width=$known[0].Width;Height=$known[0].Height}
            $nextTop+=$item.Height+24 # Дополнительный тип без своей записи раскладки размещается под остальными рисунками.
        }
        $path=Join-Path $PSScriptRoot ('SectionIllustrations/'+$asset.File)
        $shape=$source.Shapes.AddPicture($path,0,-1,$item.Left,$item.Top,$item.Width,$item.Height);$shape.Name=[string]$name
        $image=[Drawing.Image]::FromFile($path)
        try{$aspect=[double]$image.Width/[double]$image.Height}finally{$image.Dispose()}
        $fit=[Math]::Min([double]$item.Height,[double]$item.Width/$aspect)
        $crop=$shape.PictureFormat.Crop;$crop.PictureWidth=$fit*$aspect;$crop.PictureHeight=$fit;$crop.PictureOffsetX=0;$crop.PictureOffsetY=0
    }
    $item=@($targetLayout.Objects | Where-Object Name -eq 'CurrentSection')[0]
    $path=Join-Path $PSScriptRoot ('SectionIllustrations/'+$catalog.Default+'.png')
    $picture=$target.Shapes.AddPicture($path,0,-1,$item.Left,$item.Top,$item.Width,$item.Height)
    $picture.Name='CurrentSection';$picture.LockAspectRatio=0
    $image=[Drawing.Image]::FromFile($path)
    try{$aspect=[double]$image.Width/[double]$image.Height}finally{$image.Dispose()}
    $fit=[Math]::Min([double]$item.Height,[double]$item.Width/$aspect)
    $crop=$picture.PictureFormat.Crop;$crop.PictureWidth=$fit*$aspect;$crop.PictureHeight=$fit;$crop.PictureOffsetX=0;$crop.PictureOffsetY=0
}

# Сохраняет обычный список прямо в проверке данных K2; служебных ячеек и имён нет.
function Set-SectionIllustrationSelector {
    param([object]$Workbook)
    $catalog=Read-SectionIllustrationCatalog
    $cell=$Workbook.Worksheets.Item('Расчет').Range('K2')
    $cell.Validation.Delete()
    $choices=@($catalog.Names) -join [string]$Workbook.Application.International(5) # Внешний COM использует разделитель своей локали.
    $cell.Validation.Add(3,1,1,$choices)
    $cell.Validation.IgnoreBlank=$false;$cell.Validation.InCellDropdown=$true
    $cell.Validation.ShowError=$true;$cell.Validation.ErrorTitle='Выбор иллюстрации'
    $cell.Validation.ErrorMessage='Выберите одно название сечения из списка.'
    $cell.Validation.InputTitle='';$cell.Validation.InputMessage='';$cell.Validation.ShowInput=$false
    $cell.Value2=[string]$catalog.Default
    $cell.Interior.Color=15135487
    $cell.HorizontalAlignment=-4108;$cell.VerticalAlignment=-4108;$cell.WrapText=$false
    foreach($edge in @(7,8,9,10)){$cell.Borders.Item($edge).LineStyle=1;$cell.Borders.Item($edge).Weight=2;$cell.Borders.Item($edge).Color=12632256}
}

# После закрытия Excel сохраняет исходные PNG без автоматического уменьшения разрешения.
# Связи рисунков определяются по именам объектов; VBA и данные листов не меняются.
function Save-SectionIllustrationImageBytes {
    param([string]$WorkbookPath)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::Open($WorkbookPath,[IO.Compression.ZipArchiveMode]::Update)
    try {
        $media=@{};$found=@{};$assets=@(Read-SectionIllustrationAssets);$byName=@{};foreach($asset in $assets){$byName[[string]$asset.Name]=[string]$asset.File}
        foreach($entry in @($zip.Entries | Where-Object {$_.FullName -match '^xl/drawings/drawing\d+\.xml$'})){
            $reader=New-Object IO.StreamReader($entry.Open());try{$drawing=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
            if($drawing.SelectNodes('//*[local-name()="pic"]').Count-eq0){continue} # Чертёж только с кнопками или пустой схемой может не иметь файла связей.
            $relsPath='xl/drawings/_rels/'+[IO.Path]::GetFileName($entry.FullName)+'.rels'
            $reader=New-Object IO.StreamReader($zip.GetEntry($relsPath).Open());try{$rels=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
            foreach($pic in $drawing.SelectNodes('//*[local-name()="pic"]')){
                $name=$pic.SelectSingleNode('./*[local-name()="nvPicPr"]/*[local-name()="cNvPr"]').GetAttribute('name')
                $asset=$name
                if(-not $byName.ContainsKey($asset)){continue}
                $found[$asset]=$true
                $id=$pic.SelectSingleNode('./*[local-name()="blipFill"]/*[local-name()="blip"]').GetAttribute('embed','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $target=($rels.Relationships.Relationship | Where-Object Id -eq $id).Target
                $part=([Uri]::new([Uri]('http://package/'+$entry.FullName),[string]$target)).AbsolutePath.TrimStart('/')
                if($media.ContainsKey($part) -and $media[$part] -ne $asset){throw 'Разные исходные картинки ссылаются на одну часть файла.'}
                $media[$part]=$asset
            }
        }
        if($found.Count -ne $assets.Count){throw 'В сборке отсутствует иллюстрация из перечисления типов.'}
        foreach($part in $media.Keys){
            $bytes=[IO.File]::ReadAllBytes((Join-Path $PSScriptRoot ('SectionIllustrations/'+$byName[$media[$part]])))
            $zip.GetEntry($part).Delete();$stream=$zip.CreateEntry($part).Open();try{$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
        }
        $entry=$zip.GetEntry('xl/workbook.xml');$reader=New-Object IO.StreamReader($entry.Open());try{$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
        $xml.workbook.workbookPr.SetAttribute('autoCompressPictures','0')
        $entry.Delete();$writer=New-Object IO.StreamWriter($zip.CreateEntry('xl/workbook.xml').Open(),[Text.UTF8Encoding]::new($false))
        try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
    } finally {$zip.Dispose()}
}
