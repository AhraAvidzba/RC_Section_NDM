# Справочные иллюстрации и независимый выбор на листе «Расчет»; только для сборки.

# Встраивает исходные изображения в координатах утверждённой раскладки.
function Add-SectionIllustrationPictures {
    param([object]$Workbook, [object]$Layout)
    foreach($entry in $Layout.Sheets){
        foreach($item in $entry.Objects){
            if($item.Name -notin @('Circle','Rectangles','PointedSection','HollowRectangle','CurrentSection')){continue}
            $asset=if($item.Name -eq 'CurrentSection'){'PointedSection'}else{[string]$item.Name}
            $path=Join-Path $PSScriptRoot ('SectionIllustrations/'+$asset+'.png')
            $sheet=$Workbook.Worksheets.Item($entry.Name)
            $shape=$sheet.Shapes.AddPicture($path,0,-1,$item.Left,$item.Top,$item.Width,$item.Height)
            $shape.Name=[string]$item.Name
        }
    }
}

# Создаёт четыре отдельных пункта без служебных листов и связей с Geometry.Type.
function Set-SectionIllustrationSelector {
    param([object]$Workbook)
    $cell=$Workbook.Worksheets.Item('Расчет').Range('K2')
    $cell.Validation.Delete()
    $choices=@('Circle','Rectangles','PointedSection','HollowRectangle') -join [string]$Workbook.Application.International(5) # Внешний COM принимает локальный разделитель списка.
    $cell.Validation.Add(3,1,1,$choices)
    $cell.Validation.IgnoreBlank=$false;$cell.Validation.InCellDropdown=$true
    $cell.Validation.ShowError=$true;$cell.Validation.ErrorTitle='Выбор иллюстрации'
    $cell.Validation.ErrorMessage='Выберите одно название сечения из списка.'
    $cell.Validation.InputTitle=''
    $cell.Validation.InputMessage=''
    $cell.Validation.ShowInput=$false
    $cell.Value2='PointedSection'
    $cell.Interior.Color=15135487 # RGB(255,242,230): доступный для ввода выбор.
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
        $media=@{};$found=@{}
        foreach($entry in @($zip.Entries | Where-Object {$_.FullName -match '^xl/drawings/drawing\d+\.xml$'})){
            $reader=New-Object IO.StreamReader($entry.Open());try{$drawing=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
            $relsPath='xl/drawings/_rels/'+[IO.Path]::GetFileName($entry.FullName)+'.rels'
            $reader=New-Object IO.StreamReader($zip.GetEntry($relsPath).Open());try{$rels=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
            foreach($pic in $drawing.SelectNodes('//*[local-name()="pic"]')){
                $name=$pic.SelectSingleNode('./*[local-name()="nvPicPr"]/*[local-name()="cNvPr"]').GetAttribute('name')
                if($name -notin @('Circle','Rectangles','PointedSection','HollowRectangle','CurrentSection')){continue}
                $asset=if($name -eq 'CurrentSection'){'PointedSection'}else{$name};$found[$asset]=$true
                $id=$pic.SelectSingleNode('./*[local-name()="blipFill"]/*[local-name()="blip"]').GetAttribute('embed','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $target=($rels.Relationships.Relationship | Where-Object Id -eq $id).Target
                $part=([Uri]::new([Uri]('http://package/'+$entry.FullName),[string]$target)).AbsolutePath.TrimStart('/')
                if($media.ContainsKey($part) -and $media[$part] -ne $asset){throw 'Разные исходные картинки ссылаются на одну часть файла.'}
                $media[$part]=$asset
            }
        }
        if($found.Count -ne 4){throw 'В сборке отсутствует одна из четырёх справочных картинок.'}
        foreach($part in $media.Keys){
            $bytes=[IO.File]::ReadAllBytes((Join-Path $PSScriptRoot ('SectionIllustrations/'+$media[$part]+'.png')))
            $zip.GetEntry($part).Delete();$stream=$zip.CreateEntry($part).Open();try{$stream.Write($bytes,0,$bytes.Length)}finally{$stream.Dispose()}
        }
        $entry=$zip.GetEntry('xl/workbook.xml');$reader=New-Object IO.StreamReader($entry.Open());try{$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
        $xml.workbook.workbookPr.SetAttribute('autoCompressPictures','0')
        $entry.Delete();$writer=New-Object IO.StreamWriter($zip.CreateEntry('xl/workbook.xml').Open(),[Text.UTF8Encoding]::new($false))
        try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
    } finally {$zip.Dispose()}
}
