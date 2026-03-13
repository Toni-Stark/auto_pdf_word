$WPS_HOME = 'F:\wps\WPS Office\12.1.0.25225\office6'
$PDF_DIR  = 'F:\pdf_word\pdf'

$wpsExe = Join-Path $WPS_HOME 'wps.exe'
$wpsDll = Join-Path $WPS_HOME 'addons\kappessframework\kappessframework.dll'

$pdfs = Get-ChildItem -Path $PDF_DIR -Filter *.pdf -File | Select-Object -ExpandProperty FullName


if (-not $pdfs -or $pdfs.Count -eq 0) {
    Write-Host "No PDF files found:"
    Write-Host $PDF_DIR
    exit
}

$fileArg = '/file=' + ($pdfs -join '|')

& $wpsExe Run /InstanceId=kpdf2wordv2 $wpsDll /appId=kpdf2wordv2 /action=ConvertToWord $fileArg
