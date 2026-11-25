param([string]$Docx, [string]$Pdf)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($Docx, $false, $true, $false)
    $document.Repaginate()
    $document.ExportAsFixedFormat($Pdf, 17)
} finally {
    $discard = [object]0
    if ($document) { $document.Close([ref]$discard) }
    $word.Quit([ref]$discard)
}

