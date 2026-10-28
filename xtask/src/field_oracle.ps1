# Runs field codes in real Word through COM. Driven by `cargo xtask field-oracle`, which defines
# $codeHex (hex of the UTF-8 codes, one per line, newline-joined) before this text and reads one
# `index<TAB>hex-of-result` line per code from stdout. Only temporary documents are used: alerts
# are off, documents close without saving, and the Word instance this script started is quit in cleanup.
$ErrorActionPreference = 'Stop'
function FromHex([string]$h) {
    if ($h.Length -eq 0) { return '' }
    $b = New-Object byte[] ($h.Length / 2)
    for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = [Convert]::ToByte($h.Substring(2 * $i, 2), 16) }
    [Text.Encoding]::UTF8.GetString($b)
}
function ToHex([string]$s) { -join ([Text.Encoding]::UTF8.GetBytes($s) | ForEach-Object { $_.ToString('x2') }) }

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $codes = (FromHex $codeHex) -split "`n"
    for ($i = 0; $i -lt $codes.Count; $i++) {
        try {
            $end = $doc.Content.End - 1
            $f = $doc.Fields.Add($doc.Range($end, $end), -1, $codes[$i], $false)
            [void]$f.Update()
            $res = $f.Result.Text
        } catch { $res = 'COM error: ' + $_.Exception.Message }
        $doc.Content.InsertParagraphAfter()
        '{0}{1}{2}' -f $i, "`t", (ToHex (($res -replace "[\r\n\a]+", ' ').TrimEnd()))
    }
    $doc.Close(0)
} finally {
    if ($word) { try { $word.Quit(0) } catch {} }
}
