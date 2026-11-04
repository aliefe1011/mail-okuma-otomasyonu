$ErrorActionPreference = 'Stop'
# Run locally on the Windows computer that has classic Outlook and Excel.
$root = $PSScriptRoot
$python = Get-Command py -ErrorAction SilentlyContinue
if (-not $python) { throw 'Python 3.10+ ve Windows py launcher kurun, sonra yeniden calistirin.' }
& py -3 -c "import sys; assert sys.version_info >= (3,10), 'Python 3.10+ required'"
if ($LASTEXITCODE -ne 0) { throw 'Python surumu uygun degil.' }
$secureKey = Read-Host 'Yeni Gemini API anahtari (ekranda gorunmez)' -AsSecureString
$keyPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
try {
    $key = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($keyPtr)
    if ([string]::IsNullOrWhiteSpace($key)) { throw 'API anahtari bos olamaz.' }
    $env:GEMINI_API_KEY = $key
    $env:GEMINI_BRIDGE_DIR = $root
    if (-not $env:GEMINI_MODEL) { $env:GEMINI_MODEL = 'gemini-2.5-flash' }
    & py -3 (Join-Path $root 'gemini\check_connection.py')
    if ($LASTEXITCODE -ne 0) { throw 'Gemini testi basarisiz; anahtar, kota, model ve ag erisimini kontrol edin.' }
    [Environment]::SetEnvironmentVariable('GEMINI_API_KEY', $key, 'User')
    [Environment]::SetEnvironmentVariable('GEMINI_BRIDGE_DIR', $root, 'User')
    [Environment]::SetEnvironmentVariable('GEMINI_MODEL', $env:GEMINI_MODEL, 'User')
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($keyPtr)
    $key = $null
    $secureKey = $null
    Remove-Item Env:GEMINI_API_KEY -ErrorAction SilentlyContinue
}
Write-Host 'Kurulum ve canli Gemini testi tamam. Excel tamamen kapali olmali; yeniden acin.'
Write-Host 'README adimlariyla BAS modullerini bir kez ice aktarin. Sonra GelenYanitlariOku makrosunu kullanin.'
