# Script de sincronización instantánea Overleaf -> Local -> GitHub
$projectId = "6acd1f92fb154e211b0d0ba7"
$cookieFile = "$PSScriptRoot\.olauth"

if (-not (Test-Path $cookieFile)) {
    Write-Host "❌ No se encontró el archivo .olauth con la cookie overleaf_session2." -ForegroundColor Red
    exit 1
}

$cookieVal = (Get-Content $cookieFile -Raw).Trim()

Write-Host "📥 Descargando última versión desde Overleaf ($projectId)..." -ForegroundColor Cyan
$zipFile = "$env:TEMP\overleaf_sync.zip"

curl.exe -sSL "https://www.overleaf.com/project/$projectId/download/zip" `
    -H "Cookie: overleaf_session2=$cookieVal" `
    -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64)" `
    -o $zipFile

if (-not (Test-Path $zipFile) -or (Get-Item $zipFile).Length -lt 500) {
    Write-Host "❌ Error: La descarga falló. Comprueba que la cookie no haya caducado." -ForegroundColor Red
    exit 1
}

Write-Host "📦 Extrayendo archivos en local..." -ForegroundColor Cyan
tar.exe -xf $zipFile -C $PSScriptRoot
Remove-Item $zipFile -Force

Write-Host "🐙 Comprobando cambios en Git..." -ForegroundColor Cyan
git add .
git diff --staged --quiet
if ($LASTEXITCODE -ne 0) {
    git commit -m "chore: sync from Overleaf"
    git push origin main
    Write-Host "🚀 ¡Sincronizado con Overleaf y subido a GitHub!" -ForegroundColor Green
} else {
    Write-Host "✅ Todo está al día con Overleaf." -ForegroundColor Green
}
