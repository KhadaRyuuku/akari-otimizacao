$repoBase = "https://raw.githubusercontent.com/KhadaRyuuku/akari-otimizacao/main"
$dest = Join-Path $env:TEMP "AkariCentralOtimizacao"
New-Item -Path $dest -ItemType Directory -Force | Out-Null
$cacheBust = Get-Random
"CentralOtimizacao.ps1", "otimizacoes.csv" | ForEach-Object {
    Invoke-WebRequest -Uri "$repoBase/$_?v=$cacheBust" -OutFile (Join-Path $dest $_) -UseBasicParsing
}
Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $dest 'CentralOtimizacao.ps1')`""
