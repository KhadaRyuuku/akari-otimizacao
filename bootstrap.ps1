<#
    Bootstrap da Central de Otimizacao de PC (Akari)
    --------------------------------------------------
    Este e o arquivo que fica hospedado online. Quem roda o comando de
    uma linha (estilo Chris Titus) baixa este script primeiro; ele baixa
    os arquivos reais da Central pra uma pasta temporaria e abre.

    ANTES DE USAR: troque SEU_USUARIO e SEU_REPOSITORIO abaixo pelos
    dados do seu repositorio no GitHub.
#>

$repoBase = "https://raw.githubusercontent.com/KhadaRyuuku/akari-otimizacao/main"
$destDir = Join-Path $env:TEMP "AkariCentralOtimizacao"

Write-Host "==========================================="
Write-Host "  Central de Otimizacao de PC - Akari"
Write-Host "==========================================="
Write-Host "Baixando arquivos..."

if (-not (Test-Path $destDir)) {
    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
}

$arquivos = @("CentralOtimizacao.ps1", "otimizacoes.csv")

foreach ($arquivo in $arquivos) {
    $url = "$repoBase/$arquivo"
    $destino = Join-Path $destDir $arquivo
    try {
        Invoke-WebRequest -Uri $url -OutFile $destino -UseBasicParsing
        Write-Host "  OK: $arquivo"
    }
    catch {
        Write-Host "  ERRO ao baixar $arquivo : $($_.Exception.Message)"
        Write-Host "Verifique se a URL do repositorio esta correta no bootstrap.ps1."
        exit
    }
}

# ---------- Opcional: pasta de Ferramentas Externas (NVCleanstall, Nvidia Profile Inspector) ----------
# Se voce hospedar um "FerramentasExternas.zip" no repositorio, descomente as linhas abaixo:
#
# $zipUrl = "$repoBase/FerramentasExternas.zip"
# $zipDestino = Join-Path $destDir "FerramentasExternas.zip"
# try {
#     Invoke-WebRequest -Uri $zipUrl -OutFile $zipDestino -UseBasicParsing
#     Expand-Archive -Path $zipDestino -DestinationPath $destDir -Force
#     Remove-Item $zipDestino -Force
#     Write-Host "  OK: FerramentasExternas"
# }
# catch {
#     Write-Host "  AVISO: nao foi possivel baixar as Ferramentas Externas (opcional)."
# }

Write-Host "Download concluido. Abrindo a Central..."

$scriptPath = Join-Path $destDir "CentralOtimizacao.ps1"
Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
