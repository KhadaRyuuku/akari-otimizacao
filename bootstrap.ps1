[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$repoBase = "https://raw.githubusercontent.com/KhadaRyuuku/akari-otimizacao/main"
$dest = Join-Path $env:TEMP "AkariCentralOtimizacao"

if (Test-Path $dest) {
    Remove-Item -Path $dest -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -Path $dest -ItemType Directory -Force | Out-Null

$arquivo1 = "CentralOtimizacao.ps1"
$arquivo2 = "otimizacoes.csv"

function Baixar-Arquivo {
    param($nome)
    $url = "$repoBase/$nome"
    $destino = Join-Path $dest $nome
    try {
        Invoke-WebRequest -Uri $url -OutFile $destino -UseBasicParsing -ErrorAction Stop
        return $true
    }
    catch {
        Write-Host "Invoke-WebRequest falhou para $nome : $($_.Exception.Message)" -ForegroundColor Yellow
        try {
            $client = New-Object System.Net.WebClient
            $client.DownloadFile($url, $destino)
            $client.Dispose()
            return $true
        }
        catch {
            Write-Host "WebClient tambem falhou para $nome : $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
}

$ok1 = Baixar-Arquivo $arquivo1
$ok2 = Baixar-Arquivo $arquivo2

if (-not $ok1 -or -not $ok2) {
    Write-Host ""
    Write-Host "Nao foi possivel baixar os arquivos. Verifique sua conexao, firewall ou antivirus." -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit
}

Write-Host "Download concluido." -ForegroundColor Green
Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $dest $arquivo1)`""
