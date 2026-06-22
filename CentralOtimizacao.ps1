<#
    Central de Otimizacao de PC - v3 (Dark Mode)
    ----------------------------------------------
    Ferramenta para uso por um tecnico/amigo otimizando, COM CONSENTIMENTO,
    o computador de um cliente/amigo (localmente ou via acesso remoto).

    A lista de otimizacoes fica no arquivo "otimizacoes.csv" (mesma pasta
    deste script). Para adicionar, remover ou editar nome/risco de qualquer
    item, abra o otimizacoes.csv no Excel ou Notepad.

    Marque as caixas desejadas e clique em "Aplicar Selecionados".
    Um ponto de restauracao do Windows e criado automaticamente antes de
    qualquer alteracao.
#>

try { $Host.UI.RawUI.WindowTitle = "Akari - Central de Otimizacao" } catch {}

Write-Host ""
Write-Host " ###   #   #   ###   ####   ###  " -ForegroundColor Red
Write-Host "#   #  #  #   #   #  #   #   #   " -ForegroundColor Red
Write-Host "#####  ###    #####  ####    #   " -ForegroundColor Red
Write-Host "#   #  #  #   #   #  #  #    #   " -ForegroundColor Red
Write-Host "#   #  #   #  #   #  #   #  ###  " -ForegroundColor Red
Write-Host ""
Write-Host "===== Akari Tech =====" -ForegroundColor White
Write-Host "===== Central de Otimizacao de PC =====" -ForegroundColor DarkGray
Write-Host ""

# ---------- Autoelevacao (precisa rodar como Administrador) ----------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------- P/Invoke para deixar a barra de titulo nativa tambem no escuro ----------
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DarkModeHelper {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@

$logPath = "$env:USERPROFILE\Desktop\OtimizacaoPC_log.txt"
$script:needsRestart = $false
$fontFamily = "Corbel"

$logoBase64 = "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" +
"SUVORK5CYII="

$logoImage = $null
try {
    $logoBytes = [System.Convert]::FromBase64String($logoBase64)
    $logoStream = New-Object System.IO.MemoryStream(, $logoBytes)
    $logoImage = [System.Drawing.Image]::FromStream($logoStream)
}
catch {
    $logoImage = $null
}

# ---------- Paleta de cores (dark mode profundo - sem tons de azul) ----------
$colorBg = [System.Drawing.Color]::FromArgb(8, 8, 8)
$colorSidebar = [System.Drawing.Color]::FromArgb(12, 12, 12)
$colorPanel = [System.Drawing.Color]::FromArgb(16, 16, 16)
$colorAccent = [System.Drawing.Color]::FromArgb(214, 40, 70)
$colorAccentDim = [System.Drawing.Color]::FromArgb(90, 35, 42)
$colorText = [System.Drawing.Color]::FromArgb(240, 240, 240)
$colorTextDim = [System.Drawing.Color]::FromArgb(145, 145, 145)
$colorBorder = [System.Drawing.Color]::FromArgb(40, 40, 40)

function Get-RiskColor {
    param([string]$risco)
    switch ($risco) {
        "Baixo" { return [System.Drawing.Color]::FromArgb(92, 200, 110) }
        "Medio" { return [System.Drawing.Color]::FromArgb(255, 181, 71) }
        "Alto" { return [System.Drawing.Color]::FromArgb(255, 99, 71) }
        "Nenhum" { return [System.Drawing.Color]::FromArgb(180, 175, 165) }
        default { return $colorTextDim }
    }
}

function Adjust-Color {
    param([System.Drawing.Color]$color, [int]$amount)
    $r = [Math]::Min(255, [Math]::Max(0, $color.R + $amount))
    $g = [Math]::Min(255, [Math]::Max(0, $color.G + $amount))
    $b = [Math]::Min(255, [Math]::Max(0, $color.B + $amount))
    return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

function Add-ClickEffect {
    param($btn, [System.Drawing.Color]$baseColor, [System.Drawing.Color]$hoverColor, [System.Drawing.Color]$pressColor)
    $btn.Add_MouseEnter({ $this.BackColor = $hoverColor }.GetNewClosure())
    $btn.Add_MouseLeave({ $this.BackColor = $baseColor }.GetNewClosure())
    $btn.Add_MouseDown({ $this.BackColor = $pressColor }.GetNewClosure())
    $btn.Add_MouseUp({ $this.BackColor = $hoverColor }.GetNewClosure())
}

function Set-RoundedRegion {
    param($ctrl, [int]$radius)
    $w = $ctrl.Width
    $h = $ctrl.Height
    $d = $radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($w - $d, 0, $d, $d, 270, 90)
    $path.AddArc($w - $d, $h - $d, $d, $d, 0, 90)
    $path.AddArc(0, $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $ctrl.Region = New-Object System.Drawing.Region($path)
}

function Add-GradientBackground {
    param($ctrl, [System.Drawing.Color]$colorTop, [System.Drawing.Color]$colorBottom)
    $ctrl.Add_Paint({
            param($s, $e)
            $rect = $s.ClientRectangle
            if ($rect.Width -gt 0 -and $rect.Height -gt 0) {
                $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $colorTop, $colorBottom, 90)
                $e.Graphics.FillRectangle($brush, $rect)
                $brush.Dispose()
            }
        }.GetNewClosure())
}

# ---------- Carregar lista de otimizacoes do CSV externo ----------
$csvPath = Join-Path $PSScriptRoot "otimizacoes.csv"

if (-not (Test-Path $csvPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Arquivo 'otimizacoes.csv' nao encontrado na mesma pasta deste script.`n`nCaminho esperado: $csvPath",
        "Arquivo nao encontrado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit
}

try {
    $rows = Import-Csv -Path $csvPath -Encoding UTF8
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Erro ao ler otimizacoes.csv: $($_.Exception.Message)", "Erro de leitura", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit
}

$optimizations = @()
foreach ($row in $rows) {
    $sufixo = ""
    if ($row.RequerReinicio -eq "Sim") { $sufixo = "   (requer reiniciar)" }
    $optimizations += @{
        Cat     = $row.Categoria
        Nome    = $row.Nome
        Risco   = $row.Risco
        Label   = "$($row.Nome)$sufixo"
        Restart = ($row.RequerReinicio -eq "Sim")
        Action  = [scriptblock]::Create($row.Comando)
    }
}

if ($optimizations.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("O arquivo otimizacoes.csv esta vazio ou nao tem linhas validas.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    exit
}

$categories = $optimizations.Cat | Select-Object -Unique

# ---------- Janela principal ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Central de Otimizacao de PC"
$form.Size = New-Object System.Drawing.Size(980, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = $colorBg
$form.Font = New-Object System.Drawing.Font($fontFamily, 9)

$form.Add_Shown({
        $val = 1
        [DarkModeHelper]::DwmSetWindowAttribute($form.Handle, 20, [ref]$val, 4) | Out-Null
    })

Add-GradientBackground $form ([System.Drawing.Color]::FromArgb(14, 14, 14)) ([System.Drawing.Color]::FromArgb(4, 4, 4))

# ---------- Cabecalho ----------
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.BackColor = $colorSidebar
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(980, 64)
$form.Controls.Add($headerPanel)

$headerPanel.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen($colorAccent, 2)
        $e.Graphics.DrawLine($pen, 0, ($s.Height - 1), $s.Width, ($s.Height - 1))
        $pen.Dispose()
    }.GetNewClosure())

if ($logoImage) {
    $logoBox = New-Object System.Windows.Forms.PictureBox
    $logoBox.Image = $logoImage
    $logoBox.SizeMode = "Zoom"
    $logoBox.Size = New-Object System.Drawing.Size(48, 48)
    $logoBox.Location = New-Object System.Drawing.Point(14, 8)
    $headerPanel.Controls.Add($logoBox)
    $tituloX = 72
}
else {
    $tituloX = 24
}

$titleLbl = New-Object System.Windows.Forms.Label
$titleLbl.Text = "CENTRAL DE OTIMIZACAO"
$titleLbl.ForeColor = $colorText
$titleLbl.Font = New-Object System.Drawing.Font($fontFamily, 15, [System.Drawing.FontStyle]::Bold)
$titleLbl.Location = New-Object System.Drawing.Point($tituloX, 8)
$titleLbl.AutoSize = $true
$headerPanel.Controls.Add($titleLbl)

$subtitleLbl = New-Object System.Windows.Forms.Label
$subtitleLbl.Text = "Marque as otimizacoes desejadas e clique em Aplicar Selecionados"
$subtitleLbl.ForeColor = $colorTextDim
$subtitleLbl.Font = New-Object System.Drawing.Font($fontFamily, 9)
$subtitleLbl.Location = New-Object System.Drawing.Point(($tituloX + 2), 38)
$subtitleLbl.AutoSize = $true
$headerPanel.Controls.Add($subtitleLbl)

# Legenda de risco (quadradinho colorido em vez de caractere especial - evita bug de encoding)
$legendaX = 480
foreach ($item in @(@("Baixo", "Baixo"), @("Medio", "Medio"), @("Alto", "Alto"), @("Nenhum", "Info"))) {
    $swatch = New-Object System.Windows.Forms.Panel
    $swatch.BackColor = Get-RiskColor $item[0]
    $swatch.Size = New-Object System.Drawing.Size(12, 12)
    $swatch.Location = New-Object System.Drawing.Point($legendaX, 26)
    $headerPanel.Controls.Add($swatch)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $item[1]
    $lbl.ForeColor = $colorTextDim
    $lbl.Location = New-Object System.Drawing.Point(($legendaX + 18), 24)
    $lbl.AutoSize = $true
    $headerPanel.Controls.Add($lbl)
    $legendaX += 68
}

# ---------- Botao de Discord (com efeito de flash ao clicar) ----------
$discordBtn = New-Object System.Windows.Forms.Button
$discordBtn.Text = "Discord"
$discordBtn.Size = New-Object System.Drawing.Size(110, 36)
$discordBtn.Location = New-Object System.Drawing.Point(856, 14)
$discordBtn.FlatStyle = "Flat"
$discordBtn.FlatAppearance.BorderSize = 1
$discordBtn.FlatAppearance.BorderColor = $colorAccent
$discordBtn.BackColor = $colorSidebar
$discordBtn.ForeColor = $colorText
$discordBtn.Font = New-Object System.Drawing.Font($fontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
$discordBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
Set-RoundedRegion $discordBtn 18
$headerPanel.Controls.Add($discordBtn)

$discordFlashTimer = New-Object System.Windows.Forms.Timer
$discordFlashTimer.Interval = 35
$script:discordFlashStep = 0

$discordFlashTimer.Add_Tick({
        $script:discordFlashStep++
        if ($script:discordFlashStep -ge 6) {
            $discordFlashTimer.Stop()
            $discordBtn.BackColor = $colorAccent
            $discordBtn.ForeColor = [System.Drawing.Color]::White
        }
        else {
            $discordBtn.BackColor = Adjust-Color $colorAccent (40 - ($script:discordFlashStep * 8))
        }
    })

$discordBtn.Add_Click({
        $script:discordFlashStep = 0
        $discordBtn.BackColor = [System.Drawing.Color]::White
        $discordBtn.ForeColor = $colorAccent
        $discordFlashTimer.Start()
        Start-Process -FilePath "https://discord.gg/aeWH65qB5Y"
    })

$discordBtn.Add_MouseEnter({
        if (-not $discordFlashTimer.Enabled) {
            $discordBtn.BackColor = $colorAccent
            $discordBtn.ForeColor = [System.Drawing.Color]::White
        }
    })
$discordBtn.Add_MouseLeave({
        if (-not $discordFlashTimer.Enabled) {
            $discordBtn.BackColor = $colorSidebar
            $discordBtn.ForeColor = $colorText
        }
    })

# ---------- Barra lateral (navegacao por categoria) ----------
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.BackColor = $colorSidebar
$sidebar.Location = New-Object System.Drawing.Point(0, 64)
$sidebar.Size = New-Object System.Drawing.Size(210, 526)
$form.Controls.Add($sidebar)
Add-GradientBackground $sidebar ([System.Drawing.Color]::FromArgb(16, 16, 16)) ([System.Drawing.Color]::FromArgb(8, 8, 8))

# ---------- Painel de conteudo (checkboxes da categoria ativa) ----------
$contentOuter = New-Object System.Windows.Forms.Panel
$contentOuter.BackColor = $colorPanel
$contentOuter.Location = New-Object System.Drawing.Point(220, 74)
$contentOuter.Size = New-Object System.Drawing.Size(740, 506)
$form.Controls.Add($contentOuter)
Add-GradientBackground $contentOuter ([System.Drawing.Color]::FromArgb(20, 20, 20)) ([System.Drawing.Color]::FromArgb(11, 11, 11))
Set-RoundedRegion $contentOuter 14

$contentTitle = New-Object System.Windows.Forms.Label
$contentTitle.ForeColor = $colorAccent
$contentTitle.Font = New-Object System.Drawing.Font($fontFamily, 12, [System.Drawing.FontStyle]::Bold)
$contentTitle.Location = New-Object System.Drawing.Point(16, 12)
$contentTitle.AutoSize = $true
$contentOuter.Controls.Add($contentTitle)

$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.BackColor = $colorPanel
$contentPanel.AutoScroll = $true
$contentPanel.Location = New-Object System.Drawing.Point(10, 46)
$contentPanel.Size = New-Object System.Drawing.Size(720, 450)
$contentOuter.Controls.Add($contentPanel)

# ---------- Cria checkboxes (uma vez) e organiza por categoria ----------
$checkboxMap = @{}
$yByCategory = @{}
foreach ($cat in $categories) { $yByCategory[$cat] = 8 }

for ($i = 0; $i -lt $optimizations.Count; $i++) {
    $opt = $optimizations[$i]
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $opt.Label
    $cb.AutoSize = $false
    $cb.Width = 690
    $cb.Height = 30
    $cb.ForeColor = Get-RiskColor $opt.Risco
    $cb.BackColor = $colorPanel
    $cb.FlatStyle = "Flat"
    $cb.Location = New-Object System.Drawing.Point(6, $yByCategory[$opt.Cat])
    $checkboxMap[$i] = $cb
    $yByCategory[$opt.Cat] += 32
}

# ---------- Botoes de navegacao lateral ----------
$navButtons = @{}
$navY = 16
foreach ($cat in $categories) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "  $cat"
    $btn.TextAlign = "MiddleLeft"
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Size = New-Object System.Drawing.Size(195, 38)
    $btn.Location = New-Object System.Drawing.Point(8, $navY)
    $btn.BackColor = $colorSidebar
    $btn.ForeColor = $colorText
    $btn.Font = New-Object System.Drawing.Font($fontFamily, 9.5)
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Tag = $cat
    Set-RoundedRegion $btn 8
    $sidebar.Controls.Add($btn)
    $navButtons[$cat] = $btn
    $navY += 42
}

$navHoverColor = [System.Drawing.Color]::FromArgb(26, 26, 26)
foreach ($cat in $categories) {
    $navBtn = $navButtons[$cat]
    $navBtn.Add_MouseEnter({
            if ($this.Tag -ne $script:activeCategory) { $this.BackColor = $navHoverColor }
        }.GetNewClosure())
    $navBtn.Add_MouseLeave({
            if ($this.Tag -ne $script:activeCategory) { $this.BackColor = $colorSidebar }
        }.GetNewClosure())
    $navBtn.Add_MouseDown({
            if ($this.Tag -ne $script:activeCategory) { $this.BackColor = Adjust-Color $colorSidebar 10 }
        }.GetNewClosure())
}

function Select-Category {
    param([string]$cat)
    $script:activeCategory = $cat
    foreach ($c in $navButtons.Keys) {
        if ($c -eq $cat) {
            $navButtons[$c].BackColor = $colorAccent
            $navButtons[$c].ForeColor = [System.Drawing.Color]::White
        }
        else {
            $navButtons[$c].BackColor = $colorSidebar
            $navButtons[$c].ForeColor = $colorText
        }
    }
    $contentTitle.Text = $cat.ToUpper()
    $contentPanel.Controls.Clear()
    $itens = @()
    for ($i = 0; $i -lt $optimizations.Count; $i++) {
        if ($optimizations[$i].Cat -eq $cat) { $itens += $checkboxMap[$i] }
    }
    $contentPanel.Controls.AddRange($itens)
}

foreach ($cat in $categories) {
    $navButtons[$cat].Add_Click({ Select-Category $this.Tag })
}

Select-Category $categories[0]

# ---------- Botoes de acao ----------
function New-StyledButton {
    param($text, $x, $y, $w, $h, $bg, $fg)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Location = New-Object System.Drawing.Point($x, $y)
    $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = $bg
    $b.ForeColor = $fg
    $b.Font = New-Object System.Drawing.Font($fontFamily, 9.5, [System.Drawing.FontStyle]::Bold)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    Set-RoundedRegion $b 10
    return $b
}

$btnAll = New-StyledButton "Selecionar Tudo" 220 590 150 34 $colorPanel $colorText
$btnAll.Add_Click({ foreach ($cb in $checkboxMap.Values) { $cb.Checked = $true } })
Add-ClickEffect $btnAll $colorPanel (Adjust-Color $colorPanel 14) (Adjust-Color $colorPanel -10)
$form.Controls.Add($btnAll)

$btnNone = New-StyledButton "Desmarcar Tudo" 380 590 150 34 $colorPanel $colorText
$btnNone.Add_Click({ foreach ($cb in $checkboxMap.Values) { $cb.Checked = $false } })
Add-ClickEffect $btnNone $colorPanel (Adjust-Color $colorPanel 14) (Adjust-Color $colorPanel -10)
$form.Controls.Add($btnNone)

$btnApply = New-StyledButton "Aplicar Selecionados" 540 590 190 34 $colorAccent ([System.Drawing.Color]::White)
Add-ClickEffect $btnApply $colorAccent (Adjust-Color $colorAccent 22) (Adjust-Color $colorAccent -35)
$form.Controls.Add($btnApply)

$btnRestartColor = [System.Drawing.Color]::FromArgb(70, 40, 40)
$btnRestart = New-StyledButton "Reiniciar PC" 740 590 130 34 $btnRestartColor ([System.Drawing.Color]::FromArgb(255, 140, 140))
$btnRestart.Add_Click({ Restart-Computer -Force })
Add-ClickEffect $btnRestart $btnRestartColor (Adjust-Color $btnRestartColor 22) (Adjust-Color $btnRestartColor -18)
$form.Controls.Add($btnRestart)

# ---------- Log ----------
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = "LOG"
$logLabel.ForeColor = $colorTextDim
$logLabel.Location = New-Object System.Drawing.Point(220, 632)
$logLabel.AutoSize = $true
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(4, 4, 4)
$logBox.ForeColor = [System.Drawing.Color]::FromArgb(120, 220, 140)
$logBox.BorderStyle = "FixedSingle"
$logBox.Location = New-Object System.Drawing.Point(220, 650)
$logBox.Size = New-Object System.Drawing.Size(740, 56)
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

function Write-Log {
    param([string]$msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    $logBox.AppendText("$line`r`n")
    Add-Content -Path $logPath -Value $line
}

$btnApply.Add_Click({
        $btnApply.Enabled = $false
        $logBox.Clear()
        Write-Log "Iniciando processo de otimizacao..."

        try {
            Write-Log "Criando ponto de restauracao..."
            Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue
            $regFreq = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
            New-Item -Path $regFreq -Force | Out-Null
            Set-ItemProperty -Path $regFreq -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            Checkpoint-Computer -Description "Antes_CentralOtimizacao" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
            Write-Log "Ponto de restauracao criado com sucesso."
        }
        catch {
            Write-Log "AVISO: nao foi possivel criar o ponto de restauracao ($($_.Exception.Message))."
        }

        $applied = 0
        foreach ($i in ($checkboxMap.Keys | Sort-Object)) {
            $cb = $checkboxMap[$i]
            if ($cb.Checked) {
                $opt = $optimizations[$i]
                try {
                    Write-Log "Aplicando: $($opt.Nome)"
                    & $opt.Action
                    Write-Log "  -> OK"
                    $applied++
                    if ($opt.Restart) { $script:needsRestart = $true }
                }
                catch {
                    Write-Log "  -> ERRO: $($_.Exception.Message)"
                }
            }
        }

        Write-Log "Concluido. $applied otimizacao(oes) aplicada(s). Log: $logPath"

        if ($script:needsRestart) {
            [System.Windows.Forms.MessageBox]::Show("Otimizacoes aplicadas! Algumas exigem reiniciar o PC.", "Concluido", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Otimizacoes aplicadas com sucesso!", "Concluido", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        $btnApply.Enabled = $true
    })

[void]$form.ShowDialog()
