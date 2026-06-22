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

$titleLbl = New-Object System.Windows.Forms.Label
$titleLbl.Text = "CENTRAL DE OTIMIZACAO"
$titleLbl.ForeColor = $colorText
$titleLbl.Font = New-Object System.Drawing.Font($fontFamily, 15, [System.Drawing.FontStyle]::Bold)
$titleLbl.Location = New-Object System.Drawing.Point(24, 10)
$titleLbl.AutoSize = $true
$headerPanel.Controls.Add($titleLbl)

$subtitleLbl = New-Object System.Windows.Forms.Label
$subtitleLbl.Text = "Marque as otimizacoes desejadas e clique em Aplicar Selecionados"
$subtitleLbl.ForeColor = $colorTextDim
$subtitleLbl.Font = New-Object System.Drawing.Font($fontFamily, 9)
$subtitleLbl.Location = New-Object System.Drawing.Point(26, 38)
$subtitleLbl.AutoSize = $true
$headerPanel.Controls.Add($subtitleLbl)

# Legenda de risco (quadradinho colorido em vez de caractere especial - evita bug de encoding)
$legendaX = 560
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
    $legendaX += 80
}

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
