<#
.SYNOPSIS
    ПУК — Панель Управления Комиксами (v0.5).

.DESCRIPTION
    Графическое приложение на Windows Forms для автоматизации рутинных операций
    с коллекцией комиксов: перемещение файлов из папки загрузок в папку базы,
    конвертация CBZ → CBR через 7-Zip, быстрый доступ к рабочим каталогам.

    Конфигурация (пути, 7z.exe) сохраняется в config_PUK.json рядом со скриптом.

.NOTES
    Author : GlitchGoblin
    Version: 0.5
    Date   : 2026-11-25
#>

#region ──────────────────────────────────────────────────────────────────────
# Подключение сборок .NET
# ─────────────────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Нативные Win32-функции для перетаскивания borderless-окна
if (-not ([System.Management.Automation.PSTypeName]'Win32Drag').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class Win32Drag {
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Константы: цветовая палитра и шрифты
# ─────────────────────────────────────────────────────────────────────────────
$Colors = @{
    FormBg      = [System.Drawing.ColorTranslator]::FromHtml('#282C34')
    FormBorder  = [System.Drawing.ColorTranslator]::FromHtml('#5C6370')
    ButtonBg    = [System.Drawing.ColorTranslator]::FromHtml('#3E4451')
    ButtonText  = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
    ButtonBorder= [System.Drawing.ColorTranslator]::FromHtml('#5C6370')
    LogBg       = [System.Drawing.ColorTranslator]::FromHtml('#1E2127')
    LogText     = [System.Drawing.ColorTranslator]::FromHtml('#98C379')
    TitleText   = [System.Drawing.ColorTranslator]::FromHtml('#98C379')
    ExitButton  = [System.Drawing.ColorTranslator]::FromHtml('#808080')
    Info        = [System.Drawing.Color]::DeepSkyBlue
    Success     = [System.Drawing.Color]::Gold
    Error       = [System.Drawing.Color]::IndianRed
}

$Fonts = @{
    Main  = New-Object System.Drawing.Font('Segoe UI', 10)
    Log   = New-Object System.Drawing.Font('Consolas', 9)
    Title = New-Object System.Drawing.Font('Russo One', 16)
    Icon  = New-Object System.Drawing.Font('Segoe UI Symbol', 13)
}

# Размер и геометрия главного окна
$Layout = @{
    FormWidth   = 480
    FormHeight  = 410
    Padding     = 20
    ButtonHeight= 45
    SmallBtnH   = 30
    LogHeight   = 195
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Конфигурация приложения
# ─────────────────────────────────────────────────────────────────────────────
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = Get-Location }
$ConfigFile = Join-Path -Path $ScriptDir -ChildPath 'config_PUK.json'

# Значения по умолчанию (используются при отсутствии config-файла)
$script:Config = @{
    SourceDir  = 'C:\Users\Kilan\Downloads'
    ComicsBase = 'D:\comics_add_base'
    SevenZip   = 'C:\Program Files\7-Zip\7z.exe'
}

<#
.SYNOPSIS
    Загружает конфигурацию из JSON-файла.
.DESCRIPTION
    Если файл существует и валиден — перезаписывает дефолтные значения.
    Любые ошибки чтения игнорируются, чтобы не блокировать запуск.
#>
function Initialize-AppConfig {
    if (-not (Test-Path -LiteralPath $ConfigFile)) { return }

    try {
        $raw = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8
        $cfg = $raw | ConvertFrom-Json

        if ($cfg.SourceDir)  { $script:Config.SourceDir  = $cfg.SourceDir }
        if ($cfg.ComicsBase) { $script:Config.ComicsBase = $cfg.ComicsBase }
        if ($cfg.SevenZip)   { $script:Config.SevenZip   = $cfg.SevenZip }
    }
    catch {
        Write-Warning "Не удалось прочитать конфигурацию: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Сохраняет текущую конфигурацию в JSON-файл.
#>
function Save-AppConfig {
    $payload = [PSCustomObject]@{
        SourceDir  = $script:Config.SourceDir
        ComicsBase = $script:Config.ComicsBase
        SevenZip   = $script:Config.SevenZip
    }
    $payload | ConvertTo-Json | Set-Content -LiteralPath $ConfigFile -Encoding UTF8
}

Initialize-AppConfig
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Вспомогательные функции
# ─────────────────────────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Записывает сообщение в лог-окно с отметкой времени и цветом.
.PARAMETER Message
    Текст сообщения.
.PARAMETER Color
    Цвет текста. По умолчанию — LogText.
#>
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [System.Drawing.Color]$Color = $Colors.LogText
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $txtLog.SelectionStart  = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor  = $Color
    $txtLog.AppendText("[$timestamp] $Message`r`n")
    $txtLog.SelectionColor  = $txtLog.ForeColor

    [System.Windows.Forms.Application]::DoEvents()
}

<#
.SYNOPSIS
    Применяет единый визуальный стиль к кнопке.
#>
function Set-ButtonStyle {
    param(
        [System.Windows.Forms.Button]$Button,
        [switch]$IsExit
    )

    $Button.FlatStyle        = 'Flat'
    $Button.ForeColor        = $Colors.ButtonText
    $Button.FlatAppearance.BorderColor = $Colors.ButtonBorder
    $Button.FlatAppearance.BorderSize  = 1
    $Button.Cursor           = [System.Windows.Forms.Cursors]::Hand
    $Button.BackColor        = if ($IsExit) { $Colors.ExitButton } else { $Colors.ButtonBg }
}

<#
.SYNOPSIS
    Открывает проводник в указанной папке, создавая её при необходимости.
#>
function Open-FolderInExplorer {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
            Write-Log "Создана папка: $Path"
        }
        catch {
            Write-Log "ОШИБКА: не удалось создать $Path" -Color $Colors.Error
            return
        }
    }
    Start-Process explorer.exe $Path
    Write-Log "Открыт каталог: $Path" -Color $Colors.Info
}

<#
.SYNOPSIS
    Обработчик перетаскивания borderless-формы за её клиентскую область.
#>
function Start-FormDrag {
    param([System.Windows.Forms.Form]$Form)

    [Win32Drag]::ReleaseCapture() | Out-Null
    [Win32Drag]::SendMessage($Form.Handle, 0x00A1, [IntPtr]::new(2), [IntPtr]::Zero) | Out-Null
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Диалог настроек путей
# ─────────────────────────────────────────────────────────────────────────────
function Show-SettingsDialog {
    $sForm = New-Object System.Windows.Forms.Form
    $sForm.Text             = 'Настройки путей'
    $sForm.Size             = New-Object System.Drawing.Size(480, 270)
    $sForm.StartPosition    = 'CenterParent'
    $sForm.FormBorderStyle  = 'None'
    $sForm.BackColor        = $Colors.FormBg
    $sForm.ForeColor        = $Colors.ButtonText
    $sForm.Font             = $Fonts.Main

    $sForm.Add_Paint({
        $pen = New-Object System.Drawing.Pen($Colors.FormBorder, 2)
        $_.Graphics.DrawRectangle($pen, 0, 0, $sForm.Width - 1, $sForm.Height - 1)
    })
    $sForm.Add_MouseDown({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Start-FormDrag -Form $sForm }
    })

    # ── Поле 1: Папка загрузок ────────────────────────────────────────────
    $lblSrc = New-Object System.Windows.Forms.Label
    $lblSrc.Text      = 'Откуда забирать комиксы (Загрузки):'
    $lblSrc.Location  = New-Object System.Drawing.Point(20, 18)
    $lblSrc.AutoSize  = $true
    $sForm.Controls.Add($lblSrc)

    $txtSrc = New-Object System.Windows.Forms.TextBox
    $txtSrc.Text        = $script:Config.SourceDir
    $txtSrc.Location    = New-Object System.Drawing.Point(20, 40)
    $txtSrc.Size        = New-Object System.Drawing.Size(375, 25)
    $txtSrc.BackColor   = $Colors.LogBg
    $txtSrc.ForeColor   = $Colors.ButtonText
    $txtSrc.BorderStyle = 'FixedSingle'
    $sForm.Controls.Add($txtSrc)

    $btnBrowseSrc = New-Object System.Windows.Forms.Button
    $btnBrowseSrc.Text     = '...'
    $btnBrowseSrc.Location = New-Object System.Drawing.Point(400, 39)
    $btnBrowseSrc.Size     = New-Object System.Drawing.Size(40, 26)
    $btnBrowseSrc.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.SelectedPath = $txtSrc.Text
        if ($fbd.ShowDialog() -eq 'OK') { $txtSrc.Text = $fbd.SelectedPath }
    })
    $sForm.Controls.Add($btnBrowseSrc)

    # ── Поле 2: Папка базы ────────────────────────────────────────────────
    $lblBase = New-Object System.Windows.Forms.Label
    $lblBase.Text     = 'Каталог базы комиксов (Назначение и конвертация):'
    $lblBase.Location = New-Object System.Drawing.Point(20, 78)
    $lblBase.AutoSize = $true
    $sForm.Controls.Add($lblBase)

    $txtBase = New-Object System.Windows.Forms.TextBox
    $txtBase.Text        = $script:Config.ComicsBase
    $txtBase.Location    = New-Object System.Drawing.Point(20, 100)
    $txtBase.Size        = New-Object System.Drawing.Size(375, 25)
    $txtBase.BackColor   = $Colors.LogBg
    $txtBase.ForeColor   = $Colors.ButtonText
    $txtBase.BorderStyle = 'FixedSingle'
    $sForm.Controls.Add($txtBase)

    $btnBrowseBase = New-Object System.Windows.Forms.Button
    $btnBrowseBase.Text     = '...'
    $btnBrowseBase.Location = New-Object System.Drawing.Point(400, 99)
    $btnBrowseBase.Size     = New-Object System.Drawing.Size(40, 26)
    $btnBrowseBase.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.SelectedPath = $txtBase.Text
        if ($fbd.ShowDialog() -eq 'OK') { $txtBase.Text = $fbd.SelectedPath }
    })
    $sForm.Controls.Add($btnBrowseBase)

    # ── Поле 3: Путь к 7z.exe ─────────────────────────────────────────────
    $lbl7z = New-Object System.Windows.Forms.Label
    $lbl7z.Text     = 'Путь к консольному 7-Zip (7z.exe):'
    $lbl7z.Location = New-Object System.Drawing.Point(20, 138)
    $lbl7z.AutoSize = $true
    $sForm.Controls.Add($lbl7z)

    $txt7z = New-Object System.Windows.Forms.TextBox
    $txt7z.Text        = $script:Config.SevenZip
    $txt7z.Location    = New-Object System.Drawing.Point(20, 160)
    $txt7z.Size        = New-Object System.Drawing.Size(375, 25)
    $txt7z.BackColor   = $Colors.LogBg
    $txt7z.ForeColor   = $Colors.ButtonText
    $txt7z.BorderStyle = 'FixedSingle'
    $sForm.Controls.Add($txt7z)

    $btnBrowse7z = New-Object System.Windows.Forms.Button
    $btnBrowse7z.Text     = '...'
    $btnBrowse7z.Location = New-Object System.Drawing.Point(400, 159)
    $btnBrowse7z.Size     = New-Object System.Drawing.Size(40, 26)
    $btnBrowse7z.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter    = '7-Zip Console (7z.exe)|7z.exe|Все файлы (*.*)|*.*'
        $ofd.FileName  = $txt7z.Text
        if ($ofd.ShowDialog() -eq 'OK') { $txt7z.Text = $ofd.FileName }
    })
    $sForm.Controls.Add($btnBrowse7z)

    # ── Кнопки Сохранить / Отмена ─────────────────────────────────────────
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text     = 'Сохранить'
    $btnSave.Location = New-Object System.Drawing.Point(20, 208)
    $btnSave.Size     = New-Object System.Drawing.Size(200, 38)
    $btnSave.Add_Click({
        $script:Config.SourceDir  = $txtSrc.Text.Trim()
        $script:Config.ComicsBase = $txtBase.Text.Trim()
        $script:Config.SevenZip   = $txt7z.Text.Trim()

        Save-AppConfig
        Write-Log 'Настройки сохранены в config_PUK.json' -Color $Colors.Info
        $sForm.Close()
    })
    $sForm.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = 'Отмена'
    $btnCancel.Location = New-Object System.Drawing.Point(240, 208)
    $btnCancel.Size     = New-Object System.Drawing.Size(200, 38)
    $btnCancel.Add_Click({ $sForm.Close() })
    $sForm.Controls.Add($btnCancel)

    # Применяем стиль ко всем кнопкам диалога
    foreach ($ctrl in $sForm.Controls) {
        if ($ctrl -is [System.Windows.Forms.Button]) { Set-ButtonStyle -Button $ctrl }
    }

    [void]$sForm.ShowDialog($form)
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Бизнес-логика: операции над комиксами
# ─────────────────────────────────────────────────────────────────────────────

<#
.SYNOPSIS
    Перемещает все .cbr/.cbz из папки загрузок в папку базы.
#>
function Invoke-MoveComics {
    $source = $script:Config.SourceDir
    $dest   = $script:Config.ComicsBase

    Write-Log "Поиск файлов в $source..."

    if (-not (Test-Path -LiteralPath $source)) {
        Write-Log 'ОШИБКА: исходная папка не найдена.' -Color $Colors.Error
        return
    }
    if (-not (Test-Path -LiteralPath $dest)) {
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        Write-Log "Создана папка назначения: $dest"
    }

    $files = Get-ChildItem -LiteralPath $source -File |
             Where-Object { $_.Extension -in '.cbr', '.cbz' }

    if ($files.Count -eq 0) {
        Write-Log 'Файлы .cbr и .cbz не найдены.' -Color $Colors.Info
        return
    }

    $movedCount = 0
    foreach ($file in $files) {
        try {
            $destFile = Join-Path -Path $dest -ChildPath $file.Name
            Move-Item -LiteralPath $file.FullName -Destination $destFile -Force -ErrorAction Stop
            Write-Log "Перемещён: $($file.Name)"
            $movedCount++
        }
        catch {
            Write-Log "ОШИБКА ($($file.Name)): $($_.Exception.Message)" -Color $Colors.Error
        }
    }

    Write-Log "Готово. Всего перемещено файлов: $movedCount" -Color $Colors.Success
    Write-Log '----------------------------------------' -Color $Colors.Info
}

<#
.SYNOPSIS
    Конвертирует все .cbz в папке базы в .cbr через 7-Zip (без сжатия).
#>
function Invoke-ConvertCbzToCbr {
    $sevenZip   = $script:Config.SevenZip
    $comicsBase = $script:Config.ComicsBase

    if (-not (Test-Path -LiteralPath $sevenZip)) {
        Write-Log "ОШИБКА: 7z.exe не найден по пути $sevenZip" -Color $Colors.Error
        return
    }

    $cbzFiles = Get-ChildItem -LiteralPath $comicsBase -File |
                Where-Object { $_.Extension -eq '.cbz' }

    if ($cbzFiles.Count -eq 0) {
        Write-Log 'Файлы .cbz не найдены.' -Color $Colors.Info
        return
    }

    $converted = 0
    foreach ($file in $cbzFiles) {
        $tempDir = Join-Path -Path $env:TEMP -ChildPath ("PUK_" + [guid]::NewGuid().ToString())
        $cbrPath = Join-Path -Path $comicsBase -ChildPath ($file.BaseName + '.cbr')

        try {
            Write-Log "Перепаковка: $($file.Name)..."

            # 1) Распаковка .cbz во временную папку
            $extractArgs = @('x', "`"$($file.FullName)`"", "-o`"$tempDir`"", '-y')
            Start-Process -FilePath $sevenZip -ArgumentList $extractArgs -Wait -NoNewWindow

            # 2) Упаковка в .cbr (7z-контейнер, без сжатия — скорость важнее)
            $packArgs = @('a', '-t7z', '-mx0', "`"$cbrPath`"", "`"$tempDir\*`"")
            $proc = Start-Process -FilePath $sevenZip -ArgumentList $packArgs -Wait -NoNewWindow -PassThru

            if ($proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $cbrPath)) {
                Remove-Item -LiteralPath $file.FullName -Force
                Write-Log "Готово: $($file.BaseName).cbr"
                $converted++
            }
            else {
                Write-Log "ОШИБКА 7-Zip при обработке $($file.Name)" -Color $Colors.Error
            }
        }
        catch {
            Write-Log "ОШИБКА ($($file.Name)): $($_.Exception.Message)" -Color $Colors.Error
        }
        finally {
            if (Test-Path -LiteralPath $tempDir) {
                Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Log "Перепаковано файлов: $converted" -Color $Colors.Success
    Write-Log '----------------------------------------' -Color $Colors.Info
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Построение главного окна
# ─────────────────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text             = 'ПУК — Панель управления комиксами v0.5'
$form.Size             = New-Object System.Drawing.Size($Layout.FormWidth, $Layout.FormHeight)
$form.StartPosition    = 'CenterScreen'
$form.FormBorderStyle  = 'None'
$form.BackColor        = $Colors.FormBg
$form.Font             = $Fonts.Main

$form.Add_Paint({
    $pen = New-Object System.Drawing.Pen($Colors.FormBorder, 2)
    $_.Graphics.DrawRectangle($pen, 0, 0, $form.Width - 1, $form.Height - 1)
})
$form.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Start-FormDrag -Form $form }
})

# ── Заголовок ─────────────────────────────────────────────────────────────
$Lform = New-Object System.Windows.Forms.Label
$Lform.Text      = 'ПУК v0.5'
$Lform.Location  = New-Object System.Drawing.Point(20, 6)
$Lform.Size 	 = New-Object System.Drawing.Size(150, 28)
$Lform.ForeColor = $Colors.TitleText
$Lform.Font      = $Fonts.Title
$Lform.SendToBack()
$Lform.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Start-FormDrag -Form $form }
})
$form.Controls.Add($Lform)

# ── Кнопка настроек (шестерёнка) ──────────────────────────────────────────
$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Text             = [char]0x2699   # символ ⚙
$btnSettings.Font             = $Fonts.Icon
$btnSettings.Location         = New-Object System.Drawing.Point(442, 4)
$btnSettings.Size             = New-Object System.Drawing.Size(28, 26)
$btnSettings.FlatStyle        = 'Flat'
$btnSettings.FlatAppearance.BorderSize = 0
$btnSettings.BackColor        = $Colors.FormBg
$btnSettings.ForeColor        = [System.Drawing.Color]::LightGray
$btnSettings.Cursor           = [System.Windows.Forms.Cursors]::Hand
$btnSettings.Add_Click({ Show-SettingsDialog })
$form.Controls.Add($btnSettings)

# ── Основные действия (верхний ряд) ───────────────────────────────────────
$btnMove = New-Object System.Windows.Forms.Button
$btnMove.Text     = 'Переместить комиксы'
$btnMove.Location = New-Object System.Drawing.Point(20, 35)
$btnMove.Size     = New-Object System.Drawing.Size(215, $Layout.ButtonHeight)
$btnMove.Add_Click({ Invoke-MoveComics })
$form.Controls.Add($btnMove)

$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Text     = 'CBZ to CBR'
$btnConvert.Location = New-Object System.Drawing.Point(245, 35)
$btnConvert.Size     = New-Object System.Drawing.Size(215, $Layout.ButtonHeight)
$btnConvert.Add_Click({ Invoke-ConvertCbzToCbr })
$form.Controls.Add($btnConvert)

# ── Открытие рабочих каталогов (второй ряд) ───────────────────────────────
$btnOpenSrc = New-Object System.Windows.Forms.Button
$btnOpenSrc.Text     = 'Открыть каталог загрузок'
$btnOpenSrc.Location = New-Object System.Drawing.Point(20, 85)
$btnOpenSrc.Size     = New-Object System.Drawing.Size(215, $Layout.SmallBtnH)
$btnOpenSrc.Add_Click({ Open-FolderInExplorer -Path $script:Config.SourceDir })
$form.Controls.Add($btnOpenSrc)

$btnOpenBase = New-Object System.Windows.Forms.Button
$btnOpenBase.Text     = 'Открыть каталог базы'
$btnOpenBase.Location = New-Object System.Drawing.Point(245, 85)
$btnOpenBase.Size     = New-Object System.Drawing.Size(215, $Layout.SmallBtnH)
$btnOpenBase.Add_Click({ Open-FolderInExplorer -Path $script:Config.ComicsBase })
$form.Controls.Add($btnOpenBase)

# ── Окно логов ────────────────────────────────────────────────────────────
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location    = New-Object System.Drawing.Point(20, 120)
$txtLog.Size        = New-Object System.Drawing.Size(440, $Layout.LogHeight)
$txtLog.ScrollBars  = 'Vertical'
$txtLog.ReadOnly    = $true
$txtLog.BackColor   = $Colors.LogBg
$txtLog.ForeColor   = $Colors.LogText
$txtLog.Font        = $Fonts.Log
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.WordWrap    = $false
$form.Controls.Add($txtLog)

# ── Нижние кнопки: Очистить и Выйти ───────────────────────────────────────
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Очистить'
$btnClear.Location = New-Object System.Drawing.Point(20, 320)
$btnClear.Size     = New-Object System.Drawing.Size(440, 25)
$btnClear.Add_Click({ $txtLog.Clear() })
$form.Controls.Add($btnClear)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text     = 'Выйти'
$btnExit.Location = New-Object System.Drawing.Point(20, 350)
$btnExit.Size     = New-Object System.Drawing.Size(440, 35)
$btnExit.Add_Click({ $form.Close() })
$form.Controls.Add($btnExit)

# ── Финальная стилизация всех кнопок ──────────────────────────────────────
foreach ($control in $form.Controls) {
    if ($control -is [System.Windows.Forms.Button] -and $control -ne $btnSettings) {
        Set-ButtonStyle -Button $control -IsExit:($control -eq $btnExit)
    }
}
#endregion

#region ──────────────────────────────────────────────────────────────────────
# Точка входа
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'Панель управления комиксами готова к работе.'
$form.ShowDialog() | Out-Null
#endregion