Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object Windows.Forms.Form
$form.Text = "DOC → PDF Конвертер"
$form.Size = New-Object Drawing.Size(700,500)
$form.StartPosition = "CenterScreen"

$lbl = New-Object Windows.Forms.Label
$lbl.Text = "Выберите папку с DOC/DOCX файлами"
$lbl.Location = New-Object Drawing.Point(20,20)
$lbl.AutoSize = $true

$txtPath = New-Object Windows.Forms.TextBox
$txtPath.Location = New-Object Drawing.Point(20,50)
$txtPath.Size = New-Object Drawing.Size(520,25)
$txtPath.Text = "D:\"

$btnBrowse = New-Object Windows.Forms.Button
$btnBrowse.Text = "Обзор"
$btnBrowse.Location = New-Object Drawing.Point(550,48)
$btnBrowse.Size = New-Object Drawing.Size(100,30)

$btnConvert = New-Object Windows.Forms.Button
$btnConvert.Text = "Конвертировать"
$btnConvert.Location = New-Object Drawing.Point(20,90)
$btnConvert.Size = New-Object Drawing.Size(150,35)

$progress = New-Object Windows.Forms.ProgressBar
$progress.Location = New-Object Drawing.Point(20,140)
$progress.Size = New-Object Drawing.Size(630,25)

$log = New-Object Windows.Forms.TextBox
$log.Location = New-Object Drawing.Point(20,180)
$log.Size = New-Object Drawing.Size(630,250)
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.ReadOnly = $true

$form.Controls.AddRange(@(
    $lbl,
    $txtPath,
    $btnBrowse,
    $btnConvert,
    $progress,
    $log
))

$folderDialog = New-Object Windows.Forms.FolderBrowserDialog

function Get-DriveFreeSpace {
    param([string]$Path)
    try {
        $resolved = (Resolve-Path $Path).Path
        $root = [System.IO.Path]::GetPathRoot($resolved)
        if ([string]::IsNullOrWhiteSpace($root)) { return 0 }
        return (New-Object System.IO.DriveInfo($root)).AvailableFreeSpace
    }
    catch {
        return 0
    }
}

function Format-Bytes {
    param([int64]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes байт" }
    if ($Bytes -lt 1MB) { return "{0:N0} КБ" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB) { return "{0:N1} МБ" -f ($Bytes / 1MB) }
    return "{0:N2} ГБ" -f ($Bytes / 1GB)
}

$btnBrowse.Add_Click({
    if ($folderDialog.ShowDialog() -eq "OK") {
        $txtPath.Text = $folderDialog.SelectedPath
    }
})

$btnConvert.Add_Click({

    $folder = $txtPath.Text

    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path $folder)) {
        [Windows.Forms.MessageBox]::Show("Выберите корректную папку.")
        return
    }

    $pdfFolder = Join-Path $folder "PDF"
    if (!(Test-Path $pdfFolder)) {
        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
        $log.AppendText("[$timestamp] Создание папки: $pdfFolder`r`n")
        New-Item -ItemType Directory -Path $pdfFolder -Force | Out-Null
        if (!(Test-Path $pdfFolder)) {
            [Windows.Forms.MessageBox]::Show("Не удалось создать папку PDF: $pdfFolder")
            $btnBrowse.Enabled = $true
            $btnConvert.Enabled = $true
            return
        }
    }
    
    $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
    $log.AppendText("[$timestamp] Папка PDF: $pdfFolder`r`n")
    $log.Refresh()

    $files = @(Get-ChildItem $folder -File | Where-Object {
        $_.Extension -in ".doc", ".docx"
    })

    if ($files.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show("DOC/DOCX файлы не найдены.")
        return
    }

    $freeSpace = Get-DriveFreeSpace $folder
    if ($freeSpace -lt 100MB) {
        [Windows.Forms.MessageBox]::Show("На диске недостаточно свободного места: $(Format-Bytes $freeSpace). Требуется минимум 100 МБ.")
        return
    }

    $btnBrowse.Enabled = $false
    $btnConvert.Enabled = $false
    $progress.Maximum = $files.Count
    $progress.Value = 0

    $log.Clear()
    $log.AppendText("Запуск конвертации...`r`n`r`n")

    $word = $null
    try {
        $log.AppendText("Microsoft Word инициализация...`r`n")
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.ScreenUpdating = $false
        try { $word.AskToUpdateLinks = $false } catch {}
        try { $word.Options.UpdateLinksAtOpen = 0 } catch {}
        try { $word.AutomationSecurity = 3 } catch {}

        $i = 0
        foreach ($file in $files) {
            $doc = $null
            try {
                $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                $log.AppendText("[$timestamp] $($file.Name)...")
                $log.Refresh()
                [System.Windows.Forms.Application]::DoEvents()

                $doc = $null
                $openSuccess = $false
                $fileAbsolutePath = $file.FullName
                
                # Фоновый таймер для убийства зависшего процесса
                $killJob = Start-Job -ScriptBlock {
                    [System.Threading.Thread]::Sleep(5000)
                    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                }
                
                # Попытка 1: Открытие со стандартными параметрами
                try {
                    $doc = $word.Documents.Open($fileAbsolutePath, $false, $false, $false)
                    if ($doc -ne $null) {
                        $openSuccess = $true
                    }
                } catch {}
                
                # Отменяем таймер убийства
                Stop-Job $killJob -ErrorAction SilentlyContinue
                Remove-Job $killJob -ErrorAction SilentlyContinue
                
                # Проверяем, жив ли Word процесс
                $wordProcess = Get-Process WINWORD -ErrorAction SilentlyContinue
                if (-not $wordProcess -and -not $openSuccess) {
                    # Word был убит таймером - пересоздаём его
                    $word = $null
                    [System.Threading.Thread]::Sleep(300)
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0
                    $word.ScreenUpdating = $false
                    try { $word.AskToUpdateLinks = $false } catch {}
                    try { $word.Options.UpdateLinksAtOpen = 0 } catch {}
                    try { $word.AutomationSecurity = 3 } catch {}
                }
                
                # Попытка 2: Открытие с параметром ReadOnly
                if (-not $openSuccess) {
                    $killJob = Start-Job -ScriptBlock {
                        [System.Threading.Thread]::Sleep(5000)
                        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                    
                    try {
                        $doc = $word.Documents.Open($fileAbsolutePath, $false, $true, $false)
                        if ($doc -ne $null) {
                            $openSuccess = $true
                        }
                    } catch {}
                    
                    Stop-Job $killJob -ErrorAction SilentlyContinue
                    Remove-Job $killJob -ErrorAction SilentlyContinue
                    
                    $wordProcess = Get-Process WINWORD -ErrorAction SilentlyContinue
                    if (-not $wordProcess) {
                        $word = $null
                        [System.Threading.Thread]::Sleep(300)
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                        $word.DisplayAlerts = 0
                        $word.ScreenUpdating = $false
                        try { $word.AskToUpdateLinks = $false } catch {}
                        try { $word.Options.UpdateLinksAtOpen = 0 } catch {}
                        try { $word.AutomationSecurity = 3 } catch {}
                    }
                }
                
                # Попытка 3: Открытие без обновления связей + ConfirmConversions
                if (-not $openSuccess) {
                    $killJob = Start-Job -ScriptBlock {
                        [System.Threading.Thread]::Sleep(5000)
                        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                    
                    try {
                        # ConfirmConversions=False (4-й параметр) - не показывать диалоги конвертации
                        $doc = $word.Documents.Open($fileAbsolutePath, $false, $false, $true)
                        if ($doc -ne $null) {
                            $openSuccess = $true
                        }
                    } catch {}
                    
                    Stop-Job $killJob -ErrorAction SilentlyContinue
                    Remove-Job $killJob -ErrorAction SilentlyContinue
                    
                    $wordProcess = Get-Process WINWORD -ErrorAction SilentlyContinue
                    if (-not $wordProcess) {
                        $word = $null
                        [System.Threading.Thread]::Sleep(300)
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                        $word.DisplayAlerts = 0
                        $word.ScreenUpdating = $false
                        try { $word.AskToUpdateLinks = $false } catch {}
                        try { $word.Options.UpdateLinksAtOpen = 0 } catch {}
                        try { $word.AutomationSecurity = 3 } catch {}
                    }
                }
                
                # Попытка 4: Перезагружаем Word с DisableAllMacros и пробуем ещё раз
                if (-not $openSuccess) {
                    try {
                        try { $word.Quit() } catch {}
                        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
                        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
                        $word = $null
                        
                        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                        [System.Threading.Thread]::Sleep(500)
                        
                        $word = New-Object -ComObject Word.Application
                        $word.Visible = $false
                        $word.DisplayAlerts = 0
                        $word.ScreenUpdating = $false
                        try { $word.AskToUpdateLinks = $false } catch {}
                        try { $word.Options.UpdateLinksAtOpen = 0 } catch {}
                        try { $word.AutomationSecurity = 3 } catch {}
                        try { $word.Options.DisableAllMacros = 1 } catch {}
                        
                        $killJob = Start-Job -ScriptBlock {
                            [System.Threading.Thread]::Sleep(5000)
                            Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                        }
                        
                        $doc = $word.Documents.Open($fileAbsolutePath, $false, $true, $false)
                        if ($doc -ne $null) {
                            $openSuccess = $true
                        }
                        
                        Stop-Job $killJob -ErrorAction SilentlyContinue
                        Remove-Job $killJob -ErrorAction SilentlyContinue
                    } catch {}
                }

                if (-not $openSuccess) {
                    $log.AppendText(" ⚠️`r`n")
                    $log.Refresh()
                    continue
                }
                
                $log.AppendText(" ✓`r`n")

                $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                $log.AppendText("[$timestamp] Документ открыт`r`n")
                $log.Refresh()

                $pdfFile = Join-Path $pdfFolder ($file.BaseName + ".pdf")
                $pdfFile = [System.IO.Path]::GetFullPath($pdfFile)
                
                $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                $log.AppendText("[$timestamp] Путь сохранения: $pdfFile`r`n")
                $log.Refresh()
                
                if (Test-Path $pdfFile) {
                    Remove-Item $pdfFile -Force
                }

                $saved = $false
                
                try {
                    $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                    $log.AppendText("[$timestamp] SaveAs2 (PDF)...`r`n")
                    $log.Refresh()
                    
                    if (-not (Test-Path $pdfFolder)) {
                        throw "Папка PDF не существует: $pdfFolder"
                    }
                    
                    $doc.SaveAs2($pdfFile, 17)
                    
                    # Минимальная задержка для завершения сохранения
                    [System.Threading.Thread]::Sleep(100)
                    
                    $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                    if (Test-Path $pdfFile) {
                        $fileSize = (Get-Item $pdfFile).Length
                        $log.AppendText("[$timestamp] SaveAs2 успешно (размер: $fileSize байт)`r`n")
                        $saved = $true
                    } else {
                        $log.AppendText("[$timestamp] SaveAs2: файл не создан`r`n")
                    }
                } catch {
                    $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                    $log.AppendText("[$timestamp] SaveAs2 failed: $($_.Exception.Message)`r`n")
                }

                if (-not $saved) {
                    try {
                        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                        $log.AppendText("[$timestamp] SaveAs (PDF)...`r`n")
                        $log.Refresh()
                        
                        $doc.SaveAs($pdfFile, 17)
                        $saved = $true
                        
                        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                        $log.AppendText("[$timestamp] SaveAs успешно`r`n")
                    } catch {
                        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                        $log.AppendText("[$timestamp] SaveAs failed: $($_.Exception.Message)`r`n")
                    }
                }

                if ($saved) {
                    $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                    $log.AppendText("[$timestamp] Создан: $pdfFile`r`n`r`n")
                } else {
                    throw "Не удалось сохранить PDF"
                }
            }
            catch {
                $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
                $errorMsg = $_.Exception.Message
                $log.AppendText("[$timestamp] ❌ Ошибка обработки файла: $errorMsg`r`n`r`n")
                $log.Refresh()
            }
            finally {
                if ($doc -ne $null) {
                    try { $doc.Close($false, $null, $false) } catch {}
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
                    $doc = $null
                }
            }

            $i++
            $progress.Value = $i
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Единовременная очистка после обработки всех файлов
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    
        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
        $log.AppendText("[$timestamp] Готово. Обработано файлов: $i`r`n")
        [Windows.Forms.MessageBox]::Show("Конвертация завершена! Обработано файлов: $i")
    }
    catch {
        [Windows.Forms.MessageBox]::Show("Ошибка при обработке файлов: $($_.Exception.Message)")
        $timestamp = [DateTime]::Now.ToString("HH:mm:ss.fff")
        $log.AppendText("[$timestamp] Фатальная ошибка: $($_.Exception.Message)`r`n")
    }
    finally {
        if ($word -ne $null) {
            try { $word.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
            $word = $null
        }
        
        # Убедимся, что все процессы Word завершены
        Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        $btnBrowse.Enabled = $true
        $btnConvert.Enabled = $true
    }
})

$form.ShowDialog()