#Requires -Version 5.1
<#
    workbooks-gui.ps1
    Arabic right-to-left desktop window for the workbook generator.

    Runs `node build.js` (books tab) or `node image_2_cover_style/images-to-book.js`
    (images tab) as a hidden child process, streams its output into the log, and
    turns the ##STAGE / ##BOOK / ##OUTPUT / ##DONE markers into progress and
    "open file" buttons.

    This file is pure ASCII on purpose. Every Arabic string lives in gui.xaml
    and is read back with FindResource, so Windows PowerShell 5.1 can never
    mis-decode it (it reads BOM-less .ps1 files as ANSI).

    Started by start-gui.bat with -NoProfile -ExecutionPolicy Bypass -STA.
#>

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$script:Root = Split-Path -Parent $PSScriptRoot
$script:XamlPath = Join-Path $PSScriptRoot 'gui.xaml'

$script:Process = $null
$script:IsBuilding = $false
$script:StopRequested = $false
$script:JobKind = $null
$script:LogPath = $null
$script:LogOffset = 0
$script:LogRemainder = ''
$script:ActiveLog = $null
$script:ActiveResults = $null
$script:ActiveResultsList = @()
$script:ActiveStatus = $null
$script:ActiveProgress = $null
$script:BookTitles = @{}
$script:BookBoxes = @()

try {
    # -- load the window -----------------------------------------------------
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($script:XamlPath)
    $reader = New-Object System.Xml.XmlNodeReader($xml)
    $script:Window = [Windows.Markup.XamlReader]::Load($reader)

    function Get-Ui([string] $name) { return $script:Window.FindName($name) }
    function Get-Str([string] $key) { return [string]$script:Window.FindResource($key) }

    # -- controls ------------------------------------------------------------
    $script:BooksPanel          = Get-Ui 'BooksPanel'
    $script:SelectAllButton     = Get-Ui 'SelectAllButton'
    $script:SelectNoneButton    = Get-Ui 'SelectNoneButton'
    $script:OrderAbjadi         = Get-Ui 'OrderAbjadi'
    $script:OrderAlphabetical   = Get-Ui 'OrderAlphabetical'
    $script:OrderPreview        = Get-Ui 'OrderPreview'
    $script:OutDirBox           = Get-Ui 'OutDirBox'
    $script:BrowseOutButton     = Get-Ui 'BrowseOutButton'
    $script:MergeCheck          = Get-Ui 'MergeCheck'
    $script:MergedNameBox       = Get-Ui 'MergedNameBox'
    $script:BuildButton         = Get-Ui 'BuildButton'
    $script:StopButton          = Get-Ui 'StopButton'
    $script:OpenOutButton       = Get-Ui 'OpenOutButton'
    $script:StatusText          = Get-Ui 'StatusText'
    $script:BuildProgress       = Get-Ui 'BuildProgress'
    $script:LogBox              = Get-Ui 'LogBox'
    $script:ResultsPanel        = Get-Ui 'ResultsPanel'

    $script:ImagesDirBox        = Get-Ui 'ImagesDirBox'
    $script:BrowseImagesButton  = Get-Ui 'BrowseImagesButton'
    $script:ImagesOutBox        = Get-Ui 'ImagesOutBox'
    $script:BrowseImagesOutButton = Get-Ui 'BrowseImagesOutButton'
    $script:SizeBox             = Get-Ui 'SizeBox'
    $script:ImagesDocxOnly      = Get-Ui 'ImagesDocxOnly'
    $script:ImagesBuildButton   = Get-Ui 'ImagesBuildButton'
    $script:ImagesStopButton    = Get-Ui 'ImagesStopButton'
    $script:ImagesOpenOutButton = Get-Ui 'ImagesOpenOutButton'
    $script:ImagesStatusText    = Get-Ui 'ImagesStatusText'
    $script:ImagesProgress      = Get-Ui 'ImagesProgress'
    $script:ImagesLogBox        = Get-Ui 'ImagesLogBox'
    $script:ImagesResultsPanel  = Get-Ui 'ImagesResultsPanel'

    # -- strings -------------------------------------------------------------
    $script:StrOpen      = Get-Str 'Str.Open'
    $script:StrPages     = Get-Str 'Str.Pages'
    $script:StrPdfYes    = Get-Str 'Str.PdfYes'
    $script:StrPdfNo     = Get-Str 'Str.PdfNo'
    $script:StrPick      = Get-Str 'Str.PickFolder'
    $script:StrTitle     = Get-Str 'Str.Title'
    $script:StrNoBooks   = Get-Str 'Str.NoBooks'
    $script:StrNoOutDir  = Get-Str 'Str.NoOutDir'
    $script:StrNoImages  = Get-Str 'Str.NoImages'
    $script:StrBadMerged = Get-Str 'Str.BadMerged'
    $script:StrNoNode    = Get-Str 'Str.NoNode'
    $script:StrBuildErr  = Get-Str 'Str.BuildError'
    $script:StrReady     = Get-Str 'Str.Ready'
    $script:StrRunning   = Get-Str 'Str.Running'
    $script:StrStopped   = Get-Str 'Str.Stopped'
    $script:StrDone      = Get-Str 'Str.Done'
    $script:StrFailed    = Get-Str 'Str.Failed'
    $script:StrStage     = Get-Str 'Str.Stage'
    $script:StrBook      = Get-Str 'Str.Book'
    $script:PreviewAbjadi = Get-Str 'Str.Preview.Abjadi'
    $script:PreviewAlpha  = Get-Str 'Str.Preview.Alphabetical'

    # -- book list -----------------------------------------------------------
    # Read as UTF-8 explicitly: Windows PowerShell 5.1 decodes a BOM-less file
    # with the ANSI code page, which turns the Arabic titles into mojibake.
    $catalogPath = Join-Path $script:Root 'books.json'
    $catalog = [System.IO.File]::ReadAllText($catalogPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

    foreach ($book in $catalog.books) {
        $script:BookTitles[[string]$book.dir] = [string]$book.title

        $box = New-Object System.Windows.Controls.CheckBox
        $box.Content = [string]$book.title
        $box.Tag = [string]$book.dir
        $box.IsChecked = $true
        $box.MinWidth = 200
        $box.Margin = New-Object System.Windows.Thickness(0, 4, 26, 4)

        $pdfPath = Join-Path (Join-Path $script:Root $book.dir) ($book.dir + '.pdf')
        $pdfHint = if (Test-Path -LiteralPath $pdfPath) { $script:StrPdfYes } else { $script:StrPdfNo }
        $box.ToolTip = ('{0} {1}{2}{3}' -f $book.pages, $script:StrPages, [Environment]::NewLine, $pdfHint)

        $script:BooksPanel.Children.Add($box) | Out-Null
        $script:BookBoxes += $box
    }

    # -- defaults ------------------------------------------------------------
    $script:OutDirBox.Text = $script:Root
    $script:ImagesDirBox.Text = Join-Path $script:Root 'image_2_cover_style\images'
    $script:ImagesOutBox.Text = Join-Path $script:Root 'image_2_cover_style'
    $script:SizeBox.ItemsSource = @('A5', 'A4')
    $script:SizeBox.SelectedIndex = 0
    $script:StatusText.Text = $script:StrReady
    $script:ImagesStatusText.Text = $script:StrReady
    $script:OrderPreview.Text = $script:PreviewAbjadi

    # -- helpers -------------------------------------------------------------
    function Show-Warning([string] $message) {
        [void][System.Windows.MessageBox]::Show($message, $script:StrTitle,
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }

    function Select-Folder([string] $initial) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $script:StrPick
        $dialog.ShowNewFolderButton = $true
        if ($initial -and (Test-Path -LiteralPath $initial)) {
            $dialog.SelectedPath = $initial
        }
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        return $null
    }

    function Resolve-OutputFolder([string] $value) {
        if (-not $value) { return $null }
        try { $full = [System.IO.Path]::GetFullPath($value) } catch { return $null }
        if (-not (Test-Path -LiteralPath $full)) {
            try { New-Item -ItemType Directory -Path $full -Force | Out-Null }
            catch { return $null }
        }
        return $full
    }

    function Get-BookTitle([string] $dir) {
        if ($script:BookTitles.ContainsKey($dir)) { return $script:BookTitles[$dir] }
        return $dir
    }

    function Get-OrderName {
        if ($script:OrderAlphabetical.IsChecked) { return 'alphabetical' }
        return 'abjadi'
    }

    function Add-Log([string] $line) {
        if (-not $script:ActiveLog) { return }
        $script:ActiveLog.AppendText($line + [Environment]::NewLine)
        $script:ActiveLog.ScrollToEnd()
    }

    function Add-Result([string] $path) {
        if (-not $script:ActiveResults) { return }
        if (-not $path) { return }
        if ($script:ActiveResultsList -contains $path) { return }
        $script:ActiveResultsList += $path

        $button = New-Object System.Windows.Controls.Button
        # A TextBlock, not a plain string: WPF treats "_" in string content as an
        # access-key marker and would drop it from file names like
        # Arabic_Word_Analysis_Workbook.docx.
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = '{0}  {1}' -f $script:StrOpen, [System.IO.Path]::GetFileName($path)
        $button.Content = $label
        $button.Tag = $path
        $button.ToolTip = $path
        $button.Margin = New-Object System.Windows.Thickness(0, 4, 8, 4)
        $button.Padding = New-Object System.Windows.Thickness(14, 4, 14, 4)
        $button.Add_Click({
            param($sender, $eventArgs)
            $target = [string]$sender.Tag
            try {
                Start-Process -FilePath $target
            }
            catch {
                try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"' + $target + '"') }
                catch { }
            }
        })
        $script:ActiveResults.Children.Add($button) | Out-Null
    }

    function Handle-Line([string] $line) {
        $line = $line.TrimEnd("`r")

        if ($line -match '^##STAGE\s+(\d+)/(\d+)\s+(\S+)') {
            $index = [int]$Matches[1]
            $total = [int]$Matches[2]
            if ($script:ActiveStatus) {
                $script:ActiveStatus.Text = '{0} {1}/{2}' -f $script:StrStage, $index, $total
            }
            if ($script:ActiveProgress) {
                $script:ActiveProgress.Value = (($index - 1) * 100.0) / [Math]::Max($total, 1)
            }
            return
        }

        if ($line -match '^##BOOK\s+(\S+)\s+(building|ok|failed)') {
            $dir = $Matches[1]
            $state = $Matches[2]
            if ($state -eq 'building' -and $script:ActiveStatus) {
                $script:ActiveStatus.Text = '{0}: {1}' -f $script:StrBook, (Get-BookTitle $dir)
            }
            return
        }

        if ($line -match '^##OUTPUT\s+(.+)$') {
            Add-Result $Matches[1].Trim()
            return
        }

        if ($line -match '^##DONE\s+(\w+)') {
            if ($script:ActiveStatus) {
                if ($Matches[1] -eq 'ok') { $script:ActiveStatus.Text = $script:StrDone }
                else { $script:ActiveStatus.Text = $script:StrFailed }
            }
            return
        }

        if ($line -match '^##(INFO|CONFIG)\s+(.+)$') {
            Add-Log $Matches[2]
            return
        }

        Add-Log $line
    }

    function Read-NewLog {
        if (-not $script:LogPath) { return }
        if (-not (Test-Path -LiteralPath $script:LogPath)) { return }

        $stream = $null
        try {
            $stream = [System.IO.File]::Open($script:LogPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite)
            if ($stream.Length -le $script:LogOffset) { return }

            $stream.Position = $script:LogOffset
            $count = [int]($stream.Length - $script:LogOffset)
            $buffer = New-Object byte[] $count
            $read = $stream.Read($buffer, 0, $count)
            $script:LogOffset += $read

            $text = $script:LogRemainder + [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
            $parts = $text -split "`n"
            $script:LogRemainder = $parts[$parts.Count - 1]
            for ($i = 0; $i -lt $parts.Count - 1; $i += 1) {
                Handle-Line $parts[$i]
            }
        }
        catch {
            # A half-written file is normal while the child is still starting.
        }
        finally {
            if ($stream) { $stream.Dispose() }
        }
    }

    function ConvertTo-Argument([string] $value) {
        if ($value -match '^[^\s"]+$') { return $value }
        return '"' + ($value -replace '"', '\"') + '"'
    }

    function Get-NodePath {
        $node = Get-Command node -ErrorAction SilentlyContinue
        if (-not $node) { return $null }
        return [string]$node.Source
    }

    function Start-JobRun {
        param([string[]] $NodeArgs, [string] $Kind)

        $node = Get-NodePath
        if (-not $node) {
            Show-Warning $script:StrNoNode
            return
        }

        $script:JobKind = $Kind
        if ($Kind -eq 'books') {
            $script:ActiveLog = $script:LogBox
            $script:ActiveResults = $script:ResultsPanel
            $script:ActiveStatus = $script:StatusText
            $script:ActiveProgress = $script:BuildProgress
            $script:BuildProgress.IsIndeterminate = $false
            $script:BuildProgress.Value = 0
        }
        else {
            $script:ActiveLog = $script:ImagesLogBox
            $script:ActiveResults = $script:ImagesResultsPanel
            $script:ActiveStatus = $script:ImagesStatusText
            $script:ActiveProgress = $script:ImagesProgress
            $script:ImagesProgress.IsIndeterminate = $true
        }

        $script:ActiveResults.Children.Clear()
        $script:ActiveResultsList = @()
        $script:ActiveLog.Clear()
        $script:StopRequested = $false

        $script:LogPath = Join-Path $env:TEMP ('wb-gui-' + [Guid]::NewGuid().ToString('N') + '.log')
        $script:LogOffset = 0
        $script:LogRemainder = ''

        $commandLine = (ConvertTo-Argument $node) + ' ' +
            (($NodeArgs | ForEach-Object { ConvertTo-Argument ([string]$_) }) -join ' ')
        $inner = $commandLine + ' > ' + (ConvertTo-Argument $script:LogPath) + ' 2>&1'

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $env:ComSpec
        $startInfo.Arguments = '/d /s /c "' + $inner + '"'
        $startInfo.WorkingDirectory = $script:Root
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        try {
            $script:Process = New-Object System.Diagnostics.Process
            $script:Process.StartInfo = $startInfo
            [void]$script:Process.Start()
        }
        catch {
            $script:Process = $null
            [void][System.Windows.MessageBox]::Show(
                ($script:StrBuildErr + $_.Exception.Message), $script:StrTitle,
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        $script:IsBuilding = $true
        $script:BuildButton.IsEnabled = $false
        $script:ImagesBuildButton.IsEnabled = $false
        $script:StopButton.IsEnabled = ($Kind -eq 'books')
        $script:ImagesStopButton.IsEnabled = ($Kind -eq 'images')
        $script:ActiveStatus.Text = $script:StrRunning
        $script:Timer.Start()
    }

    function Stop-JobRun {
        if (-not $script:Process) { return }
        $script:StopRequested = $true
        try {
            & taskkill.exe /PID $script:Process.Id /T /F 2>$null | Out-Null
        }
        catch { }
    }

    function Complete-JobRun([int] $exitCode) {
        $script:Timer.Stop()
        Read-NewLog
        if ($script:LogRemainder) {
            Add-Log $script:LogRemainder
            $script:LogRemainder = ''
        }

        if ($script:JobKind -eq 'images' -and $exitCode -eq 0 -and -not $script:StopRequested) {
            $outDir = $script:ImagesOutBox.Text.Trim()
            Add-Result (Join-Path $outDir 'images-to-book.docx')
            if (-not $script:ImagesDocxOnly.IsChecked) {
                Add-Result (Join-Path $outDir 'images-to-book.pdf')
            }
        }

        if ($script:ActiveProgress) { $script:ActiveProgress.IsIndeterminate = $false }

        if ($script:StopRequested) {
            $script:ActiveStatus.Text = $script:StrStopped
        }
        elseif ($exitCode -eq 0) {
            $script:ActiveStatus.Text = $script:StrDone
            if ($script:ActiveProgress) { $script:ActiveProgress.Value = 100 }
        }
        else {
            $script:ActiveStatus.Text = $script:StrFailed
        }

        $script:IsBuilding = $false
        $script:BuildButton.IsEnabled = $true
        $script:ImagesBuildButton.IsEnabled = $true
        $script:StopButton.IsEnabled = $false
        $script:ImagesStopButton.IsEnabled = $false
        $script:Process = $null

        if ($script:LogPath -and (Test-Path -LiteralPath $script:LogPath)) {
            Remove-Item -LiteralPath $script:LogPath -Force -ErrorAction SilentlyContinue
        }
    }

    # -- timer ---------------------------------------------------------------
    $script:Timer = New-Object System.Windows.Threading.DispatcherTimer
    $script:Timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $script:Timer.Add_Tick({
        Read-NewLog
        if ($script:Process -and $script:Process.HasExited) {
            $code = 0
            try { $code = $script:Process.ExitCode } catch { }
            Complete-JobRun $code
        }
    })

    # -- events --------------------------------------------------------------
    $script:OrderAbjadi.Add_Checked({
        $script:OrderPreview.Text = $script:PreviewAbjadi
    })
    $script:OrderAlphabetical.Add_Checked({
        $script:OrderPreview.Text = $script:PreviewAlpha
    })

    $script:SelectAllButton.Add_Click({
        foreach ($box in $script:BookBoxes) { $box.IsChecked = $true }
    })
    $script:SelectNoneButton.Add_Click({
        foreach ($box in $script:BookBoxes) { $box.IsChecked = $false }
    })

    $script:BrowseOutButton.Add_Click({
        $picked = Select-Folder $script:OutDirBox.Text
        if ($picked) { $script:OutDirBox.Text = $picked }
    })
    $script:BrowseImagesButton.Add_Click({
        $picked = Select-Folder $script:ImagesDirBox.Text
        if ($picked) { $script:ImagesDirBox.Text = $picked }
    })
    $script:BrowseImagesOutButton.Add_Click({
        $picked = Select-Folder $script:ImagesOutBox.Text
        if ($picked) { $script:ImagesOutBox.Text = $picked }
    })

    $script:OpenOutButton.Add_Click({
        $dir = $script:OutDirBox.Text.Trim()
        if (Test-Path -LiteralPath $dir) { Start-Process -FilePath 'explorer.exe' -ArgumentList $dir }
    })
    $script:ImagesOpenOutButton.Add_Click({
        $dir = $script:ImagesOutBox.Text.Trim()
        if (Test-Path -LiteralPath $dir) { Start-Process -FilePath 'explorer.exe' -ArgumentList $dir }
    })

    $script:StopButton.Add_Click({ Stop-JobRun })
    $script:ImagesStopButton.Add_Click({ Stop-JobRun })

    $script:BuildButton.Add_Click({
        if ($script:IsBuilding) { return }

        $dirs = @()
        foreach ($box in $script:BookBoxes) {
            if ($box.IsChecked) { $dirs += [string]$box.Tag }
        }
        if ($dirs.Count -eq 0) { Show-Warning $script:StrNoBooks; return }

        $outDir = Resolve-OutputFolder $script:OutDirBox.Text.Trim()
        if (-not $outDir) { Show-Warning $script:StrNoOutDir; return }
        $script:OutDirBox.Text = $outDir

        $nodeArgs = @(
            (Join-Path $script:Root 'build.js'),
            ('--books=' + ($dirs -join ',')),
            ('--order=' + (Get-OrderName)),
            ('--out=' + $outDir)
        )

        if ($script:MergeCheck.IsChecked) {
            $merged = $script:MergedNameBox.Text.Trim()
            if (-not $merged) { $merged = 'Arabic_Workbooks_Complete.pdf' }
            if ($merged -match '[\\/:*?"<>|]') { Show-Warning $script:StrBadMerged; return }
            if (-not $merged.ToLower().EndsWith('.pdf')) { $merged = $merged + '.pdf' }
            $script:MergedNameBox.Text = $merged
            $nodeArgs += '--merge'
            $nodeArgs += ('--merged-name=' + $merged)
        }

        Start-JobRun -NodeArgs $nodeArgs -Kind 'books'
    })

    $script:ImagesBuildButton.Add_Click({
        if ($script:IsBuilding) { return }

        $images = $script:ImagesDirBox.Text.Trim()
        if (-not $images -or -not (Test-Path -LiteralPath $images)) {
            Show-Warning $script:StrNoImages
            return
        }
        $outDir = Resolve-OutputFolder $script:ImagesOutBox.Text.Trim()
        if (-not $outDir) { Show-Warning $script:StrNoOutDir; return }
        $script:ImagesOutBox.Text = $outDir

        $size = [string]$script:SizeBox.SelectedItem
        $nodeArgs = @(
            (Join-Path $script:Root 'image_2_cover_style\images-to-book.js'),
            '--images', $images,
            '--out', $outDir,
            '--size', $size.ToLower()
        )
        if ($script:ImagesDocxOnly.IsChecked) { $nodeArgs += '--no-pdf' }

        Start-JobRun -NodeArgs $nodeArgs -Kind 'images'
    })

    $script:Window.Add_Closing({
        if ($script:Process -and -not $script:Process.HasExited) { Stop-JobRun }
        if ($script:LogPath -and (Test-Path -LiteralPath $script:LogPath)) {
            Remove-Item -LiteralPath $script:LogPath -Force -ErrorAction SilentlyContinue
        }
    })

    [void]$script:Window.ShowDialog()
}
catch {
    [void][System.Windows.MessageBox]::Show(
        ($_.Exception.Message + [Environment]::NewLine + [Environment]::NewLine + $_.ScriptStackTrace),
        'workbooks-gui.ps1',
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error)
}
