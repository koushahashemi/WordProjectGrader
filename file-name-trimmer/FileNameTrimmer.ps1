param([switch]$SelfTest, [switch]$GuiSmokeTest)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-TrimPlan {
    param([string]$Folder, [bool]$Recursive = $false)
    $option = [IO.SearchOption]::TopDirectoryOnly
    if ($Recursive) { $option = [IO.SearchOption]::AllDirectories }
    $files = [IO.Directory]::GetFiles($Folder, '*', $option) |
        Where-Object { [IO.Path]::GetExtension($_) -ieq '.docx' } | Sort-Object
    $plans = @()
    $targets = @{}
    foreach ($source in $files) {
        $name = [IO.Path]::GetFileName($source)
        $index = $name.IndexOf('_')
        $status = 'Ready'
        $newName = $name
        if ($name.StartsWith('~$')) { $status = 'Skipped: Word temporary file' }
        elseif ($index -lt 0) { $status = 'Skipped: no underscore' }
        else {
            $newName = $name.Substring($index + 1)
            if ($newName -ieq '.docx') { $status = 'Skipped: empty name before extension' }
            elseif ($newName.EndsWith(' ') -or $newName.StartsWith(' ')) { $status = 'Skipped: resulting name starts or ends with a space' }
            elseif ($newName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { $status = 'Skipped: reserved Windows filename' }
        }
        $target = [IO.Path]::Combine([IO.Path]::GetDirectoryName($source), $newName)
        if ($status -eq 'Ready' -and (Test-Path -LiteralPath $target)) { $status = 'Skipped: destination already exists' }
        $plan = [pscustomobject]@{ Source = $source; Target = $target; Name = $name; NewName = $newName; Status = $status }
        $plans += $plan
        if ($status -eq 'Ready') {
            if (-not $targets.ContainsKey($target)) { $targets[$target] = @() }
            $targets[$target] += $plan
        }
    }
    foreach ($group in $targets.Values) {
        if ($group.Count -gt 1) { foreach ($plan in $group) { $plan.Status = 'Skipped: multiple files would have this name' } }
    }
    return $plans
}

function Invoke-TrimPlan {
    param([object[]]$Plans)
    foreach ($plan in $Plans) {
        if ($plan.Status -ne 'Ready') { continue }
        try {
            # File.Move never overwrites: also protects against files created after preview.
            [IO.File]::Move($plan.Source, $plan.Target)
            $plan.Status = 'Renamed'
        } catch { $plan.Status = 'Failed: ' + $_.Exception.Message }
    }
}

if ($SelfTest) {
    $folder = Join-Path ([IO.Path]::GetTempPath()) ('file-name-trimmer-' + [guid]::NewGuid())
    [IO.Directory]::CreateDirectory($folder) | Out-Null
    try {
        foreach ($name in @('123_Alice_Bears.docx','no-underscore.docx','123_.docx','123_CON.docx','first_same.docx','second_same.docx','existing.docx','123_existing.docx','123_notes.txt','_Leading.DOCX','~$123_Locked.docx','123_Zoe Lee_Bears.docx')) {
            [IO.File]::WriteAllText((Join-Path $folder $name), $name)
        }
        $sub = Join-Path $folder 'nested'; [IO.Directory]::CreateDirectory($sub) | Out-Null
        [IO.File]::WriteAllText((Join-Path $sub '456_Child.docx'), 'nested')
        $plans = @(Get-TrimPlan $folder)
        if ($plans.Count -ne 11) { throw 'Expected only DOCX files in chosen folder' }
        if (@($plans | Where-Object Status -eq 'Ready').Count -ne 3) { throw 'Expected three safe renames' }
        if (@($plans | Where-Object Status -like '*multiple files*').Count -ne 2) { throw 'Duplicate targets were not blocked' }
        Invoke-TrimPlan $plans
        if ([IO.File]::ReadAllText((Join-Path $folder 'Alice_Bears.docx')) -ne '123_Alice_Bears.docx') { throw 'Content or first-underscore rule changed' }
        if (-not (Test-Path -LiteralPath (Join-Path $folder 'Leading.DOCX'))) { throw 'Leading underscore or uppercase extension failed' }
        if (-not (Test-Path -LiteralPath (Join-Path $folder 'Zoe Lee_Bears.docx'))) { throw 'Names with spaces failed' }
        if ([IO.File]::ReadAllText((Join-Path $folder 'existing.docx')) -ne 'existing.docx') { throw 'Existing target was overwritten' }
        if (-not (Test-Path -LiteralPath (Join-Path $sub '456_Child.docx'))) { throw 'Unexpected recursive rename' }
        $recursive = @(Get-TrimPlan $folder $true)
        if (-not ($recursive | Where-Object NewName -eq 'Child.docx')) { throw 'Optional recursion failed' }
        [IO.File]::WriteAllText((Join-Path $folder 'race_Target.docx'), 'source')
        $race = @(Get-TrimPlan $folder | Where-Object Name -eq 'race_Target.docx')
        [IO.File]::WriteAllText((Join-Path $folder 'Target.docx'), 'keep')
        Invoke-TrimPlan $race
        if ($race[0].Status -notlike 'Failed:*' -or [IO.File]::ReadAllText((Join-Path $folder 'Target.docx')) -ne 'keep') { throw 'Late collision was not protected' }
        Write-Output 'File Name Trimmer self-test passed.'
    } finally { [IO.Directory]::Delete($folder, $true) }
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object Windows.Forms.Form
$form.Text = 'File Name Trimmer 1.0'
$form.Size = New-Object Drawing.Size(1000, 650)
$form.MinimumSize = New-Object Drawing.Size(800, 500)
$form.StartPosition = 'CenterScreen'
$layout = New-Object Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'; $layout.Padding = New-Object Windows.Forms.Padding(12)
$layout.ColumnCount = 1; $layout.RowCount = 5
foreach ($height in @(50, 38, 35)) { $layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', $height))) | Out-Null }
$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100))) | Out-Null
$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 45))) | Out-Null
$form.Controls.Add($layout)
$intro = New-Object Windows.Forms.Label
$intro.Text = 'Remove everything through the FIRST underscore from .docx filenames. Example: 12345_John_Bears.docx becomes John_Bears.docx. Preview first, then rename. Existing files are never overwritten.'
$intro.Dock = 'Fill'; $layout.Controls.Add($intro, 0, 0)
$folderRow = New-Object Windows.Forms.TableLayoutPanel
$folderRow.Dock = 'Fill'; $folderRow.ColumnCount = 2
$folderRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$folderRow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute', 140))) | Out-Null
$folderBox = New-Object Windows.Forms.TextBox
$folderBox.Dock = 'Fill'; $folderBox.ReadOnly = $true
$browse = New-Object Windows.Forms.Button
$browse.Text = 'Choose folder...'; $browse.Dock = 'Fill'
$folderRow.Controls.Add($folderBox, 0, 0); $folderRow.Controls.Add($browse, 1, 0)
$layout.Controls.Add($folderRow, 0, 1)
$options = New-Object Windows.Forms.FlowLayoutPanel
$options.Dock = 'Fill'
$recursive = New-Object Windows.Forms.CheckBox
$recursive.Text = 'Include subfolders'; $recursive.AutoSize = $true
$preview = New-Object Windows.Forms.Button
$preview.Text = 'Refresh preview'; $preview.AutoSize = $true
$options.Controls.Add($recursive); $options.Controls.Add($preview); $layout.Controls.Add($options, 0, 2)
$list = New-Object Windows.Forms.ListView
$list.View = 'Details'; $list.FullRowSelect = $true; $list.GridLines = $true; $list.Dock = 'Fill'
foreach ($column in @(@('Current name',220),@('New name',220),@('Status',250),@('Folder',300))) { $list.Columns.Add($column[0], $column[1]) | Out-Null }
$layout.Controls.Add($list, 0, 3)
$bottom = New-Object Windows.Forms.FlowLayoutPanel
$bottom.Dock = 'Fill'
$rename = New-Object Windows.Forms.Button
$rename.Text = 'Rename ready files'; $rename.AutoSize = $true; $rename.Enabled = $false
$status = New-Object Windows.Forms.Label
$status.Text = 'Choose a folder to preview its .docx files.'; $status.AutoSize = $true; $status.Padding = New-Object Windows.Forms.Padding(8)
$bottom.Controls.Add($rename); $bottom.Controls.Add($status); $layout.Controls.Add($bottom, 0, 4)
$script:plans = @()
function Show-Plan {
    $list.Items.Clear()
    foreach ($plan in $script:plans) {
        $item = New-Object Windows.Forms.ListViewItem($plan.Name)
        $item.SubItems.Add($plan.NewName) | Out-Null
        $item.SubItems.Add($plan.Status) | Out-Null
        $item.SubItems.Add([IO.Path]::GetDirectoryName($plan.Source)) | Out-Null
        $list.Items.Add($item) | Out-Null
    }
    $ready = @($script:plans | Where-Object Status -eq 'Ready').Count
    $done = @($script:plans | Where-Object Status -eq 'Renamed').Count
    $status.Text = "$($script:plans.Count) files; $ready ready; $done renamed."
    $rename.Enabled = $ready -gt 0
}
function Update-Preview {
    $rename.Enabled = $false; $script:plans = @(); $list.Items.Clear()
    if (-not $folderBox.Text) { return }
    try { $script:plans = @(Get-TrimPlan $folderBox.Text $recursive.Checked); Show-Plan }
    catch { $status.Text = 'Could not read folder.'; [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Folder error') | Out-Null }
}
$browse.Add_Click({
    $dialog = New-Object Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose the folder containing your Word documents'
    if ($dialog.ShowDialog() -eq 'OK') { $folderBox.Text = $dialog.SelectedPath; Update-Preview }
    $dialog.Dispose()
})
$preview.Add_Click({ Update-Preview })
$recursive.Add_CheckedChanged({ Update-Preview })
$rename.Add_Click({
    $rename.Enabled = $false
    Invoke-TrimPlan $script:plans
    Show-Plan
})
if ($GuiSmokeTest) {
    $form.Show()
    [Windows.Forms.Application]::DoEvents()
    $form.Close()
    $form.Dispose()
    Write-Output 'File Name Trimmer UI self-test passed.'
    exit 0
}
[Windows.Forms.Application]::Run($form)
$form.Dispose()
