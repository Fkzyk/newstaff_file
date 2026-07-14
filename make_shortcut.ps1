# Create/refresh a Desktop shortcut for the app (idempotent)
$ws = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$lnk = $ws.CreateShortcut((Join-Path $desktop '入社書類アプリ.lnk'))
$lnk.TargetPath = (Join-Path $PSScriptRoot 'start.bat')
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.IconLocation = 'imageres.dll,187'
$lnk.Description = '入社書類作成アプリを起動'
$lnk.Save()
