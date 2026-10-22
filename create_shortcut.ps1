$WshShell = New-Object -ComObject WScript.Shell
$Desktop = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$Shortcut = $WshShell.CreateShortcut("$Desktop\GPS Qala Qalagay School.lnk")
$Shortcut.TargetPath = "$PSScriptRoot\gps_run.bat"
$Shortcut.WorkingDirectory = "$PSScriptRoot"
$Shortcut.Description = "GPS Qala Qalagay School ERP & Examination Portal"
$Shortcut.Save()
Write-Host "[OK] Shortcut 'GPS Qala Qalagay School' created successfully on Desktop!"
