$TargetFile = "C:\Scripts\Change2.bat"
$ShortcutFile = "$env:Public\Desktop\Event Log.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.IconLocation = "C:\Windows\System32\SHELL32.dll,26"
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()