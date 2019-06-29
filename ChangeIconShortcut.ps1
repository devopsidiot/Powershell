$Shell = New-Object -ComObject ("WScript.Shell")
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.IconLocation = "$IconLocation, $IconArrayIndex"
$Shortcut.Save()