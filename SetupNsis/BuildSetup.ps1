$AppVersion = (Get-ChildItem "..\LyncFellow\bin\Release\LyncFellow.exe").VersionInfo.FileVersion

Get-Content ".\LyncFellow.nsi" | 
Foreach-Object { $_ -replace "0.8.1.5", $AppVersion } |
Set-Content ".\LyncFellowVersion.nsi"

Remove-Item ".\LyncFellowSetup*"
& "C:\Program Files (x86)\NSIS\makensis.exe" "LyncFellowVersion.nsi"
& "C:\Program Files\7-Zip\7z.exe" "a" ("LyncFellowSetup " + $AppVersion + ".zip") ("LyncFellowSetup " + $AppVersion + ".exe")
Remove-Item ".\LyncFellowVersion.nsi"

Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
