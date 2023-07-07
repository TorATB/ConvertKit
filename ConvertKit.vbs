Set WshShell = WScript.CreateObject("WScript.Shell")
Set lnk = WshShell.CreateShortcut(WshShell.SpecialFolders("SendTo") & "\ConvertKit.lnk")
WinPath = WshShell.ExpandEnvironmentStrings("%windir%")
AppPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
AppPath = mid(AppPath,1,len(AppPath) -1)

lnk.targetpath = chr(34) & WinPath & "\system32\mshta.exe" & chr(34)
lnk.arguments = chr(34) & AppPath & "\ConvertKit.hta" & chr(34)
lnk.description = "ConvertKit"
lnk.workingdirectory = AppPath
lnk.iconlocation = AppPath & "\ConvertKit.ico, 0"
lnk.save


'Setting My Values
myDir = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "")
Dim ApplicationName, ApplicationVersion, ApplicationIconPath, ApplicationPublisher, ApplicationUnInstallPath, ApplicationInstallDirectory
AppName = "ConvertKit"
AppVersion = "1.3.0.0"
AppIconPath = myDir & "\ConvertKit.ico"
AppPublisher = "TorATB"
AppUnInstallPath = myDir& "\UNINSTALL.bat"
AppInstallDirectory = myDir & "\"

WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "DisplayName", AppName, "REG_SZ"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "DisplayVersion", AppVersion, "REG_SZ"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "DisplayIcon", AppIconPath, "REG_SZ"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "Publisher", AppPublisher, "REG_SZ"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "UninstallString", AppUnInstallPath, "REG_SZ"
WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & AppName & "\" & "InstallLocation", AppInstallDirectory, "REG_SZ"
