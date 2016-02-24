'Script to create shortcut of IE
'Author: Margulan Tukhfatov
'Modified: 11/19/2015

set networkObj = CreateObject ("Wscript.Network")
set shellObj = CreateObject("Wscript.Shell")

userRoot = shellObj.ExpandEnvironmentStrings("%UserProfile%") & "\"

set ieShortcut = shellObj.CreateShortcut(userRoot+"Desktop\Internet Explorer.lnk")
ieShortcut.WindowStyle = 3
ieShortcut.TargetPath = "C:\Program Files\Internet Explorer\iexplore.exe"
ieShortcut.Save
msgbox "Link successfully created"

WScript.Quit
