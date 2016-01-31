'Script to automatically connect all pst files
'Before type in cmd: dir /b /s | find "pst" > text.txt
'Create a vbs file in cmd: type null > pst.vbs
'Copy all codes into vbs file
'Change the path to text file
'Run script
'Author: Margulan Tukhfatov

Set objOutlook = CreateObject("Outlook.Application")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject ("Scripting.FileSystemObject")
objFilename = ""

function userPrompt()
	Do While X = 0
	    strAnswer = InputBox _
	        ("Please enter a path:","Path to txt file")
	    If strAnswer = "" Then
	        Wscript.Echo "Path not found"
	    Else
	        Wscript.Echo strAnswer
	        objFilename = strAnswer
	        Exit Do
	    End If
	Loop
end function

userPrompt()
if objFSO.fileExists(objFilename) then
	Set objFile = objFSO.OpenTextFile(objFilename, 1)
	Do While objFile.AtEndOfStream = False
	    strLine = objFile.ReadLine
	    if Right(strLine, 3) = "pst" then
		    objOutlook.Session.Addstore strLine
		    Wscript.Echo "Successfully added: " & strLine
		else
			Wscript.Echo "Not Archive file: " & strLine 
		end if
	Loop
	Wscript.Echo "Completed"
else
	Wscript.Echo "File Not Found"
end if
