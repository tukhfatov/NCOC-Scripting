' Merge *.xlsx files
' Author: Margulan Tukhfatov

Set objExcel = CreateObject("Excel.Application")
Set objShellApp = CreateObject("Shell.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")

strPathSrc = WshShell.CurrentDirectory
strMaskSrc = "*.xlsx" 
iSheetSrc = 1 
strPathDst = strPathSrc & "\output.xlsx" 
iSheetDst = 1 

objExcel.Visible = True
Set objWorkBookDst = objExcel.Workbooks.Add()
objWorkBookDst.SaveAs(strPathDst)


Set objFolder = objShellApp.NameSpace(strPathSrc)
Set objItems = objFolder.Items()
objItems.Filter 64 + 128, strMaskSrc
objExcel.DisplayAlerts = False

For Each objItem In objItems
	if not objItem.Name = "output.xlsx" then
		SheetName = Left(objItem.Name, (Len(objItem.Name) - 5))

		if Len(SheetName) > 31 then
			SheetName = Left(SheetName, 15)
		end if

	    Set objWorkBookSrc = objExcel.Workbooks.Open(objItem.Path)
	    Set objSheetSrc = objWorkBookSrc.Sheets(1)

	    for each imgObj in objSheetSrc.Shapes
	    	imgObj.delete
	    next
	    
	    objSheetSrc.Name = SheetName
	    objSheetSrc.Copy objWorkBookDst.Sheets(iSheetDst)
	    iSheetDst = iSheetDst +1
	    objWorkBookSrc.Close
	end if
Next
objWorkBookDst.Save()
