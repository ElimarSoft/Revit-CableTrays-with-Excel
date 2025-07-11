Const MacroName1 = "RevitAuto"
Const MacroName2 = "EndPoints"
Const MacroName3 = "SendDataMod"

const MacroExt = ".bas"
Const MacroPath = "C:\Users\micro\Documents\"
Const WorkbookName = "RevitMacros02"

function getLastName (directoryPath, baseName, extName)

	Dim fso, folder, file, version, maxVer, verNum, finalName
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(directoryPath)

	maxVer = 0
	version = ""
	verNum = 0

	For Each file In folder.Files
		If Left(file.name,len(baseName))=baseName and Right(file.name,len(extName))=extName Then
			version = Mid(file.name, len(baseName) + 1, len(file.name)-len(baseName)-len(extName))
			if version <> "" then verNum = cint(version)
			if verNum > maxVer then maxVer = verNum
		End If
	Next

	getLastName = directoryPath+baseName+cstr(maxVer)+extName

	Set file = Nothing
	Set folder = Nothing
	Set fso = Nothing

end function

finalPath1 = getLastName(MacroPath,MacroName1,MacroExt)
finalPath2 = getLastName(MacroPath,MacroName2,MacroExt)
finalPath3 = getLastName(MacroPath,MacroName3,MacroExt)

WindowTitle = WorkbookName + " - Excel"

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

If objShell.AppActivate(WindowTitle) Then

	WScript.Sleep 1000
    objShell.SendKeys "%{F11}"
	
	WScript.sleep 200
    objShell.SendKeys "^m"
	WScript.sleep 1000
    objShell.SendKeys finalPath1 + "{ENTER}"
    
	WScript.Sleep 800    
	objShell.SendKeys "^m"
	WScript.sleep 1000
    objShell.SendKeys finalPath2 + "{ENTER}"
	
	WScript.Sleep 800    
	objShell.SendKeys "^m"
	WScript.sleep 1000
    objShell.SendKeys finalPath3 + "{ENTER}"

	WScript.Sleep 800    
	objShell.SendKeys "%q"
	WScript.Sleep 1000
    objShell.SendKeys "%{F8}"
	WScript.Sleep 1000
	objShell.SendKeys "AddToContextMenu"
	WScript.sleep 200
    objShell.SendKeys "{ENTER}"
	WScript.sleep 200
	
End If

