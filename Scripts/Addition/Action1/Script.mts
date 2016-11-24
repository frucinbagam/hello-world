Call AddLibrary ("Operations.vbs")
Call addition ("1", "2", sum)
Call AddLibrary ("Print.vbs")
Call printOperationResult ("Sum", sum)

Public Function AddLibrary(LibName)
	LibPath = "D:\Automation\POC\UFT-Git-Demo\Functions\"
	fPath = LibPath & LibName
	Set qtpApp = CreateObject("QuickTest.Application")
	If qtpApp.Test.Settings.Resources.Libraries.Find(fpath) = -1 Then
		On Error Resume Next
		Executefile (fPath)
		If StrCompt(Err.Description, "Name redefined", 1) = 0 Then
			Err.Clear
			Exit Function
		End If
	End If
End Function
