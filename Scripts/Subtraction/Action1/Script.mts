Call AddLibrary ("Operations.vbs")
Call subtraction ("4", "2", difference)
Call AddLibrary ("Print.vbs")
Call printOperationResult ("Difference", difference)

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
