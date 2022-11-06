'//*****************************************************************************
'// NOTE: I DO NOT RECOMMEND THAT YOU HARD-CODE YOUR USER NAME AND PASSWORD
'//       I only list them here as a convenience while testing a script
'//       You should retrieve these values from a prompt to the user who is
'//       running the script.
'//*****************************************************************************
'// Constants used - Change these for your specific environment
Const QUALYSUSR              = "qualysusername"
Const QUALYSPW               = "qualysuserpassword"
Const QUALYSAPIURL           = "https://qualysapi.qualys.com/"
Const adTypeText             = 2
Const adSaveCreateOverWrite  = 2

Dim oShell, oFS, AppPath, DataPath, OnLine

'// Global objects used throughout the script
Set oShell = CreateObject("WScript.Shell")
Set oFS = CreateObject("Scripting.FileSystemObject")

'// Global variables used throughout the script
'- AppPath stores the location of the current directory the script is running from
'    All paths start with AppPath as the root folder
AppPath = replace(oShell.CurrentDirectory+"\","\\","\")

'- DataPath is used to store output of API calls
DataPath = AppPath & "Data\"
VerifyFolder DataPath

'// Assume that Internet activity is avaialable, and check otherwise
OnLine=True
If Not LCase(chkInternet("http://www.qualys.com"))="ok" Then OnLine=False

If OnLine Then sessionLogin

'//*****************************************************************************
'//
'// Perform the tasks you need to perform here
'//
'//*****************************************************************************

If OnLine Then sessionLogout


Sub VerifyFolder(strFolder)
'//*****************************************************************************
'// Description: Creates folders for provided path.
'//
'// Input: Full drive letter and path (i.e.: C:\Folder1\Folder2\Folder3).
'//
'// Output: None.
'//
'//*****************************************************************************
	Dim strPath, i, arrTMP

	strPath = strFolder
	If Right(strPath,1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
	If Not oFS.FolderExists(strPath) Then
		arrTMP = Split(strPath, "\")
		strPath = arrTMP(0)
		For i = 1 To UBound(arrTMP)
			strPath = strPath & "\" & arrTMP(i)
			on error resume next
			If Not oFS.FolderExists(strPath) Then oFS.CreateFolder(strPath)
			on error goto 0
		Next
	End If
End Sub

Function chkInternet(sHREF)
'//*****************************************************************************
'// Description: Verify Internet connectivity.
'//
'// Input: External website.
'//
'// Output: Connectivity status.
'//
'//*****************************************************************************
	Dim oXMLHTTP

	Set oXMLHTTP = CreateObject("msxml2.xmlhttp.6.0")
	oXMLHTTP.Open "GET", sHREF, false
	oXMLHTTP.setRequestHeader "Content-type:", "text/xml"
	oXMLHTTP.setRequestHeader "Translate:", "f"
	On Error Resume Next
	oXMLHTTP.Send
	On Error GoTo 0
	If err.Number=0 Then
		chkInternet = oXMLHTTP.statustext
	Else
		chkInternet="Forbidden"
	End If
	Set oXMLHTTP = Nothing
End Function

Sub sessionLogin()
'//*****************************************************************************
'// Description: Performs Qualys login via APIv2.
'//
'// Input: None.
'//
'// Output: Session login status in "login.xml" file.
'//
'//*****************************************************************************
	getXMLFile 2, "login.xml", "session/?action=login&username=" & QUALYSUSR & "&password=" & QUALYSPW, "POST"
End Sub

Sub sessionLogout()
'//*****************************************************************************
'// Description: Performs Qualys Session logout via APIv2.
'//
'// Input: None.
'//
'// Output: Session logout status in "logout.xml" file.
'//
'//*****************************************************************************
	getXMLFile 2, "logout.xml", "session/?action=logout", "POST"
End Sub

Sub getXMLFile(ver, xmlFile, phpLine, Method)
'//*****************************************************************************
'// Description: Used for all API calls. Performs the requeted API task.
'//
'// Input: ver     - API version (1 or 2).
'//        xmlFile - File name to store the output of the API call.
'//        phpLine - String specifying the API arguments for the task.
'//        Method  - Either "GET" or "POST"
'//
'// Output: Results are stored in the specified file (xmlFile).
'//
'//*****************************************************************************
	Dim sHREF, oStream
	Dim oXMLHTTP

	If OnLine Then
		Set oXMLHTTP = CreateObject("msxml2.xmlhttp.6.0")
		If ver=1 Then
			sHREF=QUALYSAPIURL & "msp/" & phpLine
			oXMLHTTP.Open Method, sHREF, false, QUALYSUSR, QUALYSPW
		Else
			sHREF=QUALYSAPIURL & "api/2.0/fo/" & phpLine
			oXMLHTTP.Open Method, sHREF, false
			oXMLHTTP.setRequestHeader "X-Requested-With:", "VBScript"
		End If
		set oStream = createobject("adodb.stream")
		oStream.type = adTypeText
		oStream.Charset="ascii"
		oStream.open
		oXMLHTTP.setRequestHeader "Content-type:", "text/xml"
		oXMLHTTP.setRequestHeader "Translate:", "f"
		oXMLHTTP.Send
		oStream.writetext oXMLHTTP.ResponseText
		oStream.Position=0
		oStream.savetofile DataPath & xmlFile, adSaveCreateOverWrite
		Set oXMLHTTP = Nothing
		Set oStream = Nothing
	End If
End Sub
