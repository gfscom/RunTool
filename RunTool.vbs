' RunTool 0.1
' GSolone - Ultima modifica: 14/5/19
'
' Lo script scarica file di testo da GitHub, i quali contengono codice da eseguire via prompt dei comandi. Il file di testo viene trasformato in un CMD da eseguire sulla macchina locale.
'
' STORICO MODIFICHE
' 0.1- versione iniziale dello script.
'
' Sviluppo: 	Giovanni F. -Gioxx- Solone (dev@gfsolone.com)
' Testato su:	Windows 10 Pro
' Thanks to:	Rob van der Woude (http://www.robvanderwoude.com)
'				http://www.thaivisa.com/forum/index.php?showtopic=21832
'
Dim TestoInputBox,inp01,strFlag
strFlag = False

TestoInputBox = "Codice script da scaricare (https://go.gioxx.org/runtooldir)"
TestoInputBox = TestoInputBox & vbCrLf & "0 - esci dallo script"

Do While strFlag = False
	inp01 = InputBox(TestoInputBox,"RunTool")
	Select Case inp01
		' Exit from the script
		Case "0"
			Wscript.Quit
			strFlag = True
		Case Else
		' Download the script and create RunTool.cmd
			GitHub = "https://raw.githubusercontent.com/gioxx/RunTool/master/script/" & inp01 & ".txt"
			HTTPDownload GitHub, ".\RunTool.cmd"
			OpenCMD ".\RunTool.cmd"
			DelCMD ".\RunTool.cmd"
			strFlag = True
	End Select
loop
Wscript.Quit

Sub HTTPDownload( myURL, myPath )
	' Original script: 	Rob van der Woude (http://www.robvanderwoude.com)
	' Modified:			Giovanni F. 'Gioxx' Solone (Connection through proxy integration) (https://gioxx.org)
    ' Standard housekeeping
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    ' Create a File System Object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    ' Check if the specified target file or folder exists,
    ' and build the fully qualified path of the target file
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If
    ' Create or open the target file
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )
    ' Create an HTTP object
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
	' Connection through proxy (remove comment in the line below and change proxy address / port)
	'objHTTP.setProxy 2, "proxy.contoso.com:8080", ""
    ' Download the specified URL
    objHTTP.Open "GET", myURL, False
    objHTTP.Send
    ' Write the downloaded byte stream to the target file
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next
    ' Close the target file
    objFile.Close( )
End Sub

Sub OpenCMD( myPath )
	Set WshShell = CreateObject("WScript.Shell")
	answer= MsgBox ("Eseguo adesso lo script scaricato?", vbYesNo + vbQuestion, "RunTool.cmd")
	if answer= vbYes then
		WshShell.Run myPath
	end if
End Sub

Sub DelCMD( myPath )
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )
	answer= MsgBox ("Posso rimuovere lo script scaricato?", vbYesNo + vbQuestion, "RunTool.cmd")
	if answer= vbYes then
		objFSO.DeleteFile(myPath)
		Set objFSO = Nothing
	end if
End Sub