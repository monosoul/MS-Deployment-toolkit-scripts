Const ForReading = 1
Const ForWriting = 2

Const FILE_SHARE          = 0 
Const MAXIMUM_CONNECTIONS = 0 

On Error Resume Next

Set WSHShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileIn = objFSO.OpenTextFile("c:\old_c\shares.txt", ForReading)


Do Until objFileIn.AtEndOfStream
	ProcessShare()
Loop


Sub ProcessShare()

share_name = ""
share_path = ""
share_desc = ""

Do Until objFileIn.AtEndOfStream Or ((Len(share_name) > 0) And (Len(share_path) > 0) And (Len(share_desc) > 0))
	tmpline = objFileIn.ReadLine
	If (Len(tmpline) > 0) Then
		If (InStr(tmpline, "Share:") = 1) Then share_name = Right(tmpline, Len(tmpline) - Len("Share:"))
		If (InStr(tmpline, "Path:") = 1) Then share_path = Right(tmpline, Len(tmpline) - Len("Path:"))
		If (InStr(tmpline, "Desc:") = 1) Then share_desc = Right(tmpline, Len(tmpline) - Len("Desc:"))
	End If
Loop ' Считали все три параметра для создания общей папки - создаем

If ((Len(share_name) > 0) And (Len(share_path) > 0) And (Len(share_desc) > 0)) Then

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objNewShare = objWMIService.Get("Win32_Share")
	WScript.Echo ""
	WScript.Echo share_name & " (" & share_path & "):"
	errReturn = objNewShare.Create (share_path, share_name, FILE_SHARE, MAXIMUM_CONNECTIONS, share_desc)

	If errReturn = 0 Then
		' Создали общую папку, чистим дефолтные доступы
		WSHShell.Run "c:\utils\setacl.exe -ot shr -on """ & share_name & """ -actn ace -ace ""n:""ВСЕ"";m:revoke""", 1, True
		WSHShell.Run "c:\utils\setacl.exe -ot shr -on """ & share_name & """ -actn ace -ace ""n:""EVERYONE"";m:revoke""", 1, True

		Do
			trustee = ""
			perm = ""
			Do Until objFileIn.AtEndOfStream Or ((Len(trustee) > 0) And (Len(perm) > 0)) Or (Len(tmpline) < 2)
				tmpline = objFileIn.ReadLine
				If (Len(tmpline) > 0) Then
					If (InStr(tmpline, "Trustee:") = 1) Then trustee = Right(tmpline, Len(tmpline) - Len("Trustee:"))
					If (InStr(tmpline, "Right:") = 1) Then perm = Right(tmpline, Len(tmpline) - Len("Right:"))
				End If
			Loop ' Считали группу доступа и уровень доступа - применяем

			If ((Len(trustee) > 0) And (Len(perm) > 0)) Then
				If (InStr(UCase(trustee), "BUILTIN\") = 1) Then trustee = Right(trustee, Len(trustee) - Len("BUILTIN\"))
				If (perm = "FullControl") Then perm = "full"
				If (perm = "ReadAndExecute") Then perm = "read"
				If (perm = "Modify") Then perm = "change"
				WScript.Echo trustee & " - " & perm 
				WSHShell.Run "c:\utils\setacl -ot shr -on """ & share_name & """ -actn ace -ace ""n:" & trustee & ";p:" & perm & ";m:set""", 1, True
			End If
		Loop Until (Len(tmpline) < 2) Or objFileIn.AtEndOfStream
	Else
		MsgBox "Ошибка создания общих папок!"
	End If

End If

End Sub