Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
Set objFileOut = objFSO.OpenTextFile(SysDrive & "\backedup_shares\quotas_delete.bat", ForWriting, True)
'Set objFileOut1 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\drives_sn.list", ForWriting, True)
objFileOut.Write("@echo off" & vbCrLf)

quotasexist=0

'Получаем список квот
oShell.run "cmd /c ""dirquota q l > " & SysDrive & "\backedup_shares\all_quotas.txt""",0,bWaitOnReturn

'Если квоты есть, то будут выгружаться щаблоны и генерироваться командный файл для удаления квот
Set objFileIn = objFSO.OpenTextFile(SysDrive & "\backedup_shares\all_quotas.txt", ForReading)
Do Until objFileIn.AtEndOfStream
	tmpline = objFileIn.ReadLine
	If (Len(tmpline) > 0) Then
		If (InStr(tmpline, "Quota Path:") = 1) Then quota_path = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Path:")))
		If quota_path <> "" Then
			quotasexist=1
		End If
	End If
Loop

If objFSO.FileExists(SysDrive & "\Windows\System32\dirquota.exe") And (quotasexist = 1) Then
	'Экспортируем шаблоны
	oShell.run "dirquota template export /file:" & SysDrive & "\backedup_shares\quota_templates.xml",0,bWaitOnReturn
	
	'Генерируем командный файл для удаления квот
	For Each objDrive in colDrives
		If objDrive.DriveType = 2 Then
			objFileOut.Write("dirquota quota delete /Path:" & objDrive.DriveLetter & ":\* /Quiet" & vbCrLf)
			'objFileOut1.Write(objDrive.DriveLetter & ": " & objDrive.SerialNumber & vbCrLf)
		End If
	Next
Else
	Set objFileQtemplates = objFSO.OpenTextFile(SysDrive & "\backedup_shares\quota_templates.xml", ForWriting, True)
	objFileQtemplates.Close
End If

objFileOut.Close
objFileIn.Close
'objFileOut1.Close
