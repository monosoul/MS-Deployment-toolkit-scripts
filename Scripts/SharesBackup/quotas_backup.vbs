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

If objFSO.FileExists(SysDrive & "\Windows\System32\dirquota.exe") Then
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

'Получаем список квот
oShell.run "cmd /c ""dirquota q l > " & SysDrive & "\backedup_shares\all_quotas.txt""",0,bWaitOnReturn

objFileOut.Close
'objFileOut1.Close
