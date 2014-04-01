Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
Set objFileIn = objFSO.OpenTextFile(SysDrive & "\backedup_shares\all_quotas.txt", ForReading)
Set objFileOut = objFSO.OpenTextFile(SysDrive & "\backedup_shares\quotas_create.bat", ForWriting, True)
objFileOut.Write("@echo off" & vbCrLf)

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = "\((.*?)\)"

'Импортируем шаблоны
objFileOut.Write("dirquota template import /file:" & SysDrive & "\backedup_shares\quota_templates.xml" & vbCrLf)

Do Until objFileIn.AtEndOfStream
	ProcessQuota()
Loop

objFileOut.Close
objFileIn.Close

Sub ProcessQuota()
	
	quota_path = ""
	quota_status = ""
	quota_limit = ""
	quota_type = ""
	source_template = ""
	
	Do Until objFileIn.AtEndOfStream Or ((Len(quota_path) > 0) And (Len(quota_status) > 0) And (Len(quota_limit) > 0))
		tmpline = objFileIn.ReadLine
		If (Len(tmpline) > 0) Then
			If (InStr(tmpline, "Quota Path:") = 1) Then quota_path = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Path:")))
			If (InStr(tmpline, "Source Template:") = 1) Then source_template = LTrim(Right(tmpline, Len(tmpline) - Len("Source Template:")))
			If (InStr(tmpline, "Quota Status:") = 1) Then quota_status = LTrim(Right(tmpline, Len(tmpline) - Len("Quota Status:")))
			If (InStr(tmpline, "Limit:") = 1) Then quota_limit = LTrim(Right(tmpline, Len(tmpline) - Len("Limit:")))
		End If
	Loop
	
	If quota_path <> "" Then
	
	'Удаляем "(Does not match template)" из имени шаблона, если он содержит эту надпись
	
	If (InStr(source_template, "(Does not match template)") <> 0) Then
		source_template = RTrim(Left(source_template, Len(source_template) - Len("(Does not match template)")))
	End If
	
	'Выдёргиваем тип квоты из строки с лимитом
	Set matches = objRegEx.Execute(quota_limit)
	count = matches.count
	For i = 0 To count - 1
		quota_type=matches(i).submatches(0)
		quota_limit = replace(left(quota_limit, Len(quota_limit) - Len(" (" & quota_type & ")"))," ","")
	Next
	
	'Пересчитываем единицы измерения в единицы на порядок ниже, чтобы избавиться от разделителя в лимите
	delimpos = 0
	If (UCase(Right(quota_limit, 2)) = "GB") Then
		quota_limit = CStr(CLng(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "MB"
	ElseIf (UCase(Right(quota_limit, 2)) = "MB") Then
		quota_limit = CStr(CLng(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "KB"
	ElseIf (UCase(Right(quota_limit, 2)) = "KB") Then
		quota_limit = CStr(CLng(Left(quota_limit, Len(quota_limit) - 2))) & "KB"
	End If
	
	'Проверяем, не изменилась ли буква системного диска и если изменилась - подставляем новую
	If (UCase(left(quota_path, 2)) = WScript.Arguments(0)) And (WScript.Arguments(0) <> SysDrive) Then
		quota_path = SysDrive & Right(quota_path, Len(quota_path) - Len("C:"))
	End If
	
	' Проверяем, не изменилась ли буква диска после развёртывания
	
	drivesn = ""
	OldDrvLetter = Left(quota_path, 2)
	
	For Each rrtmpline in FileToArray(SysDrive & "\backedup_shares\drives_sn.list", False)
		If (InStr(rrtmpline, OldDrvLetter) = 1) Then
			drivesn = Right(rrtmpline, Len(rrtmpline) - Len(OldDrvLetter & " "))
		End If
	Next
	
	If (Len(drivesn) > 0) Then
		For Each objDrive in colDrives
			If objDrive.DriveType = 2 Then
				If (CStr(objDrive.SerialNumber) = drivesn) And (objFSO.FolderExists(objDrive.DriveLetter & Right(quota_path, Len(quota_path) - 1))) Then
					quota_path = objDrive.DriveLetter & Right(quota_path, Len(quota_path) - 1)
				End If
			End If
		Next
	End If
	
	If source_template = "None" Then
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & "" & vbCrLf)
	Else 'Если квота была создана на основе шаблона, то используем его снова
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & " /SourceTemplate:""" & source_template & """" & vbCrLf)
	End If
	
	End If
End Sub

Function FileToArray(ByVal strFile, ByVal blnUNICODE)
  Const FOR_READING = 1
  Dim objFSO, objTS, strContents
  FileToArray = Split("")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(strFile) Then
    On Error Resume Next
    Set objTS = objFSO.OpenTextFile(strFile, FOR_READING, False, blnUNICODE)
    If Err = 0 Then
      strContents = objTS.ReadAll
      objTS.Close
      FileToArray = Split(strContents, vbNewLine)
    End If
  End If
End Function
