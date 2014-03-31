Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileIn = objFSO.OpenTextFile(SysDrive & "\backedup_shares\all_quotas.txt", ForReading)
Set objFileOut = objFSO.OpenTextFile(SysDrive & "\backedup_shares\quotas_create.bat", ForWriting, True)
objFileOut.Write("@echo off" & vbCrLf)

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = "\((.*?)\)"

'Импортируем шаблоны
objFileOut.Write("dirquota template import /file:" & SysDrive & "\backedup_shares\quota_templates.xml" & vbCrLf)
'oShell.run "dirquota template import /file:" & SysDrive & "\backedup_shares\quota_templates.xml",0,bWaitOnReturn

Do Until objFileIn.AtEndOfStream
	ProcessQuota()
Loop

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
		quota_limit = CStr(CInt(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "MB"
	ElseIf (UCase(Right(quota_limit, 2)) = "MB") Then
		quota_limit = CStr(CInt(CDbl(Left(quota_limit, Len(quota_limit) - 2)) * 1024)) & "KB"
	ElseIf (UCase(Right(quota_limit, 2)) = "KB") Then
		quota_limit = CStr(CInt(Left(quota_limit, Len(quota_limit) - 2))) & "KB"
	End If
	
	'Проверяем, не изменилась ли буква системного диска и если изменилась - подставляем новую
	If (UCase(left(quota_path, 2)) = WScript.Arguments(0)) And (WScript.Arguments(0) <> SysDrive) Then
		quota_path = SysDrive & Right(quota_path, Len(quota_path) - Len("C:"))
	End If
	
	' Проверяем, не изменилась ли буква диска после развёртывания
	' Типы дисков:
	' 0 - Unknown
	' 1 - Removable drive
	' 2 - Fixed drive
	' 3 - Network drive
	' 4 - CD-ROM drive
	' 5 - RAM Disk
	DrvLetter = objFSO.GetDrive(objFSO.GetDriveName(quota_path)).DriveLetter
	If objFSO.GetDrive(objFSO.GetDriveName(quota_path)).DriveType <> 2 Then
		Do
			DrvNumber = Asc(UCase(DrvLetter)) - 1
			If objFSO.GetDrive(objFSO.GetDriveName(Chr(DrvNumber) & ":\")).DriveType = 2 Then
				DrvLetter = Chr(DrvNumber)
				suggested_quota_path = DrvLetter & Right(quota_path, Len(quota_path) - Len("C"))
			End If
		Loop Until ((objFSO.GetDrive(objFSO.GetDriveName(DrvLetter & ":\")).DriveType = 2) And objFSO.FolderExists(suggested_quota_path)) Or (DrvNumber = 65) ' Перебираем буквы дисков, пока не попадётся жётский диск, на котором существует нужный каталог или, если каталог так и не был найден, пока не доберёмся до диска A (Код - 65)
		If (DrvNumber <> 65) Then
			quota_path = suggested_quota_path
		End If
	End If
	
	If source_template = "None" Then
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & "" & vbCrLf)
		'oShell.run "dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & "",0,bWaitOnReturn
	Else 'Если квота была создана на основе шаблона, то используем его снова
		objFileOut.Write("dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & " /SourceTemplate:" & source_template & "" & vbCrLf)
		'oShell.run "dirquota quota add /Overwrite /Path:""" & quota_path & """ /Limit:" & quota_limit & " /Type:" & quota_type & " /Status:" & quota_status & " /SourceTemplate:" & source_template & "",0,bWaitOnReturn
	End If
	
	End If
End Sub