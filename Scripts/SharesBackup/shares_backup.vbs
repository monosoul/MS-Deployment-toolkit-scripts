Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

' WMI Constants

Const WBEM_RETURN_IMMEDIATELY = &h10
Const WBEM_FORWARD_ONLY = &h20

' Constants and storage arrays for security settings

' GetSecurityDescriptor Return values

Dim objReturnCodes : Set objReturnCodes = CreateObject("Scripting.Dictionary")
Const SUCCESS = 0
Const ACCESS_DENIED = 2
Const UNKNOWN_FAILURE = 8
Const PRIVILEGE_MISSING = 9
Const INVALID_PARAMETER = 21

' Security Descriptor Control Flags

Dim objControlFlags : Set objControlFlags = CreateObject("Scripting.Dictionary")
objControlFlags.Add 32768, "SelfRelative"
objControlFlags.Add 16384, "RMControlValid"
objControlFlags.Add 8192, "SystemAclProtected"
objControlFlags.Add 4096, "DiscretionaryAclProtected"
objControlFlags.Add 2048, "SystemAclAutoInherited"
objControlFlags.Add 1024, "DiscretionaryAclAutoInherited"
objControlFlags.Add 512, "SystemAclAutoInheritRequired"
objControlFlags.Add 256, "DiscretionaryAclAutoInheritRequired"
objControlFlags.Add 32, "SystemAclDefaulted"
objControlFlags.Add 16, "SystemAclPresent"
objControlFlags.Add 8, "DiscretionaryAclDefaulted"
objControlFlags.Add 4, "DiscretionaryAclPresent"
objControlFlags.Add 2, "GroupDefaulted"
objControlFlags.Add 1, "OwnerDefaulted"

' ACE Access Right

Dim objAccessRights : Set objAccessRights = CreateObject("Scripting.Dictionary")
objAccessRights.Add 2032127, "FullControl"
objAccessRights.Add 1048576, "Synchronize"
objAccessRights.Add 524288, "TakeOwnership"
objAccessRights.Add 262144, "ChangePermissions"
objAccessRights.Add 197055, "Modify"
objAccessRights.Add 131241, "ReadAndExecute"
objAccessRights.Add 131209, "Read"
objAccessRights.Add 131072, "ReadPermissions"
objAccessRights.Add 65536, "Delete"
objAccessRights.Add 278, "Write"
objAccessRights.Add 256, "WriteAttributes"
objAccessRights.Add 128, "ReadAttributes"
objAccessRights.Add 64, "DeleteSubdirectoriesAndFiles"
objAccessRights.Add 32, "ExecuteFile"
objAccessRights.Add 16, "WriteExtendedAttributes"
objAccessRights.Add 8, "ReadExtendedAttributes"
objAccessRights.Add 4, "AppendData"
objAccessRights.Add 2, "CreateFiles"
objAccessRights.Add 1, "ReadData"

' ACE Types

Dim objAceTypes : Set objAceTypes = CreateObject("Scripting.Dictionary")
objAceTypes.Add 0, "Allow"
objAceTypes.Add 1, "Deny"
objAceTypes.Add 2, "Audit"

' ACE Flags

Dim objAceFlags : Set objAceFlags = CreateObject("Scripting.Dictionary")
objAceFlags.Add 128, "FailedAccess"
objAceFlags.Add 64, "SuccessfulAccess"
objAceFlags.Add 16, "Inherited"
objAceFlags.Add 8, "InheritOnly"
objAceFlags.Add 4, "NoPropagateInherit"
objAceFlags.Add 2, "ContainerInherit"
objAceFlags.Add 1, "ObjectInherit"

Sub ReadNTFSSecurity(objWMI, strPath)
  objFileOut.Write("  Displaying NTFS Security" & vbCrLf)

  Dim objSecuritySettings : Set objSecuritySettings = _
    objWMI.Get("Win32_LogicalFileSecuritySetting='" & strPath & "'")
  Dim objSD : objSecuritySettings.GetSecurityDescriptor objSD

  Dim strDomain : strDomain = objSD.Owner.Domain
  If strDomain <> "" Then strDomain = strDomain & "\"
  objFileOut.Write("  Owner: " & strDomain & objSD.Owner.Name & vbCrLf)
  objFileOut.Write("  Owner SID: " & objSD.Owner.SIDString & vbCrLf)

  objFileOut.Write("  Basic Control Flags Value: " & objSD.ControlFlags & vbCrLf)
  objFileOut.Write("  Control Flags:" & vbCrLf)

  DisplayValues objSD.ControlFlags, objControlFlags

  objFileOut.Write(vbCrLf)

  Dim objACE

  ' Display the DACL
  objFileOut.Write("  Discretionary Access Control List:" & vbCrLf)
  For Each objACE in objSD.DACL
    DisplayACE objACE
  Next

  ' Display the SACL (if there is one)
  If Not IsNull(objSD.SACL) Then
    objFileOut.Write("  System Access Control List:" & vbCrLf)
    For Each objACE in objSD.SACL
      DisplayACE objACE
    Next
  End If
End Sub

Sub ReadShareSecurity(objWMI, strName)

	Dim objSecuritySettings : Set objSecuritySettings = objWMI.Get("Win32_LogicalShareSecuritySetting='" & strName & "'")
	Dim objSD : objSecuritySettings.GetSecurityDescriptor objSD
	Dim objACE

	For Each objACE in objSD.DACL
		DisplayACE objACE
	Next

End Sub

Sub DisplayValues(dblValues, objSecurityEnumeration)

	Dim dblValue

	For Each dblValue in objSecurityEnumeration
		If dblValues >= dblValue Then
			If (objSecurityEnumeration(dblValue) <> "Synchronize") Then objFileOut.Write("Right:" & objSecurityEnumeration(dblValue) & vbCrLf)
			dblValues = dblValues - dblValue
	    End If
	Next

End Sub

Sub DisplayACE(objACE)

	Dim strDomain : strDomain = objAce.Trustee.Domain

	If strDomain <> "" Then strDomain = strDomain & "\"

	If (UCase(objAceTypes(objACE.AceType)) = "ALLOW") Then
		objFileOut.Write("Trustee:" & UCase(strDomain & objAce.Trustee.Name) & vbCrLf)
		DisplayValues objACE.AccessMask, objAccessRights
	End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(SysDrive & "\backedup_shares\")
Set objFolder = Nothing
Set colDrives = objFSO.Drives
Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = ".*\\"
Set objREx = CreateObject("VBScript.RegExp")
objREx.Global = True   
objREx.IgnoreCase = True
objREx.Pattern = "[\:\\ ]"
Set objFileOut = objFSO.OpenTextFile(SysDrive & "\backedup_shares\shares.txt", ForWriting, True)
Set objFileOut1 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\drives_sn.list", ForWriting, True)
Set objFileOut2 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\shares_backup.cmd", ForWriting, True)
Set objFileOut3 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\shares_restore.cmd", ForWriting, True)
Set objFileOut5 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\remove_inheritance.cmd", ForWriting, True)
Set objFileOut6 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\own_shares.cmd", ForWriting, True)
objFileOut2.Write("@echo off" & vbCrLf)
objFileOut3.Write("@echo off" & vbCrLf)
objFileOut5.Write("@echo off" & vbCrLf)
objFileOut6.Write("@echo off" & vbCrLf)

'Создаём список сопоставления букв разделов и серийных номеров
For Each objDrive in colDrives
	If objDrive.DriveType = 2 Then
		objFileOut1.Write(objDrive.DriveLetter & ": " & objDrive.SerialNumber & vbCrLf)
	End If
Next

Dim strComputer : strComputer = "."
Dim objWMI : Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Dim colItems : Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Share WHERE Type='0'", "WQL", WBEM_RETURN_IMMEDIATELY + WBEM_FORWARD_ONLY)
Dim objItem

dircounter = 0

copycommand = "copy /A /Y "

'Создаём папку для хранения флагов существования путей шар (для защиты от каталогов, имеющих более 1 шары)
If (Not objFSO.FolderExists(SysDrive & "\backedup_shares\flags")) Then
	oShell.run "cmd /c ""mkdir " & SysDrive & "\backedup_shares\flags""",0,bWaitOnReturn
End If

For Each objItem in colItems
	exclude = 0
	For Each Argument In WScript.Arguments
		If (InStr(objItem.Name, Argument) = 1) Then
			exclude = 1
		End If
	Next
	'Ставить флаг установки IIS, если в названии хотя бы одной шары присутствует "SCCM" или "SMS"
	If (((InStr(objItem.Name, "SCCM") = 1) Or (InStr(objItem.Name, "SMS") = 1)) And (Not objFSO.FileExists(SysDrive & "\backedup_shares\iis"))) Then
		Set objFileOut4 = objFSO.OpenTextFile(SysDrive & "\backedup_shares\iis", ForWriting, True)
		objFileOut4.Write("iis" & vbCrLf)
		objFileOut4.Close
	End If
	If (exclude = 0) Then
		If ((InStr(objItem.Name, "SCCM") = 0) And (InStr(objItem.Name, "SMS") <> 1)) Then
			objFileOut.Write("Share:" & objItem.Name & vbCrLf)
			objFileOut.Write("Path:" & objItem.Path & vbCrLf)
			objFileOut.Write("Desc:" & objItem.Caption & vbCrLf)
			ReadShareSecurity objWMI, objItem.Name
			'ReadNTFSSecurity objWMI, objItem.Path
			objFileOut.Write(vbCrLf)
		End If
		If (UCase(left(objItem.Path, 2)) = SysDrive) And (Not objFSO.FileExists(SysDrive & "\backedup_shares\flags\" & objREx.Replace(objItem.Path,"_"))) Then
			'Генерируем shares_backup.cmd
			objFileOut2.Write("move /y """ & "%1" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """ """ & "%1\backedup_shares\" & objItem.Name & """" & vbCrLf)
			objFileOut2.Write("if not %errorlevel%==0 exit %errorlevel%" & vbCrLf)
			
			'Генерируем shares_restore.cmd
			containpath = Left(objItem.Path, Len(objItem.Path) - Len(objRegEx.Replace(objItem.Path,"")) - 1)
			objFileOut3.Write("mkdir ""%SystemDrive%" & Right(containpath, Len(containpath) - Len("C:")) & """" & vbCrLf)
			objFileOut3.Write("move /y """ & "%SystemDrive%\backedup_shares\" & objItem.Name & """ """ & "%SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """" & vbCrLf)
			
			'Генерируем own_shares.cmd
			objFileOut6.Write("echo Taking ownership on %SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & " ..." & vbCrLf)
			objFileOut6.Write("takeown /R /A /D ""Y"" /F ""%SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """ > NUL" & vbCrLf)
			objFileOut6.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			objFileOut6.Write("echo Adding Administrators and SYSTEM to ACL on %SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & " ..." & vbCrLf)
			objFileOut6.Write("%SystemDrive%\backedup_shares\setacl.exe -silent -ot file -on ""%SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """ -actn ace -ace ""n:S-1-5-18;p:full"" -ace ""n:S-1-5-32-544;p:full""" & vbCrLf)
			objFileOut6.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			
			' Отключаем наследование и удаляем CREATOR-OWNER (remove_inheritance.cmd)
			objFileOut5.Write("echo Disabling ACL inheritance on %SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & " ..." & vbCrLf)
			objFileOut5.Write("%SystemDrive%\backedup_shares\setacl.exe -silent -ot file -on ""%SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """ -actn setprot -op ""dacl:p_c;sacl:p_c""" & vbCrLf)
			objFileOut5.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			objFileOut5.Write("echo Removing CREATOR-OWNER from ACL on %SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & " ..." & vbCrLf)
			objFileOut5.Write("%SystemDrive%\backedup_shares\setacl.exe -silent -ot file -on ""%SystemDrive%" & Right(objItem.Path, Len(objItem.Path) - Len("C:")) & """ -actn trustee -trst ""n1:S-1-3-0;ta:remtrst;w:dacl,sacl""" & vbCrLf)
			objFileOut5.Write("if not %errorlevel%==0 (" & vbCrLf & "	echo Failed." & vbCrLf & ") else (" & vbCrLf & "	echo Done." & vbCrLf & ")" & vbCrLf)
			
			'Бэкапим ACL NTFS для каталогов, которые будем перемещать
			oShell.Run SysDrive & "\backedup_shares\setacl.exe -on """ & objItem.Path & """ -ot file -actn list -lst ""f:sddl;w:d,s,o,g"" -bckp """ & SysDrive & "\backedup_shares\" & dircounter & ".acl""",0,bWaitOnReturn
			If (dircounter = 0) Then
				copycommand = copycommand & SysDrive & "\backedup_shares\" & dircounter & ".acl"
			Else
				copycommand = copycommand & "+" & SysDrive & "\backedup_shares\" & dircounter & ".acl"
			End If
			dircounter = dircounter + 1
			
			'Создаём флаг, указывающий, что шара с таким каталогом уже есть в списке
			Set objFlagObj = objFSO.OpenTextFile(SysDrive & "\backedup_shares\flags\" & objREx.Replace(objItem.Path,"_"), ForWriting, True)
			objFlagObj.Close
		End If
	End If
Next

'Удаляем папку для хранения флагов существования путей шар
If (objFSO.FolderExists(SysDrive & "\backedup_shares\flags")) Then
	oShell.run "cmd /c ""rd /s /q " & SysDrive & "\backedup_shares\flags""",0,bWaitOnReturn
End If

'Склеиваем файлы со списком ACL для каждого из каталогов
copycommand = copycommand & " " & SysDrive & "\backedup_shares\acllist.lca"
oShell.run "cmd /c """ & copycommand & """",0,bWaitOnReturn
oShell.run "cmd /c ""del /F /Q " & SysDrive & "\backedup_shares\*.acl""",0,bWaitOnReturn

'Меняем кодировку файла со списокм ACL с UCS-2 LE (UTF-16) на UTF-8
Set ADODBStream = CreateObject("ADODB.Stream")
ADODBStream.Type = 2
ADODBStream.Charset = "UTF-16LE"
ADODBStream.Open()
ADODBStream.LoadFromFile(SysDrive & "\backedup_shares\acllist.lca")
Text = ADODBStream.ReadText()
ADODBStream.Close()
ADODBStream.Charset = "UTF-8"
ADODBStream.Open()
ADODBStream.WriteText(Text)
ADODBStream.SaveToFile SysDrive & "\backedup_shares\acllist.lca", 2
ADODBStream.Close()

objFileOut3.Write("cscript.exe " & "%SystemDrive%\backedup_shares\shares_restore.vbs %1" & vbCrLf)
objFileOut3.Write("%SystemDrive%\backedup_shares\setacl.exe -ignoreerr -on ""%SystemDrive%"" -ot file -actn restore -bckp ""%SystemDrive%\backedup_shares\acllist.lca""" & vbCrLf)
objFileOut3.Write("%SystemDrive%\backedup_shares\remove_inheritance.cmd" & vbCrLf)
objFileOut.Close
objFileOut1.Close
objFileOut2.Close
objFileOut3.Close
objFileOut5.Close
objFileOut6.Close
