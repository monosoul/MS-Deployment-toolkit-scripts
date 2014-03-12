Const ForReading = 1
Const ForWriting = 2

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

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder("c:\old_c\")
Set objFolder = Nothing
Set objFileOut = objFSO.OpenTextFile("c:\old_c\shares.txt", ForWriting, True)
Set objFileOut2 = objFSO.OpenTextFile("c:\old_c\shares_backup.cmd", ForWriting, True)
Set objFileOut3 = objFSO.OpenTextFile("c:\old_c\shares_restore.cmd", ForWriting, True)

Dim strComputer : strComputer = "."
Dim objWMI : Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Dim colItems : Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Share WHERE Type='0'", "WQL", WBEM_RETURN_IMMEDIATELY + WBEM_FORWARD_ONLY)
Dim objItem

For Each objItem in colItems
	objFileOut.Write("Share:" & objItem.Name & vbCrLf)
	objFileOut.Write("Path:" & objItem.Path & vbCrLf)
	objFileOut.Write("Desc:" & objItem.Caption & vbCrLf)
	objFileOut2.Write("move /y """ & objItem.Path & """ ""c:\old_c\" & objItem.Name & """" & vbCrLf)
	objFileOut3.Write("move /y ""c:\old_c\" & objItem.Name & """ """ & objItem.Path & """" & vbCrLf)
	ReadShareSecurity objWMI, objItem.Name
'	ReadNTFSSecurity objWMI, objItem.Path
	objFileOut.Write(vbCrLf)
Next

objFileOut2.Write("REM !!! cscript.exe c:\old_c\clear_c.vbs" & vbCrLf)
objFileOut3.Write("cscript.exe c:\old_c\shares_restore.vbs" & vbCrLf)
objFileOut.Close
objFileOut2.Close
