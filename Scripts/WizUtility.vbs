' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      WizUtility.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Common Utility functions used by wizard UI Scripts
' // 
' // ***************************************************************************


Function BrowseForFolder(sDefaultValue)
	Dim iRetVal


	iRetVal = Success

	' Workaround for BrowseForFolder problem when called from an HTA:  sometimes it doesn't show up.

	oEnvironment.Item("DefaultFolderPath") = sDefaultValue
	iRetVal = oShell.Run("wscript.exe """ & oUtility.ScriptDir & "\LTIGetFolder.wsf""", 1, true)
	If iRetVal = 0 then
		BrowseForFolder = oEnvironment.Item("FolderPath")
	Else
		BrowseForFolder = sDefaultValue
	End if


End Function



'
' Call a VBScript and handle the errors
'

Function DisplayErrorIfAny ( sStatement )
	Dim sError

	If Err.Number = 0 then
		Exit function
	End if

	'sError = sError & "Use Ctrl-C to copy the text of this message to the clipboard!" &  vbNewLine
	sError = sError & "A VBScript Runtime Error has occurred: " & vbNewLine & vbNewLine
	sError = sError & "Error: " & Err.Number & " = " & Err.Description & vbNewLine & vbNewLine
	sError = sError & "VBScript Code:" & vbNewLine & "-------------------" &  vbNewLine
	sError = sError & left(sStatement,800)

	oLogging.CreateEntry  sError ,LogTypeError

	If oLogging.Debug then
		sError = sError & vbNewLine & "-------------------" &  vbNewLine
		sError = sError & "Do you wish to attempt debugging on this script?"
		DisplayErrorIfAny = MsgBox ( sError, vbYesNo , "VBScript Runtime Error" ) = vbYes
	Else
		Alert sError
		DisplayErrorIfAny = FALSE
	End if

End function


Function ExecuteWithErrorHandling ( statements )
	Dim RunAgain

	RunAgain = FALSE

	On error resume Next
	Err.Clear
	ExecuteGlobal statements
	ExecuteWithErrorHandling = err.number = 0

	RunAgain = DisplayErrorIfAny (statements)
	On error goto 0

	If RunAgain then
		ExecuteGlobal statements
	End if

End function


Function EvalWithErrorHandling ( fn )

	RunAgain = FALSE

	On error resume Next
	Err.Clear
	EvalWithErrorHandling = eval(fn)
	EvalWIthErrorHandling = EvalWIthErrorHandling and (err.number = 0)

	RunAgain = DisplayErrorIfAny (fn)
	On error goto 0

	If RunAgain then
		EvalWithErrorHandling = eval(fn)
	End if

End function


'
' Create an XML Document Node with assoticated Elements and Attributes
'
'  Parameters:
'      oXMLDoc - XML DOM object ( Created from MSXML2.DOMDocument )
'  oTargetNode - XML Document Element where the element is created
'    sNodeName - Name of the Element Created
'  oAttributes - VBScript Dictionary object containing a list of XML Attributes to add
'    oElements - VBScript Dictionary object containing a list of XML Elements to add
'
'  Example Usage:
'   dim     xmlDoc, oAttributes, oElements
'   set  oAttributes = CreateObject("Scripting.Dictionary")
'   set    oElements = CreateObject("Scripting.Dictionary")
'   Set       xmlDoc = CreateXMLDOMObject
'
'   xmlDoc.Load "c:\SomeFile.xml"     or   xmlDoc.appendChild xmlDoc.createElement("MyRootElement")
'
'   oAttributes.Add "guid", left( CreateObject("Scriptlet.TypeLib").GUID, 38 ) ' Strip trailing NULL's
'     oElements.Add "Name","BuildName.value"
'     oElements.Add "Version","1.0"
'
'   CreateXMLNode xmlDoc, xmlDoc.documentElement, "MyElement", oAttributes, oElements
'
'   xmlDoc.Save "c:\SomeFile.xml"
'

Function CreateXMLNode ( xmlDoc, oTargetNode, sNodeName, oAttributes , oElements)

	Dim Key, oElement, oNodeRoot

	Set oNodeRoot = xmlDoc.createElement(sNodeName)

	For each key in oAttributes
		If IsEmpty(key) then
			Exit For
		End if
		oNodeRoot.setAttribute key, oAttributes.Item(key)
	Next

	For each key in oElements.Keys
		If isempty(key) then exit for
		Set oElement = xmlDoc.createElement(key)
		oElement.text = oElements.Item(key)
		oNodeRoot.appendChild oElement
	Next

	oTargetNode.appendChild oNodeRoot

End function


'
' hide/unhide an DHTML element
'
Sub HideUnhide ( oHTMLElement, isVisible )

	If isVisible then
		oHTMLElement.style.display = "inline"
	Else
		oHTMLElement.style.display = "none"
	End if

End sub


'
' Copy Files to a folder
'
'  More information:
'      http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/objects/folder/copyhere.asp
'
' WARNING: There is no way to tell if the user hits the Cancel Button during copy.
'
' Parameters:
'       Source - Item or items to copy. Can be a file name string, a FolderItem object, or a FolderItems object.
'     sDestDir - Target Destination Directory. If missing then this procedure will create the directory.
'        Flags - Copy Flags for CopyHere (Typical settings: 2048 or 16 ===> 2064 ):
'   CleanFirst - Clean the directory first.
'
Sub CopyFileWithProgressEx ( Source , sDestDir, Flags, CleanFirst )

	Dim objFolder, hr

	If CleanFirst and oFso.FolderExists( sDestDir ) then
		oFSO.DeleteFolder sDestDir
	End if

	' Open the Destination Directory (Create if missing)
	oUtility.VerifyPathExists sDestDir
	Set objFolder = CreateObject("Shell.Application").NameSpace(sDestDir)

	If objFolder is nothing then
		Err.Raise 507,,"Destination Folder not set: " & sDestDir
		Exit sub
	End if

	objFolder.CopyHere Source, Flags
	Set objFolder = nothing

End sub


Sub CopyFileWithProgress ( Source , sDestDir )
	' Common Settings
	'    16 Respond with "Yes to All" for any dialog box that is displayed.
	'  2048 Version 4.71. Do not copy the security attributes of the file.
	CopyFileWithProgressEx Source , sDestDir, 2064, TRUE
End sub

'
' Tests a filename for invalid characters.
'
Function IsValidFileName (FileName)

	Dim regEx

	Set regEx = New RegExp
	regEx.Pattern = "[\x00-\x1F\<\>\:\""\/\\\|\%\*\?\']"   'Strict Subset
	IsValidFileName = (not regEx.Test ( FileName )) and (trim(FileName) <> "") and len(trim(FileName)) <= 253

	Select Case UCase(Trim(FileName))
	Case "CON", "AUX", "COM1", "COM2", "COM3", "COM4", "LPT1", "LPT2", "LPT3", "PRN", "NUL"
		IsValidFileName = FALSE
	End select

End function


Function IsValidPath (FilePath)

	Dim regEx

	Set regEx = New RegExp
	regEx.Pattern = "[\x00-\x1F\<\>\""\%\*\?\']"   'Strict Subset
	IsValidPath = (not regEx.Test ( FilePath )) and (trim(FilePath) <> "") and len(trim(FilePath)) <= 253

End function


Function IsValidComputerName ( OSDComputerName )

	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "[^a-zA-Z0-9\-\_]"   'Strict Subset
	' regEx.Pattern = "[^a-zA-Z0-9\!\@\#\$\%\^\&\'\)\(\.\-\_\{\}\~ ]" ' SetComputerName compliant
	IsValidComputerName 	= not regEx.Test ( OSDComputerName ) and len(OSDComputerName) <= 15

End function



Sub AddToList(byref List, Item, Delimiter)
	oStrings.AddToList List,Item,Delimiter
End sub

Function ForceAsString ( InputVar )
	ForceAsString = oStrings.ForceAsString ( InputVar )
End function

Function IsWhiteSpace (MyChar)
	IsWhiteSpace = oStrings.IsWhiteSpace(MyChar)
End function


Function TrimAllWS( MyString )
	TrimAllWS = oStrings.TrimAllWS( MyString )
End function


'
' Validate Credentials against either a Domain or a UNC Disk Resource
'   Function will return TRUE if successfull, otherwise it will return an ERROR string!
'
Function CheckCredentials(sValidateUNC, sUserName, sDomain, sPassword)

	Dim sFullUserName
	Dim sJustUserName
	Dim sUserObjectPath
	Dim oWinNT
	Dim oDomain ' As IADsDomain
	Dim sDriveLetter


	sFullUserName = sUserName
	If sDomain <> "" then
		sFullUserName = sDomain & "\" & sFullUserName
	End if


	If sValidateUNC = "" then

		If oEnvironment.Item("ValidateAgainstUNCPath") <> "" then
			sValidateUNC = oEnvironment.Item("ValidateAgainstUNCPath")
		Else

			sValidateUNC = oEnvironment.Item("DeployRoot")
			For Each oDomain in objWMI.InstancesOf("Win32_ComputerSystem")
				if oDomain.DomainRole = 1 or oDomain.DomainRole = 3 or oDomain.DomainRole = 4 or oDomain.DomainRole = 5 then
					oLogging.CreateEntry "Computer is part of a domain, valiadate against domain.", LogTypeInfo
					sValidateUNC = "" 
				End if
			next

		End if

	End if


	If sValidateUNC <> "" then

		'
		' Validate the credentials against an actual UNC disk resource
		'

		sDriveLetter = oUtility.MapNetworkDriveEx (sValidateUNC, sFullUserName, sPassword, LogTypeInfo )

		If len(sDriveLetter) = 2 then
			If bLeaveShareOpen <> TRUE and ucase(sValidateUNC) <> ucase(oEnvironment.Item("DeployRoot")) then
				oNetwork.RemoveNetworkDrive sDriveLetter
			End if
			CheckCredentials = TRUE
		Else
			CheckCredentials = sDriveLetter ' If not a drive letter, then this is an error string
		End if

	Else
	
		CheckCredentials = CheckCredentialsAD( "", sUserName, sDomain, sPassword)

	End if

	If CheckCredentials <> TRUE then
		oLogging.CreateEntry "Credentials Script: " & CheckCredentials, LogTypeInfo
		Exit function
	End if

	oEnvironment.Item("UserCredentials") = Cstr(TRUE)

	If bDoNotSaveParameters = TRUE then
		window.close
		CheckCredentials = FALSE
	End if

End Function

'
' Validate Credentials against a Domain
'   Function will return TRUE if successfull, otherwise it will return an ERROR string!
'
Function CheckCredentialsAD(sJoinDomain, sUserName, sDomain, sPassword)

	Dim sFullUserName
	Dim sJustUserName
	Dim sUserObjectPath
	Dim oWinNT
	Dim oDomain ' As IADsDomain
	Dim sDriveLetter
	Dim sJoinDomainNew
	
	sJoinDomainNew = sJoinDomain


	sFullUserName = sUserName
	If sDomain <> "" then
		sFullUserName = sDomain & "\" & sFullUserName
	End if


	'
	' Validate the credentials against a domain or computer server using Active Directory authentication.
	'

	' The credentials can be in the form "BillG", "redmond\BillG", or "BillG@Microsoft.Com". Cleanup for use.

	sJustUserName = sUserName
	If Instr(1,sJustUserName,"@") <> 0 then
		' Username is in form: BillG@redmond.corp.Microsoft.com, remove domain.
		sJustUserName = left(sJustUserName, instr(1,sJustUserName,"@") - 1)
	ElseIf instr(1,sJustUserName,"\") <> 0 then
		' Username is in form: redmond\BillG, remove domain.
		sJustUserName = mid(sJustUserName, instr(1,sJustUserName,"\") + 1 )
	End if

	On Error Resume Next
	Set oWinNT = GetObject("WinNT:")
	If Err then
		If oEnvironment.Item("OSVersion") = "WinPE" then
			oLogging.CreateEntry "Unable to verify domain credentials in Windows PE since ADSI is not available", LogTypeInfo
			CheckCredentialsAD = TRUE
			Err.Clear
			Exit Function
		Else
			CheckCredentialsAD = Err.Description & " (" & Hex(Err.Number) & ")"
			Exit Function
		End if
	End if



	If sJoinDomainNew = "" then
		sJoinDomainNew = sDomain
	ElseIf Instr(1, sUserName, "\") <> 0 then
		sJoinDomainNew = left(sUserName, instr(1, sUserName, "\") - 1)
	ElseIf Instr(1, sUserName, "@") <> 0 then
		sJoinDomainNew = mid(sJustUserName, instr(1,sJustUserName,"@") + 1 )
	End if


	' 1 = ADS_SECURE_AUTHENTICATION
	Set oDomain = oWinNT.OpenDSObject("WinNT://" & sJoinDomainNew & "/" & sJustUserName &  ",user" ,  sFullUserName, sPassword, 1 )
		


	If Err.Number = &h80070035 then
		CheckCredentialsAD = "Network path not found (80070035)"
	ElseIf Err.Number = &H8007054B then
		CheckCredentialsAD = "Domain could not be contacted (8007054B)"
	ElseIf Err.Number = &h8007052E then
		CheckCredentialsAD = "User ID or password is invalid (8007052E)"
	ElseIf Err.Number = &h800708AD then
		CheckCredentialsAD = "User ID is not valid (800708AD)"
	ElseIf Err then
		CheckCredentialsAD = Err.Description & " (" & Hex(Err.Number) & ")"
	ElseIf oDomain is nothing then
		CheckCredentialsAD = "Domain validation failed - " & Err.Description & " (" & Err.Number & ")"
	Else
		Err.Clear
		CheckCredentialsAD = TRUE
	End if
	On error goto 0

	If CheckCredentialsAD <> TRUE then
		oLogging.CreateEntry "Credentials Script: " & CheckCredentialsAD, LogTypeInfo
		Exit function
	End if

	oEnvironment.Item("UserCredentials") = Cstr(TRUE)

	If bDoNotSaveParameters = TRUE then
		window.close
		CheckCredentialsAD = FALSE
	End if

End Function

Function GetDomainDefault
	Dim oComputer


	' Get Local Domain

	GetDomainDefault = ""
	if oEnvironment.Item("UserDomain") <> "" then
		GetDomainDefault = oEnvironment.Item("USERDOMAIN")
	else
		On Error Resume Next
		For each oComputer in objWMI.InstancesOf("Win32_ComputerSystem")
			If oComputer.DomainRole <> 0 then
				GetDomainDefault = oComputer.Domain
			End if
		Next
		On Error Goto 0
	end if

End Function


Function GetDestDisk

	' Preference search order is: oProperties, oEnvironment/CS.INI , and TS.XML ( DefaultDestinationXxx )
	GetDestDisk = Property("DestinationDisk")
	If GetDestDisk = "" then
		GetDestDisk = Property("DefaultDestinationDisk")
	End if
	If GetDestDisk = "" then
		GetDestDisk = "0"
	End if

End function

Function GetDestPart

	' Preference search order is: oProperties, oEnvironment/CS.INI , and TS.XML ( DefaultDestinationXxx )
	GetDestPart = Property("DestinationPartition")
	If GetDestPart = "" then
		GetDestPart = Property("DefaultDestinationPartition")
	End if
	If GetDestPart = "" then
		GetDestPart = "1"
	End if

End function

Function HasGoodDestDisk( sDestDisk )

	Dim oDisk, oDisks

	Set oDisks = objWMI.ExecQuery("Select index from Win32_DiskDrive where MediaType like 'Fixed%hard disk%'")

	HasGoodDestDisk = False
	For Each oDisk in oDisks
		If Cstr(oDisk.Index) = sDestDisk Then
			oLogging.CreateEntry "Validated Disk exists", LogTypeInfo
			HasGoodDestDisk = True
			exit for
		End If
	Next

End function

Function HasGoodDestPart ( sDestDisk, sDestPart )

	Dim sDestDrive
	
	HasGoodDestPart = true
	If Property("DeploymentType") = "REFRESH" Then

		sDestDrive = oUtility.DetermineDriveFromDiskPart( sDestDisk, sDestPart )
		HasGoodDestPart = sDestDrive = oEnv("SystemDrive")
		oLogging.CreateEntry "oUtility.DetermineDriveFromDiskPart( " & sDestDisk& ", " & sDestPart& " )  = " & sDestDrive, LogTypeInfo

	End if

End function


Function RmPropIfFound( Prop )

	If oProperties.Exists(Prop) then
		oProperties.Remove(Prop)
	End if

End function


