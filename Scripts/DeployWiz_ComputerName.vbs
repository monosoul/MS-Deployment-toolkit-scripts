' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_Initialization.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Main Client Deployment Wizard Initialization routines
' // 
' // ***************************************************************************


Option Explicit


Function InitializeComputerName

	If oProperties("OSDComputerName") = "" then
		OSDComputerName.Value = oUtility.ComputerName 
	End if

	If UCase(oEnvironment.Item("SkipComputerName")) = "YES" then
		OSDComputerName.disabled = true
	End if

End Function


Function ValidateComputerName

	' Check Warnings
	ParseAllWarningLabels


	If Len(OSDComputerName.value) > 15 then
		InvalidChar.style.display = "none"
		TooLong.style.display = "inline"
		ValidateComputerName = false
		ButtonNext.disabled = true
	ElseIf IsValidComputerName ( OSDComputerName.Value ) then
		ValidateComputerName = TRUE
		InvalidChar.style.display = "none"
		TooLong.style.display = "none"
	Else
		InvalidChar.style.display = "inline"
		TooLong.style.display = "none"
		ValidateComputerName = false
		ButtonNext.disabled = true
	End if

End function




''''''''''''''''''''''''''''

Function AddItemToMachineObjectOUOpt(item)
	Dim oOption

	set oOption = document.createElement("OPTION")
	oOption.Value = item
	oOption.Text = item
	oOption.Title = item
	MachineObjectOUOptional.Add oOption
	MachineObjectOUOptionalBtn.style.display = "inline"

End function


Function InitializeDomainMembership
	Dim oLDAP, oOptOU, oItem
	Dim sFoundFile
	Dim iRetVal


	' Prepopulate the join account details if they are presently blank

	If Property("DomainAdmin") = "" and Property("DomainAdminDomain") = "" and Property("DomainAdminPassword") = "" then
		If Property("UserID") <> "" and Property("UserDomain") <> "" and Property("UserPassword") <> "" Then

			DomainAdmin.value  = Property("UserID")
			DomainAdminDomain.value = Property("UserDomain")
			DomainAdminPassword.value  = Property("UserPassword")

		End if
	End if


	If JoinWorkgroup.value <> "" then
		JDRadio2.checked = TRUE
	ElseIf JoinDomain.Value = "" then
		On Error Resume Next
		JoinDomain.value = CreateObject("ADSystemInfo").DomainDNSName
		On Error Goto 0

		If JoinDomain.Value = "" then
			JoinWorkgroup.value = "WORKGROUP"
			JDRadio2.checked = TRUE
		Else
			JDRadio1.checked = TRUE
			' Domain.value = JoinDomain.Value

			On error resume next
			' Will extract out the existing OU (if any) for the current machine.
			set oLDAP = GetObject("LDAP://" & CreateObject("ADSystemInfo").ComputerName)
			MachineObjectOU.Value = oLDAP.Get("Organizational-Unit-Name")
			On error goto 0

		End if
	End if


	''''''''''''''''''''''''''''''''
	'
	' Populate OU method #1 - Query ADSI
	'

	MachineObjectOUOptionalBtn.style.display =  "none"


	''''''''''''''''''''''''''''''''
	'
	' Populate OU method #2 - Read MachineObjectOUOptional[1...n] property
	'

	If MachineObjectOUOptionalBtn.style.display <> "inline" then
		oOptOU = Property("DomainOUs")
		If isarray(oOptOU) then

			For each oItem in oOptOU
				AddItemToMachineObjectOUOpt oItem
			Next
			MachineObjectOUOptionalBtn.style.display = "inline"

		ElseIf oOptOU <> "" then
			AddItemToMachineObjectOUOpt oOptOU
		End if
	End if


	''''''''''''''''''''''''''''''''
	'
	' Populate OU method #3 - Read ...\control\DomainOUList.xml
	'
	' Example:
	'	<?xml version="1.0" encoding="utf-8"?>
	'	<DomainOUs>
	'		<DomainOU>OU=Test1</DomainOU>
	'		<DomainOU>OU=Test2</DomainOU>
	'	</DomainOUs>
	'

	If MachineObjectOUOptionalBtn.style.display <> "inline" then
	
		iRetVal = oUtility.FindFile( "DomainOUList.xml" , sFoundFile)
		if iRetVal = SUCCESS then
			For each oItem in oUtility.CreateXMLDOMObjectEx( sFoundFile ).selectNodes("//DomainOUs/DomainOU")
				AddItemToMachineObjectOUOpt oItem.text
			Next
		End if
	End if

	If MachineObjectOUOptionalBtn.style.display = "inline" then

		document.body.onMouseDown = getRef("DomainMouseDown")
		document.body.onKeyDown   = getRef("MachineObjectOUOptionalKeyPress")

	End if

	ValidateDomainMembership

End Function


Function MachineObjectOUOptionalKeyPress

	dim OUOpt

	on error resume next
	set OUOpt = MachineObjectOUOptional
	on error goto 0

	If isempty(OUOpt) then
		KeyHandler
	ElseIf window.event.srcElement is MachineObjectOUOptional then
		If window.event.keycode = 13 then
		' Enter
			MachineObjectOU.value = MachineObjectOUOptional.value
			PopupBox.style.display = "none"

		ElseIf window.event.keycode = 27 then
			' escape
			PopupBox.style.display = "none"
		End if
	Else
		KeyHandler
	End if

End function


Function DomainMouseDown
	If not window.event.srcElement is MachineObjectOUOptional and not window.event.srcElement is MachineObjectOUOptionalBtn then
		PopupBox.style.display = "none"
	End if
End function


Function HideUnHideComboBox

	If UCase(PopupBox.style.display) <> "NONE" then

		HideUnhide PopupBox, FALSE

		document.body.onMouseDown = ""
		document.body.onKeyDown   = getRef("KeyHandler")

	Else

		HideUnhide PopupBox, TRUE
		MachineObjectOUOptional.focus

		document.body.onMouseDown = getRef("DomainMouseDown")
		document.body.onKeyDown   = getRef("MachineObjectOUOptionalKeyPress")

	End if

End function


Function ValidateDomainMembership_Final

	If JDRadio1.checked then

		RMPropIfFound("JoinWorkgroup")
		JoinWOrkgroup.Value = ""

	Else

		RMPropIfFound("JoinDomain")
		JoinDomain.Value = ""

	End if

	ValidateDomainMembership_Final = true


End function

'''''''''''''''''''''''''''''''''''''
'  Validate Domain Membership
'

Function ValidateDomainMembership
	Dim IsDomain
	Dim r

	MissingCredentials.style.display = "none"
	InvalidCredentials.style.display = "none"
	InvalidOU.style.display = "none"

	isDomain = JDRadio1.checked

	If not isDomain then

		RMPropIfFound("BdeInstall")
		RMPropIfFound("BdeInstallSuppress")
		RMPropIfFound("DoCapture")
		RMPropIfFound("BackupFile")
		
		If Property("DeploymentType") <> "REFRESH" and Property("DeploymentType") <> "REPLACE" then
			RMPropIfFound("ComputerBackupLocation")
		End if

	End if

	If UCase(oEnvironment.Item("SkipDomainMembership")) = "YES" then

		' Hide all the domain/workgroup settings
		DomainSection.style.display = "none"

		JDRadio1.disabled = true
		JDRadio2.disabled = true

		JoinDomain.disabled = true
		DomainAdmin.disabled = true
		DomainAdminDomain.disabled = true
		DomainAdminPassword.disabled = true

		MachineObjectOU.disabled = true
		MachineObjectOUOptionalBtn.disabled = true
		MachineObjectOUOptional.disabled = true

		JoinWorkgroup.disabled = true

		ValidateDomainMembership = true

		' Don't do any more validation because this has been disabled
		Exit Function

	Else

		JoinDomain.disabled = not isDomain
		DomainAdmin.disabled = not isDomain
		DomainAdminDomain.disabled = not isDomain
		DomainAdminPassword.disabled = not isDomain

		MachineObjectOU.disabled = not isDomain
		MachineObjectOUOptionalBtn.disabled = not isDomain
		MachineObjectOUOptional.disabled = not isDomain

		JoinWorkgroup.disabled = isDomain

	End if


	' Check Warnings

	ValidateDomainMembership = ParseAllWarningLabels


	' Check domain settings (without validation of credentials)

	If IsDomain then

		' Make sure the join account details are specified

		If Trim(DomainAdmin.value) = "" or Trim(DomainAdminPassword.value) = "" or (Instr(DomainAdmin.Value, "@") = 0 and Trim(DomainAdminDomain.Value) = "") then
			MissingCredentials.style.display = "inline"
			ValidateDomainMembership = false
		End if


		' Check OU to make sure it is a valid format

		If MachineObjectOU.Value <> "" then

			' Make sure it starts with "OU=" or "OU " (equal could be preceeded by spaces)

			If Left(UCase(Trim(MachineObjectOU.Value)), 3) <> "OU=" and Left(UCase(Trim(MachineObjectOU.Value)), 3) <> "OU " then
				InvalidOU.style.display = "inline"
				ValidateDomainMembership = false
			End if
		End if

	End if


	' Check credentials

	If IsDomain and ValidateDomainMembership and (not window.event is Nothing) then
	
		' Only check credentials when the next button is clicked

		If window.event.srcElement is ButtonNext or window.event.KeyCode = 13 then

			oLogging.CreateEntry "Validate Domain Credentials [" & DomainAdminDomain.value & "\" & DomainAdmin.value & "]", LogTypeInfo

			If oEnvironment.Item("OSVersion") <> "WinPE" then

				' Check using ADSI (not possible in Windows PE)

				r = CheckCredentialsAD(DomainAdminDomain.value, DomainAdmin.value, DomainAdminDomain.value, DomainAdminPassword.value)
				If r <> TRUE then

					InvalidCredentials.innerText = "* Invalid credentials: " & r
					InvalidCredentials.style.display = "inline"
					ValidateDomainMembership = false

				End if

			ElseIf oEnvironment.Item("ValidateDomainCredentialsUNC") <> "" then

				' Check using ADSI (not possible in Windows PE)

				oLogging.CreateEntry "Validate Domain Credentials against UNC:  " & oEnvironment.Item("ValidateDomainCredentialsUNC")  , LogTypeInfo

				r = CheckCredentials( oEnvironment.Item("ValidateDomainCredentialsUNC") , DomainAdmin.value, DomainAdminDomain.value, DomainAdminPassword.value)

				oLogging.CreateEntry "Validate Domain Credentials against UNC:  result = " & r , LogTypeInfo

				If r <> TRUE then

					InvalidCredentials.innerText = "* Invalid credentials: " & r
					InvalidCredentials.style.display = "inline"
					ValidateDomainMembership = false

				End if

			End if

		End if

	End if


	' We need to clean up the keyboard hook

	If ValidateDomainMembership then
		document.body.onMouseDown = ""
	End if

End Function
