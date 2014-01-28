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


Function InitializeUserDataRestoreLocation

	' If the user data location is AUTO or NETWORK, reset it to none for bare metal	
	
	If UCase(Property("UserDataLocation")) = "AUTO" or UCase(Property("UserDataLocation")) = "NETWORK" then
		oProperties("UserDataLocation") = "NONE"
	End if


	' Make sure the right radio button is set

	If UCase(Property("UserDataLocation")) = "NONE" OR Property("UserDataLocation") = "" then
		UDRadio1.click
	Else
		StatePath.value = Property("UserDataLocation")
		UDRadio2.click
	End if
	
	if not isempty("UDShare") then
		StatePathT.value = Property("UUID")
		UDRadio2.click
	End if

End Function


'''''''''''''''''''''''''''''''''''''
'  Validate UserData Location
'''''''''''''''''''''''''''''''''''''

Function ValidateUserDataRestoreLocation

	StatePath.value = Property("UDShare")+"\"+StatePathT.value

	UDRadio2.Value = StatePath.Value

	document.GetElementByID("StatePathT").Disabled = not UDRadio2.Checked
	document.GetElementByID("StatePathBrowse").Disabled = not UDRadio2.Checked

	InvalidPath.style.display = "none"
	ValidateUserDataRestoreLocation = TRUE
	If UDRadio2.Checked and StatePath.value <> "" then

		If Left(StatePath.value, 2) = "\\" and len(StatePath.value) > 6 and ubound(split(StatePath.value,"\")) >= 3 then
			oUtility.ValidateConnection StatePath.value
		End if

		If (oFSO.FileExists(StatePath.value & "\USMT3.MIG" ) or oFSO.FileExists(StatePath.value & "\USMT.MIG" )) or ( oFSO.FileExists(StatePath.value & "\MIGSTATE.DAT" ) and _
			oFSO.FileExists(StatePath.value & "\catalog.mig" ) ) then

			' Just in case the user selects the USMT3 directory.
			StatePath.value = StatePath.value & "\.."

		End if

		If not (oFSO.FolderExists(StatePath.value & "\USMT3" ) or oFSO.FolderExists(StatePath.value & "\USMT" )) then
			ValidateUserDataRestoreLocation = FALSE
			InvalidPath.style.display = "inline"
		Elseif not (oFSO.FileExists(StatePath.value & "\USMT3\USMT3.MIG" ) or oFSO.FileExists(StatePath.value & "\USMT\USMT.MIG" )) and _
			not (oFSO.FileExists(StatePath.value & "\USMT3\MIGSTATE.DAT" ) or oFSO.FileExists(StatePath.value & "\USMT\MIGSTATE.DAT" )) and _
			not (oFSO.FileExists(StatePath.value & "\USMT3\catalog.mig" ) or oFSO.FileExists(StatePath.value & "\USMT\catalog.mig" )) then

			ValidateUserDataRestoreLocation = FALSE
			InvalidPath.style.display = "inline"
		End if



	End if

	ValidateUserDataRestoreLocation = ValidateUserDataRestoreLocation and ParseAllWarningLabels

End Function
