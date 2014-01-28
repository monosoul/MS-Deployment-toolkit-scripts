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


Function InitializeCapture

	If Ucase(Property("ComputerBackupLocation")) = "NETWORK" OR Ucase(oEnvironment.Item("ComputerBackupLocation")) = "NETWORK" Then
		If Property("BackupShare") <> ""AND Property("BackupDir") <> "" Then
			ComputerBackupLocation.value = Property("BackupShare") & "\" & Property("BackupDir")
		ElseIF oEnvironment.Item("BackupShare") <> "" AND oEnvironment.Item("BackupDir") <> "" Then
			ComputerBackupLocation.value = oEnvironment.Item("BackupShare") & "\" & oENvironment.Item("BackupDir")
		Else
			ComputerBackupLocation.value = Property("DeployRoot") & "\Captures"
		End If
	End If
	If Property("ComputerBackupLocation") = "" then
		ComputerBackupLocation.value = Property("DeployRoot") & "\Captures"
	End if
	If Property("BackupFile") = "" then
		BackupFile.value = Property("TaskSequenceID") & ".wim"
	End if
	
	RMPropIfFound("BdePin")
	RMPropIfFound("BdeModeSelect1")
	RMPropIfFound("BdeModeSelect2")
	RMPropIfFound("BdeKeyLocation")
	RMPropIfFound("OSDBitLockerWaitForEncryption")
	RMPropIfFound("BdeRecoveryKey")
	RMPropIfFound("BdeRecoveryPassword")
	RMPropIfFound("BdeInstallSuppress")

	
End Function



'''''''''''''''''''''''''''''''''''''
'  Validate Capture
'

Function ValidateCaptureLocation

	InvalidCaptureLocation.style.display = "none"
	ValidateCaptureLocation = true

	If not CaptureRadio1.Checked then
		ComputerBackupLocation.value = ""
		BackupFile.Value = ""
		RMPropIfFound("ComputerBackupLocation")
		RMPropIfFound("BackupFile")
		Exit Function
	End if

	If Left(ComputerBackupLocation.value, 2) = "\\" and len(ComputerBackupLocation.value) > 6 and ubound(split(ComputerBackupLocation.value,"\")) >= 3 then

		If not oUtility.ValidateConnection(ComputerBackupLocation.value) = Success then
				InvalidCaptureLocation.style.display = "inline"
				ValidateCaptureLocation = FALSE
		End if

	Else
		InvalidCaptureLocation.style.display = "inline"
		ValidateCaptureLocation = FALSE
	End if

End Function

Function ValidateCapture

	document.GetElementByID("ComputerBackupLocation").Disabled = not CaptureRadio1.Checked
	document.GetElementByID("BackupFile").Disabled = not CaptureRadio1.Checked

	if not CaptureRadio4.Checked then

		RMPropIfFound("BdeInstall")
		RMPropIfFound("BdeInstallSuppress")

	End if

	ValidateCapture = ParseAllWarningLabels

End Function
