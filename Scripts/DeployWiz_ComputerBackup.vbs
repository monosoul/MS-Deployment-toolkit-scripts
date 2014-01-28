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


Function InitializeComputerBackupLocation

	' We muck around with the values, so we need to do some manual cleanup
	
	If Ucase(property("DeploymentType")) = "REPLACE" then
	
		CBRadio1.Disabled = TRUE
		AllowLocal.Checked = FALSE

		If UCase(property("ComputerBackupLocation")) = "NONE" or Ucase(oEnvironment.Item("ComputerBackupLocation")) = "NONE" then
			CBRadio3.click

		Else
			CBRadio2.click
			If UCase(property("ComputerBackupLocation")) <> "NETWORK" then
				DataPath.Value = property("ComputerBackupLocation")
			End if

		End if

	ElseIf UCase(property("ComputerBackupLocation")) = "" then

		If Property("BackupShare") <> ""AND Property("BackupDir") <> "" Then
			DataPath.value = Property("BackupShare") & "\" & Property("BackupDir")
			CBRadio2.click
		End If
		
	ElseIf UCase(property("ComputerBackupLocation")) = "AUTO" or Ucase(oEnvironment.Item("ComputerBackupLocation")) = "AUTO" then
		AllowLocal.Checked = TRUE
		CBRadio1.click

	ElseIf UCase(property("ComputerBackupLocation")) = "NONE" or Ucase(oEnvironment.Item("ComputerBackupLocation")) = "NONE" then
		CBRadio3.click

	ElseIf UCase(property("ComputerBackupLocation")) = "NETWORK" or Ucase(oEnvironment.Item("ComputerBackupLocation")) = "NETWORK" then
		CBRadio1.Disabled = TRUE
		AllowLocal.Checked = FALSE
		CBRadio2.click
		If Property("BackupShare") <> ""AND Property("BackupDir") <> "" Then
			DataPath.value = Property("BackupShare") & "\" & Property("BackupDir")
		End if

	Else
		DataPath.Value = property("ComputerBackupLocation")
		CBRadio2.click
	End if

	ValidateComputerBackupLocation

End function




'''''''''''''''''''''''''''''''''''''
'  Validate Computer Backup Location
'

Function ValidateComputerBackupLocation
	Dim HasErrors

	HasErrors = FALSE
	document.GetElementByID("CBRadio2").Value = document.GetElementByID("DataPath").Value

	document.GetElementByID("AllowLocal").Disabled = not CBRadio1.Checked

	document.GetElementByID("DataPath").Disabled = not CBRadio2.Checked
	document.GetElementByID("DataPathBrowse").Disabled = not CBRadio2.Checked

	ValidateComputerBackupLocation = ParseAllWarningLabels

End Function

