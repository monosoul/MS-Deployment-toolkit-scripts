' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_Validation.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Main Client Deployment Wizard Validation routines
' // 
' // ***************************************************************************

Option Explicit


'''''''''''''''''''''''''''''''''''''
'  Validate Password
'

Function ValidatePassword

	ValidatePassword = ParseAllWarningLabels

	NonMatchPassword.style.display = "none"
	If Password1.Value <> "" then
		If Password1.Value <> Password2.Value then
			ValidatePassword = FALSE
			NonMatchPassword.style.display = "inline"
		End if
	End if

	ButtonNext.Disabled = not ValidatePassword

End Function
