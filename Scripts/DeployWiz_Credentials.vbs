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


Dim UserID_isDirty
UserID_isDirty = FALSE

Function ValidateCredentials

	UserID_isDirty = TRUE
	ValidateCredentials = ParseAllWarningLabelsEx(userdomain, username )

End Function

Function ValidateCredentialsEx
	Dim r

	ValidateCredentialsEx = ValidateCredentials

	InvalidCredentials.style.display = "none"

	If ValidateCredentialsEx and oEnvironment.Item("OSVersion") <> "WinPE" then

			' Check using ADSI (not possible in Windows PE)

			r = CheckCredentials("", username.value, userdomain.value, userpassword.value)
			If r <> TRUE then

				InvalidCredentials.innerText = "* Invalid credentials: " & r
				InvalidCredentials.style.display = "inline"
				ValidateCredentialsEx = false

			End if


	End if

End function
