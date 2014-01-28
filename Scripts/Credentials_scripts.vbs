' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      Credentials_scripts.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Scripts to initialize and validate credential wizard
' // 
' // ***************************************************************************

Option explicit

Dim sValidateUNC
Dim sValidateDomain
Dim bDoNotSaveParameters
Dim bLeaveShareOpen


Function UserCredentialsInitialize
	Dim oArguments
	Dim SWbemObject
	Dim sNewDomain
	Dim aUNCPath


	' Parse the Command Line

	sValidateDomain = oUtility.Arguments.Item("ValidateAgainstDomain")
	aUNCPath = Split(oUtility.Arguments.Item("ValidateAgainstUNCPath"),"\")
	bDoNotSaveParameters = oUtility.Arguments.Exists("DoNotSave")
	bLeaveShareOpen = oUtility.Arguments.Exists("LeaveShareOpen")

	oEnvironment.Item("UserCredentials") = CStr(FALSE)

	If UBound(aUNCPath) > 2 then
		If aUNCPath(0)="" and aUNCPath(1)="" and aUNCPath(2)<>"" and aUNCPath(3)<>"" then
			sValidateUNC = "\\" & aUNCPath(2) & "\" & aUNCPath(3)
			oLogging.CreateEntry "Validate Against UNC: " & sValidateUNC , LogTypeInfo
		End if
	End if

	If not IsEmpty(sValidateDomain) then

		oLogging.CreateEntry "Validate Against Domain: " & sValidateDomain , LogTypeInfo

	ElseIf not IsEmpty(sValidateUNC) then

		sValidateDomain = GetDomainDefault
		oLogging.CreateEntry "Validate Against UNC: " & sValidateUNC , LogTypeInfo
		' Do we need to do any checking for IP address ( as compared to server names )

	Else

		sValidateDomain = GetDomainDefault

	End if

	If UserName.Value = "" and oNetwork.UserName <> "" and oNetwork.UserName <> "SYSTEM" then
		UserName.Value = oNetwork.UserName
	End if

	If oProperties("userdomain") = "" and sValidateDomain <> "" then
		userdomain.Value = sValidateDomain
	ElseIf oproperties("userdomain") = "" and oNetwork.UserDomain <> "" then
		userdomain.Value = oNetwork.UserDomain
	End if

End function


Function ValidateCredentials
	Dim r

	InvalidCredentials.style.display = "none"
	ValidateCredentials = ParseAllWarningLabelsEx(userdomain, username)

	If ValidateCredentials then

		r = CheckCredentials(sValidateUNC, UserName.value, userdomain.Value, userpassword.value)
		If r <> TRUE then
			InvalidCredentials.innerText = "* Invalid credentials: " & r
			InvalidCredentials.style.display = "inline"
			ValidateCredentials = false
		End if

	End if

End Function
