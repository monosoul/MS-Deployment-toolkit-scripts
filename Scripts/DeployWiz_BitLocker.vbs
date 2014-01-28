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


Function InitializeBDE

	Dim sType
	
	sType = ucase(Property("BdeInstall"))
	If sType = "" then
		sType = ucase(Property("OSDBitLockerMode"))
	End if
	
	Select Case sType
	Case "TPM"
		BdeRadio2.checked = true
		BdeModeRadio1.checked = true
	Case "TPMKEY"
		BdeRadio2.checked = true
		BdeModeRadio2.checked = true
		If Property("BdeKeyLocation") <> "" then
			BdeModeSelect1.Value = ucase(Property("BdeKeyLocation"))
		ElseIf Property("OSDBitLockerStartupKeyDrive") <> "" then
			BdeModeSelect1.Value = ucase(Property("OSDBitLockerStartupKeyDrive"))
		End if
	Case "KEY"
		BdeRadio2.checked = true
		BdeModeRadio3.checked = true
		If Property("BdeKeyLocation") <> "" then
			BdeModeSelect2.Value = ucase(Property("BdeKeyLocation"))
		ElseIf Property("OSDBitLockerStartupKeyDrive") <> "" then
			BdeModeSelect2.Value = ucase(Property("OSDBitLockerStartupKeyDrive"))
		End if
	Case "TPMPIN"
		BdeRadio2.checked = true
		BdeModeRadio4.checked = true
	Case Else
		BdeRadio1.Checked = true
	End Select

	If UCase(Property("BdeRecoveryKey")) = "AD" or UCase(Property("OSDBitLockerCreateRecoveryPassword")) = "AD" Then
		ADButton1.checked = True
	Else
		ADButton2.Checked = True
	End if 
	
	WaitForEncryption.checked = ucase(Property("OSDBitLockerWaitForEncryption")) = "TRUE" or  ucase(Property("BdeWaitForEncryption")) = "TRUE"

	BdeInstallSuppress.value = "YES"
End Function



Function ValidateBDE

	Dim regEx

	

	' Enable and disable

	If BDERadio2.checked then

		' Enable second set of radio buttons

		BdeModeRadio1.disabled = false
		BdeModeRadio2.disabled = false
		BdeModeRadio3.disabled = false
		BdeModeRadio4.disabled = false

		BdePin.disabled = false
		ADButton1.disabled = false
		ADButton2.disabled = false

		WaitForEncryption.disabled = false
	Else

		' Disable second set of radio buttons

		BdeModeRadio1.disabled = true
		BdeModeRadio2.disabled = true
		BdeModeRadio3.disabled = true
		BdeModeRadio4.disabled = true


		BdeModeSelect1.disabled = true
		BdeModeSelect2.disabled = true

		BdePin.disabled = true

		ADButton1.disabled = true
		ADButton2.disabled = true

		WaitForEncryption.disabled = true
	End if




	' Set BdeInstall based on choices

	If BDERadio2.checked then
		BdeInstallSuppress.value = "NO"

		' Mode/location
		If BdeModeRadio1.checked then
			BdeInstall.value = "TPM"
			BdePin.disabled = true
		ElseIf BdeModeRadio2.checked then
			BdeInstall.value = "TPMKey"
			OSDBitLockerStartupKeyDrive.value = BdeModeSelect1.Value
			BdeModeSelect1.disabled = false
			BdeModeSelect2.disabled = true
			BdePin.disabled = true
		ElseIf BdeModeRadio3.checked then
			BdeInstall.value = "Key"
			OSDBitLockerStartupKeyDrive.value = BdeModeSelect2.Value
			BdeModeSelect1.disabled = true
			BdeModeSelect2.disabled = false
			BdePin.disabled = true
		Else
			BdeInstall.value = "TPMPin"
			BdeModeSelect1.disabled = true
			BdeModeSelect2.disabled = true
			BdePin.disabled = false
		End if


		If ADButton1.checked Then
			BdeRecoveryKey.value = "AD"
		Else
			BdeRecoveryKey.value = ""
		End if

		OSDBitLockerWaitForEncryption.value = WaitForEncryption.checked

	Else ' IF BDERadio1.Checked then
	
		BdeInstall.value = ""
		BdeInstallSuppress.value = "YES"
		
	End if

	ValidateKey
	
	' Scan required fields

	ValidateBDE = ParseAllWarningLabels
	
End Function

Function ValidateKey

	InvalidKey.style.display = "none"
	ValidateKey = TRUE

	If not BdeModeRadio4.Checked or not BDERadio2.Checked then
		BdePin.value = ""
		Exit Function
	End if
	
	If len(BdePin.value) = 0 Then
	
		InvalidKey.innerText = "* Required (MISSING)"
		InvalidKey.style.display = "inline"
		ValidateKey = FALSE
		
	ElseIf Not IsEmpty(oEnvironment.item("BDEPinMinLength")) And isNumeric(oEnvironment.item("BDEPinMinLength")) then

		if len(BdePin.value) < cLng(oEnvironment.item("BDEPinMinLength")) then
			InvalidKey.innerText = "* Pin must be " & oEnvironment.item("BDEPinMinLength") & " charaters or longer"
			InvalidKey.style.display = "inline"
			ValidateKey = FALSE
		End if
		
	ElseIf (oEnvironment.item("BDEAllowAlphaNumericPin") <> UCase("YES")) and Not IsNumeric(BdePin.value) then
	
		InvalidKey.innerText = "* Pin must be numeric"
		InvalidKey.style.display = "inline"
		ValidateKey = FALSE

	End if

End Function
