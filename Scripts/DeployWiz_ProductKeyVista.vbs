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


Function InitializeProductKey

	' Figure out how to initialize the pane.

	If Property("ProductKey") <> "" or Left(Property("ImageBuild"), 1) < "6" then
		locProductKey.disabled = false
		locProductKey.value = Property("ProductKey")
		ProductKey.value = locProductKey.value
		If Left(Property("ImageBuild"), 1) >= "6" then
			PKRadio3.click
			locOverrideProductKey.disabled = true
			OverrideProductKey.value = ""
		End if
	ElseIf Property("OverrideProductKey") <> "" then
		PKRadio2.click
		locOverrideProductKey.disabled = false
		locProductKey.disabled = true
		locOverrideProductKey.value = Property("OverrideProductKey")
		OverrideProductKey.value = locOverrideProductKey.value
		ProductKey.value = ""
	Else
		PKRadio1.click
		locOverrideProductKey.disabled = true
		locProductKey.disabled = true
		ProductKey.value = ""
		OverrideProductKey.value = ""
	End if

End Function

Function ValidateProductKey

	ValidateProductKey = False

	If Left(Property("ImageBuild"), 1) < "6" then

		' Make sure the product key is valid

		If locProductKey.value = "" then
			PKBlank.style.display = "inline"
			PKInvalid.style.display = "none"
		ElseIf IsEmpty(GetProductKey(locProductKey.value)) then
			PKBlank.style.display = "none"
			PKInvalid.style.display = "inline"
		Else
			PKBlank.style.display = "none"
			PKInvalid.style.display = "none"
			ProductKey.value = GetProductKey(locProductKey.value)
			ValidateProductKey = True
		End if

	ElseIf PKRadio1.checked then

		locOverrideProductKey.disabled = true
		locProductKey.disabled = true

		OverrideBlank.style.display = "none"
		OverrideInvalid.style.display = "none"
		PKBlank.style.display = "none"
		PKInvalid.style.display = "none"

		ProductKey.value = ""
		OverrideProductKey.value = ""

		ValidateProductKey = True


	ElseIf PKRadio2.checked then

		locOverrideProductKey.disabled = false
		locProductKey.disabled = true

		PKBlank.style.display = "none"
		PKInvalid.style.display = "none"


		' Make sure the MAK key is valid

		If locOverrideProductKey.value = "" then
			OverrideBlank.style.display = "inline"
			OverrideInvalid.style.display = "none"
		ElseIf IsEmpty(GetProductKey(locOverrideProductKey.value)) then
			OverrideBlank.style.display = "none"
			OverrideInvalid.style.display = "inline"
		Else
			OverrideBlank.style.display = "none"
			OverrideInvalid.style.display = "none"
			OverrideProductKey.value = GetProductKey(locOverrideProductKey.value)
			ProductKey.value = ""
			ValidateProductKey = True
		End if

	Else

		locOverrideProductKey.disabled = true
		locProductKey.disabled = false

		OverrideBlank.style.display = "none"
		OverrideInvalid.style.display = "none"


		' Make sure the product key is valid

		If locProductKey.value = "" then
			PKBlank.style.display = "inline"
			PKInvalid.style.display = "none"
		ElseIf IsEmpty(GetProductKey(locProductKey.value)) then
			PKBlank.style.display = "none"
			PKInvalid.style.display = "inline"
		Else
			PKBlank.style.display = "none"
			PKInvalid.style.display = "none"
			ProductKey.value = GetProductKey(locProductKey.value)
			OverrideProductKey.value = ""
			ValidateProductKey = True
		End if

	End if

End Function


const PRODUCT_KEY_TEST = "([0-9A-Z]+)?[^0-9A-Z]*([0-9A-Z]{5})[^0-9A-Z]?([0-9A-Z]{5})[^0-9A-Z]?([0-9A-Z]{5})[^0-9A-Z]?([0-9A-Z]{5})[^0-9A-Z]?([0-9A-Z]{5})[^0-9A-Z]*([0-9A-Z]+)?" '


Function GetProductKey( pk )

	Dim regEx, match

	Set regEx = New RegExp
	regEx.Pattern = PRODUCT_KEY_TEST
	regex.IgnoreCase = TRUE

	For each match in regEx.Execute( UCase(pk) )
		If IsEmpty(match.SubMatches(0)) and IsEmpty(match.SubMatches(6)) then
			GetProductKey = ucase( match.SubMatches(1) & "-" & match.SubMatches(2) & "-" & _
			match.SubMatches(3) & "-" & match.SubMatches(4) & "-" & match.SubMatches(5) )
		End if
		Exit function
	Next

End function


Function AssignProductKey

	If not IsEmpty(GetProductKey(locProductKey.value)) then
		locProductKey.value = GetProductKey(locProductKey.value)
	End if
	If Left(Property("ImageBuild"), 1) >= "6" then
		If not IsEmpty(GetProductKey(locOverrideProductKey.value)) then
			locOverrideProductKey.value = GetProductKey(locOverrideProductKey.value)
		End if
	End if

End Function


