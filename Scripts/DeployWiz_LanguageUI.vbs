' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_LanguageUI.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Script methods used for the language UI (locale, timezone) pane
' // 
' // ***************************************************************************


Option Explicit


dim g_oXMLLanguageList

Function oXMLLanguageList

	If IsEmpty(g_oXMLLanguageList) then
		Set g_oXMLLanguageList = oUtility.LoadConfigFileSafe( "scripts\ListOfLanguages.xml" )
	End if
	Set oXMLLanguageList = g_oXMLLanguageList

End function

dim g_oPackageGroup

''''''''''''''''''''''''''''''''''''''

function AddLanguage( LanugageToAdd )

	Dim oLang, oOption2, sLangToAdd
	
	' sLangToAdd = GetParentLanguageFromLocale(LanugageToAdd)
	sLangToAdd = LanugageToAdd

	For each oLang in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE[IFLAGS='1' and SNAME]")

		If ucase(oUtility.SelectSingleNodeString(oLang,"SNAME")) = ucase(LanugageToAdd) then
			Set oOption2 = document.createElement("OPTION")
			oOption2.Text = oUtility.SelectSingleNodeString(oLang,"SENGDISPLAYNAME")
			If not oLang.SelectSingleNode("SNAME") is nothing then
				oOption2.Value = lcase(oUtility.SelectSingleNodeString(oLang,"SNAME"))
				UILanguage.add oOption2
			End if 
			
			Exit for
		End if

	Next

End function


Function Locale_Initialization

	Dim oItem
	Dim oXMLPackageList
	Dim FoundLocale
	Dim AllreadyAddedLanguages
	Dim aLangPack
	Dim oOption2
	Dim thisLocale
	
	' g_sPackageDialogBox = ""

	oLogging.CreateEntry "###### Locale_Initialization ###### " , LogTypeInfo
	

	' Disable fields as appropriate

	If UCase(oEnvironment.item("SkipLocaleSelection")) = "YES" then
		UILanguage.Disabled = true
		UserLocale_Edit.Disabled = true
		KeyboardLocale_Edit.Disabled = true
	ElseIf Left(oEnvironment.Item("ImageBuild"), 1) = "5" then
		UILanguage.Disabled = true
	End if


	' Add a Language for each package selected

	set oXMLPackageList = new ConfigFile
	oXMLPackageList.sFileType = "Packages"

	
	If IsInstallationUltimateEnterprise then
	
		' Add langauges allready installed on the image.
		AllreadyAddedLanguages = ""
		For each oItem in oEnvironment.ListItem("ImageLanguage").Keys
			If oItem <> "" then
				AddLanguage oItem
				AllreadyAddedLanguages = AllreadyAddedLanguages & vbTab & oItem
			End if
		Next

	End if 

	aLangPack = property("LanguagePacks")
	If not isArray(aLangPack) then
		' Force LanguagePacks variable as an array for non Ultimate/Enterprise builds.		
		aLangPack = array(aLangPack)
		oProperties.Item("LanguagePacks") = aLangPack
	End if
	
	For each oItem in aLangPack

		FoundLocale = ""
		If oItem = "DEFAULT" then
			' Skip...
		ElseIf not oXMLPackageList is nothing then

			If oXMLPackageList.FindAllItems.Exists(oItem) then
				FoundLocale = oUtility.SelectSingleNodeString(oXMLPackageList.FindAllItems.Item(oItem),"./Language")
			End if

		End if

		If FoundLocale <> "" and Instr(1,AllreadyAddedLanguages, FoundLocale, vbTextCompare ) = 0 then
			AddLanguage FoundLocale
			AllreadyAddedLanguages = AllreadyAddedLanguages & vbTab & FoundLocale
		End if

	Next
	
	If AllreadyAddedLanguages = "" then
	
		For each oItem in oEnvironment.ListItem("ImageLanguage").Keys
			If oItem <> "" then
				AddLanguage oItem
				AllreadyAddedLanguages = AllreadyAddedLanguages & vbTab & oItem
			End if
		Next

	End if 

	ForceLCase "UILanguage"
	ForceLCase "UserLocale"
	ForceLCase "KeyboardLocale"

	oLogging.CreateEntry "Languages Displayed: " & AllreadyAddedLanguages , LogTypeInfo
	oLogging.CreateEntry "UILanguage: " & property("UILanguage") , LogTypeVerbose
	
	' Populate the Locale
	For each oItem in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE[IFLAGS='1']")

		Set oOption2 = document.createElement("OPTION")
		oOption2.Text = oUtility.SelectSingleNodeString(oItem,"SENGDISPLAYNAME")

		oOption2.Value = lcase(oUtility.SelectSingleNodeString(oItem,"./SNAME") & ";" & oUtility.SelectSingleNodeString(oItem,"ILANGUAGE"))
		UserLocale_Edit.add oOption2
	Next

	PopulateElements

	' Get default Language and populate
	If UILanguage.Value <> "" then
		thisLocale = UILanguage.Value

	Elseif Property("UILanguage") <> "" then
		thisLocale = Property("UILanguage")
	
	ElseIf oEnvironment.Item("ImageLanguage001") <> "" then
		thisLocale = oEnvironment.Item("ImageLanguage001")
		
	Else
		thisLocale = GetDefaultInstallationLocaleString
		
	End if
	If IsEmpty(thisLocale) then
		thisLocale = "en-US" ' WinPE *may* not have the locale defined
	End if


	SetNewLanguageEx thisLocale
	SetNewLocaleEx thisLocale

End function

Function ForceLCase( sPropertyName )
	If Property(sPropertyName) <> lcase(Property(sPropertyName)) then
		If oProperties.Exists(sPropertyName) then
			oProperties.Item(sPropertyName) = lcase(oProperties.Item(sPropertyName))
		Else
			oProperties.Add  sPropertyName, lcase(Property(sPropertyName))
		End if 
		
	End if 

End function

Function SetNewLanguage
	If (Not UILanguage.Disabled = true) then
		If oProperties.exists("UserLocale") then
			oProperties.remove "UserLocale"
		End if
		SetNewLanguageEx UILanguage.Value
		SetNewLocale
	End if
End Function


Function SetNewLocale
	If (Not KeyboardLocale_Edit.Disabled = true) then
		If instr(1,UserLocale_Edit.Value,";",vbTextCompare) <> 0 then
			If oProperties.exists("KeyboardLocale") then
				oProperties.remove "KeyboardLocale"
			End if
			SetNewLocaleEx mid(UserLocale_Edit.Value,1,instr(1,UserLocale_Edit.Value,";",vbTextCompare) - 1)
		End if
	End if
End Function


Function SetNewLanguageEx( thisLocale )
	Dim LCID
	
	oLogging.CreateEntry "SetNewLanguageEx " & thisLocale, LogTypeVerbose

	' Get the default UserLocale
	If Property("UserLocale") <> "" then
		UserLocale_Edit.Value = lcase(Property("UserLocale") & ";" & GetLCIDFromSName( Property("UserLocale") ))
	ElseIf instr(1,Property("KeyboardLocale"),":",vbTextCompare) <> 0 and len(trim(Property("KeyboardLocale"))) = len("0000:12345678") then
		LCID = left(Property("KeyboardLocale"),instr(1,Property("KeyboardLocale"),":",vbTextCompare)-1)
		UserLocale_Edit.Value = lcase(GetSNameFromLCID(LCID) & ";" & LCID)
	Else
		UserLocale_Edit.Value = lcase(thisLocale & ";" & GetLCIDFromSName( thisLocale ))
	End if
	
	oLogging.CreateEntry "UserLocale : " & Property("UserLocale") & " - " &   UserLocale_Edit.Value, LogTypeVerbose
	
	If Property("ImageLanguage001") = "" then
		oProperties.Add  "ImageLanguage001", thisLocale
	End if


End function


Function SetNewLocaleEx (thisLocale)
	Dim sKeyboard
	Dim sItem
	Dim sNew
	oLogging.CreateEntry "SetNewLocaleEx " & thisLocale & " - " & Property("KeyboardLocale"), LogTypeVerbose
	
	' Set the default Keyboard

	sNew = ""
	If Property("KeyboardLocale") <> "" then
		sKeyboard = Property("KeyboardLocale")
		If instr(1,sKeyboard,";",vbTExtCompare) <> 0 then
			' Use the 1st instance of a ; delimited array that contains a :
			for each sItem in split(sKeyboard,";")
				If instr(1,sItem,":",vbTextCompare) <> 0 then
					sKeyboard = trim(sItem)
					exit for
				End if
			next
			If isempty(sKeyboard) then
				sKeyboard = trim(left(sKeyboard,instr(1,sKeyboard,";",vbTExtCompare)-1))
			End if 
		End if
		If instr(1,sKeyboard,":",vbTextCompare) <> 0 and len(sKeyboard) = len("0000:12345678") then
			' KeyboardLocale appears to be in the format 0409:00000409 format
			sNew = mid(sKeyboard,instr(1,sKeyboard,":",vbTextCompare)+1)
		ElseIf instr(1,sKeyboard,"-",vbTextCompare) <> 0 Then
			sNew = right("00000000" & GetLCIDFromSName(sKeyboard),8)
		End if 
	End if

	If sNew = "" Then
		sNew = right("00000000" & GetKeyboardFromSName(thisLocale),8)
	End if 
	KeyboardLocale_Edit.Value = lcase(sNew)
	
	If KeyboardLocale_Edit.Value = "" then
		KeyboardLocale_Edit.Value = "00000409"
	End if

	oLogging.CreateEntry "KeyboardLocale: " & Property("KeyboardLocale") & " - " & KeyboardLocale_Edit.Value & " - " & sNew, LogTypeVerbose

End function


''''''''''''''''''''''''''''''''

Function GetDefaultInstallationLocaleString

	Dim oItem


	' First see if a UserLocal value was specified

	If Property("UserLocale") <> "" then

		Set oItem = oXMLLanguageList.SelectSingleNode("//LOCALEDATA/LOCALE[@ID = '" & lcase(Property("UserLocale")) & "']/SNAME")

		If not oItem is nothing then
			GetDefaultInstallationLocaleString = oItem.Text
			Exit Function
		End if

	End if


	' No, so get the default locale

	Set oItem = oXMLLanguageList.SelectSingleNode("//LOCALEDATA/LOCALE[@ID = '" & lcase(right("0000" & hex(GetLocale),4)) & "']/SNAME")

	If not oItem is nothing then
		GetDefaultInstallationLocaleString = oItem.Text
	End if


End function


Function GetDefaultInstallationLanguageString

	GetDefaultInstallationLanguageString = GetParentLanguageFromLocale(GetDefaultInstallationLocaleString)

End function


Function GetParentLanguageFromLocale( Locale)
	Dim oLocale

	For each oLocale in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE/SNAME")

		If UCase(Locale) = UCase(oLocale.text) then
			GetParentLanguageFromLocale = oUtility.SelectSingleNodeString(oLocale.ParentNode,"SISO639LANGNAME")
			Exit for
		End if

	Next

End function

Function GetSNameFromLCID ( LCID )

	GetSNameFromLCID  = oUtility.SelectSingleNodeString(oXMLLanguageList,"/LOCALEDATA/LOCALE[@ID='" & lcase(LCID) & "']/SNAME")
	
End function 

Function GetKeyboardFromSName ( sName )

	Dim oLocale

	For each oLocale in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE/SNAME")

		If UCase(sName) = UCase(oLocale.text) then
			GetKeyboardFromSName = oUtility.SelectSingleNodeString(oLocale.ParentNode,"DEFAULTKEYBOARD")
			Exit for
		End if

	Next

End Function 

Function GetLCIDFromSName ( sName )

	Dim oLocale

	For each oLocale in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE/SNAME")

		If UCase(sName) = UCase(oLocale.text) then
			GetLCIDFromSName = oUtility.SelectSingleNodeString(oLocale.ParentNode,"ILANGUAGE")
			Exit for
		End if

	Next

End Function 

''''''''''''''''''''''''''''''''

Function IsInstallationUltimateEnterprise

	IsInstallationUltimateEnterprise = oUtility.IsHighEndSKUEx( oEnvironment.Item("ImageFlags") )

End function



'''''''''''''''''''''''''''''''''''''
'  Validate LanguagePack
'

''''''''''''''''''''''''''''''''''''''

Function Locale_Validation

	Dim iSplit

	Locale_Validation = TRUE

	If not UILanguage.Disabled then
		UILanguage_err.style.display = "none"
		If UILanguage.SelectedIndex = -1 then
			UILanguage_Err.style.display = "inline"
			Locale_Validation = FALSE
		End if
	End if

	If not UserLocale_Edit.Disabled then
		UserLocale_Err.style.display = "none"
		If UserLocale_Edit.SelectedIndex = -1 then
			UserLocale_Err.style.display = "inline"
			Locale_Validation = FALSE
		End if
	End if

	If not KeyboardLocale_Edit.Disabled then
		KeyboardLocale_Err.style.display = "none"
		If KeyboardLocale_Edit.SelectedIndex = -1 then
			KeyboardLocale_Err.style.display = "inline"
			Locale_Validation = FALSE
		End if
	End if	

	If not Locale_Validation then
		Exit Function
	End if
	
	iSplit = instr(1,UserLocale_Edit.Value,";",vbTextCompare)
	TestAndLog iSplit <> 0 , "Verify UserLocale_Edit contains Comma Delimiter: " & UserLocale_Edit.Value

	If (Not KeyboardLocale_Edit.Disabled = true) then
		If instr(1,UserLocale_Edit.Value,";",vbTextCompare) <> 0 then
			' Take the LCID From UserLocale and add it to the KeyboardLocale
			KeyboardLocale.Value = UserLocale_Edit.Value & ":" & KeyboardLocale_Edit.Value
		Else
			' Some kind of Error
			KeyboardLocale.Value = right("0000" & hex(GetLocale),4) & ":" & KeyboardLocale_Edit.Value
		End if
	End if
	
	If (Not UserLocale_Edit.Disabled = true) then
		If iSplit <> 0 then
			UserLocale.Value = mid(UserLocale_Edit.Value,1,instr(1,UserLocale_Edit.Value,";",vbTextCompare) - 1)
		Else
			UserLocale.Value = UserLocale_Edit.Value
		End if
	End if

End function



'''''''''''''''''''''''''''''''''''''
'  Timezone functions
'
'''''''''''''''''''''''''''''''''''''

Function SetTimeZoneValue
	' When the user selects a value in the TimeZoneList we must populate the hidden Text Values
	Dim TimeSplit

	TimeSplit = split( TimeZoneList.value , ";" )
	If ubound(TimeSplit) < 1 then
	ElseIf not isNumeric(TimeSplit(0)) then
	Else
		TimeZoneName.Value = TimeSplit(1)
		TimeZone.Value = TimeSplit(0)
	End if

End function


Function TimeZone_Initialization

	Dim TimeZone, i, TimeSplit, Item, test

	' Disable fields as appropriate

	If UCase(oEnvironment.item("SkipTimeZone")) = "YES" then
		TimeZoneList.Disabled = true
	End if


	' If either of the TimeZone Properties have been set, then select the coresponding list item.

	If Property("TimeZone") <> "" or Property("TimeZoneName") <> "" then
		For i = 0 to TimeZoneList.Options.Length - 1

			TimeSplit = split( TimeZoneList.Options(i).value , ";" )

			If ubound(TimeSplit) >= 1 then
				If Property("TimeZone") <> "" then
					If IsNumeric(Property("TimeZone")) then
						' Check Windows XP style Name
						If CInt(Property("TimeZone")) = cint(TimeSplit(0)) then
							TimeZoneList.SelectedIndex = i
							SetTimeZoneValue
							Exit function
						End if
					Else
						' Check Windows Vista Style Name
						If UCase(Property("TimeZone")) = UCase(TimeSplit(1)) then
							TimeZoneList.SelectedIndex = i
							SetTimeZoneValue
							Exit function
						End if
					End if
				ElseIf Property("TimeZoneName") <> "" then
					' Check Windows Vista Style Name
					If UCase(Property("TimeZoneName")) = UCase(TimeSplit(1)) then
						TimeZoneList.SelectedIndex = i
						SetTimeZoneValue
						Exit function
					End if
				End if
			End if
		Next
	End if


	' Extract out the current TimeZone

	For each TimeZone in objWMI.InstancesOf("Win32_TimeZone")
		Exit for ' Take the first entry and break out of loop
	Next

	If IsEmpty(TimeZone) then
		Exit function
	End if


	'Try to match the timezone against the current Timezone Name

	For i = 0 to TimeZoneList.Options.Length - 1

		TimeSplit = split( TimeZoneList.Options(i).value , ";" )

		If UBound(TimeSplit) >= 1 then
			' Compare the Description
			If UCase(TimeZoneList.Options(i).Text) = UCase(TimeZone.Description) then
				TimeZoneList.SelectedIndex = i
				SetTimeZoneValue
				Exit function
			End if

			' See if there is a match in the alternate description or other Values
			For each test in array(TimeZone.Description,TimeZone.StandardName)
				If test <> "" then
					For each item in TimeSplit
						If Item <> "" then

							If UCase(test) = UCase(Item) then
								TimeZoneList.SelectedIndex = i
								SetTimeZoneValue
								Exit function
							End if

						End if
					Next
				End if
			Next

		End if
	Next


	' Try to match against the closest GMT value (This *May* select an entry that is not *exact*)

	For i = 0 to TimeZoneList.Options.Length - 1

		test = Instr(1, TimeZone.Description," ")
		If test <> 0 then

			If left(TimeZone.Description, test) = left(TimeZoneList.Options(i).Text, test) then
				TimeZoneList.SelectedIndex = i
				SetTimeZoneValue
				Exit function
			End if

		End if
	Next

End function


Function TimeZone_Validation

	' Return true if disabled or if a value has been specified

	If UCase(oEnvironment.item("SkipTimeZone")) = "YES" then
		TimeZone_Validation = True
	Else
		TimeZone_Validation = TimeZoneName.Value <> "" and TimeZone.Value <> ""
	End if

End function
