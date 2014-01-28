' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      WelcomeWiz_Choice.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Scripts for welcome wizard choice pane
' // 
' // ***************************************************************************


Option Explicit


Dim sSelectedBtn

Function GetValueFromID( oItem ) 

	Select Case oItem.ID
		Case buttonitem1.ID
			GetValueFromID = "DEPLOYWIZARD"
		Case buttonitem2.ID
			GetValueFromID = "RECOVERY"
		Case buttonitem3.ID
			GetValueFromID = "DART"
		Case buttonitem4.ID
			GetValueFromID = "COMMANDPROMPT"
	End select

End Function 


Function RunSelCmd

	Dim sValue
	sValue = GetValueFromID(window.event.srcElement)

	Select Case (window.event.type)
	
		Case "mouseout", "deactivate"
		
			If window.event.srcElement.ID <> sSelectedBtn then
				window.event.srcElement.style.backgroundimage = "url(btnout.png)"
			End if
		
		Case "mouseover"
		
			If window.event.srcElement.ID <> sSelectedBtn then
				window.event.srcElement.style.backgroundimage = "url(btnover.png)"
			End if

		Case "activate"
		
			ActivateItem window.event.srcElement
		
		Case "click", "dblclick"
			ActivateItem window.event.srcElement
			ButtonNextClick
			
	End Select

End function


Function ActivateItem ( oItemNew ) 

	if sSelectedBtn <> "" then
		document.GetElementByID(sSelectedBtn).style.backgroundimage = "url(btnout.png)"
	End if
	oItemNew.style.backgroundimage = "url(btnsel.png)"

	sSelectedBtn = oItemNew.ID
	oItemNew.Focus

End function


Sub KeyHandlerCustom

	if window.event.srcElement.tagName = "INPUT" then
		KeyHandler
		exit sub
	End if

	select case window.event.KeyCode

		Case 40  ' Down

			If window.event.srcElement.ID = "buttonItem1" and buttonItem2.style.display <> "none" then
				ActivateItem buttonItem2
			Elseif (window.event.srcElement.ID = "buttonItem1" or window.event.srcElement.ID = "buttonItem2") and buttonItem3.style.display <> "none" then
				ActivateItem buttonItem3
			Else
				ActivateItem buttonItem4
			End if

		Case 38  ' Up		

			If window.event.srcElement.ID = "buttonItem4" and buttonItem3.style.display <> "none" then
				ActivateItem buttonItem3
			ElseIf (window.event.srcElement.ID = "buttonItem4" or window.event.srcElement.ID = "buttonItem3") and buttonItem2.style.display <> "none" then
				ActivateItem buttonItem2
			Else
				ActivateItem buttonItem1
			End if

		End select
	
End sub


Function RunSelectedCommand 

	Dim sCmd

	Select case GetValueFromID(document.GetElementByID(sSelectedBtn))

		Case "RECOVERY"

			document.body.style.cursor = "Wait"
			sCmd = "x:\sources\recovery\RecEnv.exe"
			oShell.Run sCmd, 1, true
			document.body.style.cursor = "default"
			RunSelectedCommand = false

		Case "DART"

			document.body.style.cursor = "Wait"
			sCmd = "x:\sources\recovery\tools\MSDartTools.exe"
			oShell.Run sCmd, 1, true
			document.body.style.cursor = "default"
			RunSelectedCommand = false

		Case "COMMANDPROMPT"

			oEnvironment.Item("WizardComplete") = "N"
			window.Close
			Exit function

		Case else ' "DEPLOYWIZARD"

			RunSelectedCommand = true

	End select


End function


Function SafeRegRead( KeyValue )
   on error resume next
      SafeRegRead = oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE\KeyboardLayouts\" & GetLocale & "\" & KeyValue)
   on error goto 0
End function


Function WelcomeWizard_CustomInitializationCloseout
	buttonItem1.focus
End function 


Function WelcomeWizard_CustomInitialization
	Dim KeyboardLocale
	Dim FoundPrimary
	Dim sName, sID, i, j, Insert, oOption


	' Set the window title

	If oEnvironment.Item("_SMSTSOrgName") <> "" then
		document.title = oEnvironment.Item("_SMSTSOrgName")
	Else
		document.title = "Microsoft Deployment Toolkit"
	End if


	FoundPrimary = False
	ActivateItem buttonItem1

	document.body.onkeyDown = getref("KeyHandlerCustom")
	MyContentArea.style.backgroundimage = "url(WelcomeWiz_Background.jpg)"


	' Disable buttons for items not present in the boot image
	
	If not oFSO.FileExists("x:\sources\recovery\RecEnv.exe") then
		buttonItem2.Style.display = "none"
	End if

	If not oFso.FileExists("x:\sources\recovery\tools\MSDartTools.exe") then
		buttonItem3.Style.display = "none"
	End if


	' Test for the 1st registry entry

	if isempty(SafeRegRead( "0\Name" )) then
		' Not Found, run WpeUtil again
		oLogging.CreateEntry "Run WPEUtil.exe ListKeyboardLayouts " & GetLocale, LogTypeInfo
		oShell.Run "wpeutil.exe ListKeyboardLayouts " & GetLocale, 0, TRUE
	end if

	if isempty(SafeRegRead( "0\Name" )) or isempty(SafeRegRead( "0\ID" ))then
		' Still not found, some kind of problem with wpeutil.exe
		oLogging.CreateEntry "Could not enumerate Keyboard list through WPEUtil.exe", LogTypeWarning
	end if

	KeyboardLocale = oEnvironment.Item("KeyboardLocalePE")
	if KeyboardLocale = "" then
		KeyboardLocale = hex(GetLocale)
		while len(KeyboardLocale) < 4
			KeyboardLocale = "0" & KeyboardLocale
		wend
		KeyboardLocale = KeyboardLocale & ":0000" & KeyboardLocale
	end if


	i = 0
	sName = SafeRegRead( i & "\Name" )
	sID = SafeRegRead( i & "\ID" )
	do while not isempty(sName) and not isempty(sID)

		Insert = -1  ' Default

		for j = 0 to WinPEKeyboard.options.length - 1
			if StrComp(sName,WinPEKeyboard.Options(j).Text,VbTextCompare) < 0 then
				Insert = j
				exit For
			end if
		next

		' Skip if pre-existing
		for j = 0 to WinPEKeyboard.options.length - 1
				if WinPEKeyboard.options(j).value = sID then
				WinPEKeyboard.options(j).Selected = sID = KeyboardLocale
				Insert = empty
				exit for
			end if
		next

		' Add entry to the display.
		if not isempty(Insert) then
			set oOption = document.CreateElement("OPTION")

			if ucase(sID) = ucase(KeyboardLocale) then
				FoundPrimary = True
				oOption.Selected = True
			elseif FoundPrimary = False and ucase(right(sID,8)) = ucase(right(KeyboardLocale,8)) then
				oOption.Selected = True
			end if
			oOption.text = sName
			oOption.Value = sID
			WinPEKeyboard.Add oOption, Insert
		end if

		i = i + 1
		sName = SafeRegRead( i & "\Name" )
		sID = SafeRegRead( i & "\ID" )
	loop

end function


Function ConfigureStaticIP
	oShell.run "MSHTA.exe " & oUtility.ScriptDir & "\Wizard.hta /definition:NICSettings_Definition_ENU.xml"
End function
