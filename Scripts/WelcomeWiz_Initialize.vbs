' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      WelcomeWiz_Initialization.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Script logic for initialization/progress display
' // 
' // ***************************************************************************


Option Explicit


Function InitializeProgress

	' Set the background image

	MyContentArea.style.backgroundimage = "url(WelcomeWiz_Background.jpg)"


	' Set the window title

	If oEnvironment.Item("_SMSTSOrgName") <> "" then
		document.title = oEnvironment.Item("_SMSTSOrgName")
	Else
		document.title = "Microsoft Deployment Toolkit"
	End if


	' Schedule a progress update in a half second

	window.SetTimeout GetRef("DisplayProgress"),500

End function


Function DisplayProgress

	Dim sPercent, sMessage


	' Retrieve and display the updated progress

	On Error Resume Next
	sPercent = oShell.RegRead("HKLM\Software\Microsoft\Deployment 4\ProgressPercent")
	sMessage = oShell.RegRead("HKLM\Software\Microsoft\Deployment 4\ProgressText")
	ProgressPercent.InnerHTML = sPercent
	ProgressMessage.InnerHTML = sMessage
	MyProgress.style.width = CStr(sPercent) & "%"
	On Error Goto 0


	' If at 100%, click next.  Otherwise, schedule us again.

	If sPercent = 100 then
		window.SetTimeout GetRef("ProgressNext"),700
	Else
		window.SetTimeout GetRef("DisplayProgress"),500
	End if

End function


Function ProgressNext

	' Click next

	ButtonNextClick

End Function
