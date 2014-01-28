' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_Applications.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Deployment Wizard application initialization routines
' // 
' // ***************************************************************************


Option Explicit



'''''''''''''''''''''''''''''''''''''''''''
'  Application List
'

Dim g_sApplicationDialog
Dim g_oXMLAppList

Function IsThereAtLeastOneApplicationPresent

	Dim oXMLAppList
	Dim dXMLCollection
	Dim oTaskList
	Dim oAction

	' Bail out early if there is no "Install Application" step in the task sequence

	If not FindTaskSequenceStep( "//step[@type='BDD_InstallApplication' and ./defaultVarList/variable[@name='ApplicationGUID'] and ./defaultVarList[variable='']]", "ZTIApplications.wsf" ) then
		IsThereAtLeastOneApplicationPresent = false
		Exit function
	End if


	' Load and cache the application list

	If IsEmpty(g_oXMLAppList) then
		Set g_oXMLAppList = new ConfigFile
		g_oXMLAppList.sFileType = "Applications"
	End if


	' Get the filtered list of applications

	g_oXMLAppList.sSelectionProfile = oEnvironment.Item("WizardSelectionProfile")
	g_oXMLAppList.sCustomSelectionProfile = oEnvironment.Item("CustomWizardSelectionProfile")
	Set dXMLCollection = g_oXMLAppList.FindItems

	If dXMLCollection.count = 0 then
		IsThereAtLeastOneApplicationPresent = False
		g_sApplicationDialog = ""
		Exit Function
	End if

	g_sApplicationDialog = g_oXMLAppList.GetHTMLEx( "CheckBox", "Applications" ) 

	IsThereAtLeastOneApplicationPresent = True
	
End function

Function InitializeApplicationList

	AppListBox.InnerHTML = g_sApplicationDialog
	PopulateElements

End Function


Function ReadyInitializeApplicationList
	Dim oInput, oApplicationList, oAppItem

	If not ImageList.readystate = "complete" then
		Exit function
	End if

	Set oApplicationList = document.getElementsByName("Applications")

	If oApplicationList is nothing then
		Exit function
	ElseIf oApplicationList.Length < 1 then
		Exit function
	End if

	For each oInput in oApplicationList
		If UCase(document.all.item(oInput.SourceIndex - 1).TagName) = "INPUT" then
			If oInput.Value = "" then
				document.all.item(oInput.SourceIndex - 1).Disabled = TRUE
				document.all.item(oInput.SourceIndex - 1).Style.Display = "none"
			Else
				document.all.item(oInput.SourceIndex - 1).Style.Display = "inline"
				If not IsEmpty(Property("Applications"))then
					For each oAppItem in Property("Applications")
						If UCase(oAppItem) = UCase(oInput.Value) then
							document.all.item(oInput.SourceIndex - 1).checked = TRUE
							Exit for
						End if
					Next
				End if
				If not IsEmpty(Property("MandatoryApplications"))then
					For each oAppItem in Property("MandatoryApplications")
						If UCase(oAppItem) = UCase(oInput.Value) then
							document.all.item(oInput.SourceIndex - 1).disabled = TRUE
							document.all.item(oInput.SourceIndex - 1).checked = TRUE
							Exit for
						End if
					Next
				End if

			End if
		End if

	Next

End function


Sub AppItemChange

	document.all.item(window.event.srcElement.SourceIndex + 1).Disabled = not window.event.SrcElement.checked

End sub

