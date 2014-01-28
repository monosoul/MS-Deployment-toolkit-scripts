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


'''''''''''''''''''''''''''''''''''''
'  Image List
'

Dim g_AllOperatingSystems

Function AllOperatingSystems


	Dim oOSes

	If isempty(g_AllOperatingSystems) then
	
		set oOSes = new ConfigFile
		oOSes.sFileType = "OperatingSystems"
		oOSes.bMustSucceed = false
		
		set g_AllOperatingSystems = oOSes.FindAllItems
		
	End if

	set AllOperatingSystems = g_AllOperatingSystems

End function


Function InitializeTSList
	Dim oItem, sXPathOld
	
	If oEnvironment.Item("TaskSequenceID") <> "" and oProperties("TSGuid") = "" then
		
		sXPathOld = oTaskSequences.xPathFilter
		for each oItem in oTaskSequences.oControlFile.SelectNodes( "/*/*[ID = '" & oEnvironment.Item("TaskSequenceID")&"']")
			oLogging.CreateEntry "TSGuid changed via TaskSequenceID = " & oEnvironment.Item("TaskSequenceID"), LogTypeInfo
			oEnvironment.Item("TSGuid") = oItem.Attributes.getNamedItem("guid").value
			exit for
		next
		
		oTaskSequences.xPathFilter = sXPathOld 
		
	End if

	TSListBox.InnerHTML = oTaskSequences.GetHTMLEx ( "Radio", "TSGuid" )
	
	PopulateElements
	TSItemChange

End function


Function TSItemChange

	Dim oInput
	ButtonNext.Disabled = TRUE
	
	for each oInput in document.getElementsByName("TSGuid")
		If oInput.Checked then
			oLogging.CreateEntry "Found CHecked Item: " & oInput.Value, LogTypeVerbose
		
			ButtonNext.Disabled = FALSE
			exit function
		End if
	next

End function


'''''''''''''''''''''''''''''''''''''
'  Validate task sequence List
'

Function ValidateTSList

	Dim oTaskList
	Dim oTS
	Dim oItem
	Dim oOSItem
	Dim sID
	Dim bFound
	Dim sTemplate
	
	set oTS = new ConfigFile
	oTS.sFileType = "TaskSequences"

	SaveAllDataElements

	If Property("TSGuid") = "" then
		oLogging.CreateEntry "No valid TSGuid found in the environment.", LogTypeWarning
		ValidateTSList = false
	End if

	oLogging.CreateEntry "TSGuid Found: " & Property("TSGuid"), LogTypeVerbose
	
	sID = ""
	sTemplate = ""
	If oTS.FindAllItems.Exists(Property("TSGuid")) then
		sID = oUtility.SelectSingleNodeString(oTS.FindAllItems.Item(Property("TSGuid")),"./ID")
		sTemplate = oUtility.SelectSingleNodeString(oTS.FindAllItems.Item(Property("TSGuid")),"./TaskSequenceTemplate")
	End if
	
	oEnvironment.item("TaskSequenceID") = sID
	TestAndLog sID <> "", "Verify Task Sequence ID: " & sID
	Set oTaskList = oUtility.LoadConfigFileSafe( sID & "\TS.XML" )

	If not FindTaskSequenceStep( "//step[@type='BDD_InstallOS']", "" ) then

		oLogging.CreateEntry "Task Sequence does not contain an OS and does not contain a LTIApply.wsf step, possibly a Custom Step or a Client Replace.", LogTypeInfo
		
		oProperties.Item("OSGUID")=""
		If not (oTaskList.SelectSingleNode("//group[@name='State Restore']") is nothing) then
			oProperties("DeploymentType") = "StateRestore"
		ElseIf sTemplate <> "ClientReplace.xml" and oTaskList.SelectSingleNode("//step[@name='Capture User State']") is nothing then
			oProperties("DeploymentType")="CUSTOM"
		Else
			oProperties("DeploymentType")="REPLACE"

			RMPropIfFound("ImageIndex")
			RMPropIfFound("ImageSize")
			RMPRopIfFound("ImageFlags")
			RMPropIfFound("ImageBuild")
			RMPropIfFound("InstallFromPath")
			RMPropIfFound("ImageMemory")

			oEnvironment.Item("ImageProcessor")=Ucase(oEnvironment.Item("Architecture"))
		End if

	Elseif oEnvironment.Item("OSVERSION")="WinPE" Then

		oProperties("DeploymentType")="NEWCOMPUTER"

	Else

		oLogging.CreateEntry "Task Sequence contains a LTIApply.wsf step, and is not running within WinPE.", LogTypeInfo
		oProperties("DeploymentType") = "REFRESH"
		oEnvironment.Item("DeployTemplate")=Ucase(Left(sTemplate,Instr(sTemplate,".")-1))

	End if

	oLogging.CreateEntry "DeploymentType = " & oProperties("DeploymentType"), LogTypeInfo

	
	set oTaskList = nothing
	set oTS = nothing


	' Set the related properties

	oEnvironment.Item("ImageProcessor") = ""
	oEnvironment.Item("OSGUID")=""
	oUtility.SetTaskSequenceProperties sID


	If Left(Property("ImageBuild"), 1) < "6" then
		RMPropIfFound("LanguagePacks")
		RMPropIfFound("UserLocaleAndLang")
		RMPropIfFound("KeyboardLocale")
		RMPropIfFound("UserLocale")
		RMPropIfFound("UILanguage")
		RMPropIfFound("BdePin")
		RMPropIfFound("BdeModeSelect1")
		RMPropIfFound("BdeModeSelect2")
		RMPropIfFound("OSDBitLockerStartupKeyDrive")
		RMPropIfFound("WaitForEncryption")
		RMPropIfFound("BdeInstall")
		RMPropIfFound("OSDBitLockerWaitForEncryption")
		RMPropIfFound("BdeRecoveryKey")
		RMPropIfFound("BdeInstallSuppress")
	End If

	If oEnvironment.Item("OSGUID") <> "" and oEnvironment.Item("ImageProcessor") = "" then
		' There was an OSGUID defined within the TS.xml file, however the GUID was not found 
		' within the OperatingSystems.xml file. Which is a dependency error. Block the wizard.
		ValidateTSList = False
		ButtonNext.Disabled = True
		Bad_OSGUID.style.display = "inline"
	Else
		ValidateTSList = True
		ButtonNext.Disabled = False
		Bad_OSGUID.style.display = "none"
	ENd if


End Function
