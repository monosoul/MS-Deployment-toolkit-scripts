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

Dim g_oTaskSequences


Function oTaskSequences

	If isempty(g_oTaskSequences) then
	
		oLogging.CreateEntry "Begin InitializeTSList...", LogTypeVerbose

		set g_oTaskSequences = new ConfigFile
		g_oTaskSequences.sFileType = "TaskSequences"
		g_oTaskSequences.sSelectionProfile = oEnvironment.Item("WizardSelectionProfile")
		g_oTaskSequences.sCustomSelectionProfile = oEnvironment.Item("CustomWizardSelectionProfile")
		g_oTaskSequences.sHTMLPropertyHook = " onPropertyChange='TSItemChange'"
		set g_oTaskSequences.fnCustomFilter = GetRef("CustomTSFilter")
		
		oLogging.CreateEntry "Finished InitializeTSList...", LogTypeVerbose
		
	End if
	set oTaskSequences = g_oTaskSequences
	
End function


function CustomTSFilter( sGuid, oItem )

	' Hook for ZTIConfigFile.vbs. Return True only if the Item should be displayed, otherwise false.
	Dim oTaskList
	Dim oTaskOsGuid	
	Dim oOS
	DIm sOSPlatform
	Dim sOSBuild
	
	Set oTaskList = oUtility.LoadConfigFileSafe( "Control\" & oUtility.SelectSingleNodeString(oItem,"ID") & "\TS.xml")
	Set oTaskOsGuid = oTaskList.SelectSingleNode("//globalVarList/variable[@name='OSGUID']")
	
	CustomTSFilter = True

	If oTaskOsGuid is Nothing then

		' This Task Sequence does not have any associated OS, allways include

	ElseIf not AllOperatingSystems.Exists(oTaskOsGuid.text) then

		' This Task Sequence does not have any associated OS, allways include
		oLogging.CreateEntry "ERROR: Invalid OS GUID " & oTaskOsGuid.text & " specified for task sequence " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeInfo

	Else
	
		set oOS = AllOperatingSystems.Item(oTaskOsGuid.text)
		
		If not oOS.selectSingleNode("SMSImage") is nothing then
			If ucase(oUtility.SelectSingleNodeString(oOS,"SMSImage")) = "TRUE" then
				oLogging.CreateEntry "Skip SMS OS " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeVerbose
				CustomTSFilter = False
				exit function
			End if
		End if
		
		if not oOS.selectSingleNode("Platform") is nothing then
		
			sOSPlatform = oUtility.SelectSingleNodeString(oOS,"Platform")
			sOSBuild = oUtility.SelectSingleNodeString(oOS, "Build")

			If UCase(sOSPlatform) = UCase(oEnvironment.Item("Architecture")) then

				' Same Archtecture as current OS, No problems.

			ElseIf Instr(1, Property("CapableArchitecture"), sOSPlatform, vbTextCompare) = 0 then

				oLogging.CreateEntry "Not Capable of running Platform: " & sOSPlatform & "   " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeInfo
				CustomTSFilter = False

			ElseIf oEnv("SystemDrive") <> "X:"  then

				' We are not in WinPE, so we can still apply any OS

			ElseIf ucase(oEnvironment.Item("ForceApplyFallback")) = "NEVER" and ucase(oUtility.SelectSingleNodeString(oOS, "IncludesSetup")) = "TRUE" then

				oLogging.CreateEntry "Skip cross platform unattended install disabled (ForceApplyFallback = NEVER). " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeInfo
				CustomTSFilter = False

			ElseIf UCase(sOSPlatform) = "X86" and UCase(oEnvironment.Item("Architecture")) = "X64" then

				oLogging.CreateEntry "Skip cross platform x86 install from x64 Windows PE. " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeInfo
				CustomTSFilter = False			

			ElseIf Left(sOSBuild, 3) < "6.1" and ucase(oUtility.SelectSingleNodeString(oOS, "IncludesSetup")) = "TRUE" then

				oLogging.CreateEntry "Skip cross platform unattended install for OS'es earlier than Windows 7. " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeInfo
				CustomTSFilter = False
			
			End if

		End if
		
	End if 

	If not oItem.selectSingleNode("SupportedPlatform") is nothing then
		If not oUtility.IsSupportedPlatform(oUtility.SelectSingleNodeString(oItem,"SupportedPlatform")) then
			oLogging.CreateEntry "Skip unsupported platform " & oUtility.SelectSingleNodeString(oItem,"SupportedPlatform") & " in " & oUtility.SelectSingleNodeString(oItem,"ID"), LogTypeVerbose
			CustomTSFilter = False
			Exit function
		End if
	End if
	
End function


Dim sCachedTSID
Dim oCachedTaskList

Function FindTaskSequenceStep(sStepType, sScriptCmd )
	Dim oAction
	Dim oItem
	Dim oOptionDiableVal


	' Is there a task sequence chosen yet?  If not, the step can't possibly be present

	If Property("TaskSequenceID") = "" then
		oLogging.CreateEntry "No task sequence has been selected yet.", LogTypeVerbose
		FindTaskSequenceStep = false
		Exit Function
	End if


	' For efficiency, only load the task sequence if it has changed from the last time we loaded it

	If sCachedTSID <> Property("TaskSequenceID") then
		Set oCachedTaskList = oUtility.LoadConfigFileSafe( Property("TaskSequenceID") & "\TS.XML" )
		sCachedTSID = Property("TaskSequenceID")
	End if


	' Get the list of nodes of the specified type

	set oItem = oCachedTaskList.SelectNodes(sStepType)
	

	If not oItem is nothing then
		oLogging.CreateEntry "Found Task Sequence Item: " & sStepType, LogTypeInfo
		
	ElseIf len(sScriptCmd) > 0 then

		oLogging.CreateEntry "Unable to find Task Sequence step of type " & sStepType & ", performing more exhaustive search...", LogTypeInfo
		For each oAction in oCachedTaskList.SelectNodes("//action")
			If instr(1,oAction.XML,sScriptCmd,vbTExtCompare) <> 0 then
				oLogging.CreateEntry "Found Task Sequence Item: " & sScriptCmd, LogTypeInfo
				set oItem = oAction
				exit for
			End if
		Next

	End if
	

	' Verify this step is not "disabled"...

	If oItem is nothing then 
		oLogging.CreateEntry "Unable to find Task Sequence step of type " & sStepType , LogTypeInfo
		FindTaskSequenceStep = False
	Else
		'Loop through each step in the collection until first enabled step and exit the loop
		For each oOptionDiableVal in oItem
			set oAction = oOptionDiableVal.Attributes.getNamedItem("disable")
			If  oAction is nothing then
				FindTaskSequenceStep = true
				
			Else
				FindTaskSequenceStep = lcase(oAction.Value) <> "true"
				
				If FindTaskSequenceStep = true then
					Exit For
				End If
			End if
		Next

		oLogging.CreateEntry "Found Task Sequence step of type " & sStepType & " = " & FindTaskSequenceStep, LogTypeInfo

	End if

End function
