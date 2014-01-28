' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_DeployRoot.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Main Client Deployment Wizard Initialization routines
' // 
' // ***************************************************************************


Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
'  DeployRoot!
'

Dim oXMLSiteData

Function InitializeDeployRoot

	Dim oXMLDefault
	Dim oItem
	Dim oOption
	Dim iRetVal
	Dim sLocationServer

	Set oXMLSiteData = nothing
	document.body.onkeyDown = GetRef("KeyHandler")

	' Save the local DeployRoot location
	If property("LocalDeployRoot") = "" then
		oProperties("LocalDeployRoot") = property("DeployRoot")
	End if

	' Find the LocationServer.xml file if it exists.  If it doesn't, then exit.
	iRetVal = oUtility.FindFile("LocationServer.xml", sLocationServer)
	If iRetVal <> SUCCESS then

		' Force manual entry
		RadioCT2.checked = TRUE
		SiteList.disabled = TRUE
		DisplayLocal.style.display = "none"
		DisplayValidateDeployRoot

		oLogging.CreateEntry "No LocationServer.xml file was found, so no additonal DeployRoot pane initialization is required.", LogTypeInfo
		Exit Function
	End if
	
	' Load the Site Configuration XML file.
	Set oXMLSiteData = oUtility.CreateXMLDOMObjectEx( sLocationServer )
	If oXMLSiteData is nothing or oXMLSiteData.ParseError.ErrorCode <> 0 then

		' Force manual entry
		RadioCT2.checked = TRUE
		SiteList.disabled = TRUE
		DisplayLocal.style.display = "none"
		DisplayValidateDeployRoot

		oLogging.CreateEntry "The LocationServer.xml file was found at " & sLocationServer & " but it could not be loaded, probably because it was invalid.", LogTypeWarning
		Exit Function
	End if

	If not ( oXMLSiteData.selectNodes("//servers/server") is nothing ) then
		While SiteList.options.length > 0
			SiteList.remove 0
		Wend
	End if

	For each oItem in oXMLSiteData.selectNodes("//servers/server")

		Set oOption = document.createElement("OPTION")
		oOption.Value = oUtility.SelectSingleNodeString(oItem,"serverid")
		oOption.Text = oUtility.SelectSingleNodeString(oItem,"friendlyname")
		SiteList.Add oOption

	Next

	' Now attempt to get a default from a server!
	If oUtility.SelectSingleNodeString(oXMLSiteData,"//servers/QueryDefault") <> "" then

		Set oXMLDefault = oUtility.CreateXMLDOMObjectEx( oUtility.SelectSingleNodeString(oXMLSiteData,"//servers/QueryDefault") )
		If not (oXMLDefault is nothing) then
			For each oItem in oXMLDefault.selectNodes("//DefaultSites/DefaultSite")
				SiteList.Value = oItem.Text
				If SiteList.Value = oItem.Text then
					Exit for
				End if
			Next
		End if
		Set oXMLDefault = nothing

	End if

	DisplayValidateDeployRoot

End Function



'''''''''''''''''''''''''''''''''''''
'  Validate DeployRoot and Credentials
'


Function ChangeServerFromSite

	Dim oItem
	dim UpperBound
	dim oServer
	dim Index

	dim oServerList

	if oXMLSiteData is nothing then
		exit function
	end if

	for each oItem in oXMLSiteData.selectNodes("//servers/server")
		if SiteList.value = oUtility.SelectSingleNodeString(oItem,"serverid") then

			set oServerList = oItem.selectNodes("(server1|server2|server3|server4|server5|server6|server7|server8|server9|server|UNCPath)")

			' Get the Weighted Value UpperBound
			UpperBound = 0
			for each oServer in oServerList
				if oServer.Attributes.getQualifiedItem("weight","") is nothing then
					UpperBound = UpperBound + 1
				else
					UpperBound = UpperBound + cint(oServer.Attributes.getQualifiedItem("weight","").Value)
				end if
			next

			randomize
			Index = int(rnd * UpperBound + 1)

			' Pick a random server entry based on Weighted Value.
			UpperBound = 0
			for each oServer in oServerList
				if oServer.Attributes.getQualifiedItem("weight","") is nothing then
					UpperBound = UpperBound + 1
				else
					UpperBound = UpperBound + cint(oServer.Attributes.getQualifiedItem("weight","").Value)
				end if

				if Index <= UpperBound then
					DeployRoot.value = oServer.Text
					DisplayValidateDeployRoot
					exit function
				end if
			next

		end if
	next

	DisplayValidateDeployRoot

end function


Function DisplayValidateDeployRoot

	DeployRoot.readonly = RadioCT1.checked
	if RadioCT1.checked then
		DeployRoot.style.color = "graytext"
	else
		DeployRoot.style.color = ""
	end if

	SiteList.Disabled = RadioCT2.Checked

	DisplayValidateDeployRoot = ParseAllWarningLabels

end function


Function ValidateDeployRoot
	Dim oItem
	Dim oVariable
	Dim oName
	Dim sCmd

	ValidateDeployRoot = DisplayValidateDeployRoot

	If ValidateDeployRoot = FALSE then
		Exit function
	End if


	' Test the share for network access.

	ValidateDeployRoot = FALSE

	Do
		On Error Resume Next
		Err.Clear
		If oFSO.FileExists(DeployRoot.value & "\Control\OperatingSystems.xml" ) then
			ValidateDeployRoot = TRUE
			Exit Do
		End if
		On Error Goto 0

		If Mid(DeployRoot.value, 2, 2) = ":\" then
			Alert "Invalid or unrecognized path specified!"  ' For example, if they specified W:\Deploy and that didn't exist
			ValidateDeployRoot = FALSE
			Exit Function
		ElseIf not ValidateDeployRoot then

			' Get the credentials and connect to the share!

			oEnvironment.Item("UserID") = ""
			oEnvironment.Item("UserDomain") = ""
			oEnvironment.Item("UserPassword") = ""

			oShell.Run "mshta.exe " & window.document.location.href & " /NotWizard /LeaveShareOpen /ValidateAgainstUNCPath:" & DeployRoot.value & " /Definition:Credentials_ENU.xml", 1, true

			If UCase(oEnvironment.Item("UserCredentials")) <> "TRUE" then
				Alert "Could not validate Credentials!"
				Exit function
			End if

		End if

	Loop until ValidateDeployRoot = TRUE


	' Flush the value to variables.dat, before we continue.

	SaveAllDataElements
	SaveProperties

	' Process full rules

	sCmd = "wscript.exe """ & oUtility.ScriptDir & "\ZTIGather.wsf"""
	oItem = oSHell.Run(sCmd, , true)

	' Extract out other fields within the XML Data Object.

	If oXMLSiteData is nothing then
		Exit function
	End if

	For each oItem in oXMLSiteData.selectNodes("//servers/server")
		If SiteList.value = oUtility.SelectSingleNodeString(oItem,"serverid") then
			For each oVariable in oItem.selectNodes("otherparameters/parameter")
				Set oName = oVariable.Attributes.getQualifiedItem("name","")
				If not oName is Nothing then
					oProperties(oName.Value) = oVariable.Text
				End if
			Next

		End if
	Next

End Function

