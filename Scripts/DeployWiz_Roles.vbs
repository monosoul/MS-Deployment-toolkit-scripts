' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_Roles.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Allow the selection of roles and features.
' // 
' // ***************************************************************************

Option Explicit

Dim oRoles

Function InitializeOSRoles()

	Dim sPreviousOSRoles
	Dim sPreviousOSRoleServices
	Dim sPreviousOSFeatures
	Dim iSvrMgrVal
	Dim oSvrMgrXml
	Dim sFoundXmlFile
	Dim sServer
	Dim sServerCore
	Dim sVersion
	Dim oStyle
	Dim oResult
	Dim sCurrent
	Dim oNode


	' Get the current list of selected roles and features

	If OptionalOSRoles.value <> "" Then
		sPreviousOSRoles = OptionalOSRoles.Value
	ElseIf oEnvironment.Item("OptionalOSRoles") <> "" Then
		sPreviousOSRoles = oEnvironment.Item("OptionalOSRoles")
	Else
		sPreviousOSRoles = ""
	End If

	If OptionalOSRoleServices.value <> "" Then
		sPreviousOSRoleServices = OptionalOSRoleServices.Value
	ElseIf oEnvironment.Item("OptionalOSRoleServices") <> "" Then
		sPreviousOSRoleServices = oEnvironment.Item("OptionalOSRoleServices")
	Else
		sPreviousOSRoleServices = ""
	End If

	If OptionalOSFeatures.value <> "" Then
		sPreviousOSFeatures = OptionalOSFeatures.Value
	ElseIf oEnvironment.Item("OptionalOSFeatures") <> "" Then
		sPreviousOSFeatures = oEnvironment.Item("OptionalOSFeatures")
	Else
		sPreviousOSFeatures = ""
	End If


	' Load the XML file

	iSvrMgrVal = oUtility.FindFile( "ServerManager.xml" , sFoundXmlFile)
	TestAndFail iSvrMgrVal, 9000, "Did not find ServerManager.xml"

	On Error Resume Next
	Err.Clear
	Set oSvrMgrXml = oUtility.CreateXMLDOMObjectEx(sFoundXmlFile)
	If Err then
		' Unable to create XML object
		Err.Clear
	End if
	On Error Goto 0


	' Get the appropriate OS's role list

	If oEnvironment.Item("OSVersion") <> "WinPE" and (Property("DeploymentType") = "CUSTOM" or Property("DeploymentType") = "StateRestore") then

		If UCase(oEnvironment.Item("IsServerOS")) = "TRUE" then
			sServer = "yes"
		Else
			sServer = "no"
		End if

		If UCase(oEnvironment.Item("IsServerCoreOS")) = "TRUE" then
			sServerCore = "yes"
		Else
			sServerCore = "no"
		End if

		sVersion = Left(oEnvironment.Item("OSCurrentVersion"),3)

	Else

		sServerCore = "no"
		If Instr(UCase(oEnvironment.Item("ImageFlags")), "SERVER") > 0 then
			sServer = "yes"
			If Instr(UCase(oEnvironment.Item("ImageFlags")), "CORE") > 0 then
				sServerCore = "yes"
			End if
		Else
			sServer = "no"
		End if

		sVersion = Left(oEnvironment.Item("ImageBuild"),3)


		' Because XP/2003 don't have image flags, check the version and manually force server

		If sVersion = "5.2" then
			sServer = "yes"
		End if

	End if


	Set oRoles = oSvrMgrXml.SelectSingleNode("//Roles[@Server='" & sServer & "' and @Core='" & sServerCore & "' and @OS='" & sVersion & "']")
	If oRoles is Nothing then
		Exit Function
	End if	


	' Load the stylesheet

	Set oStyle = oUtility.CreateXMLDOMObjectEx("DeployWiz_Roles.xsl")


	' Format the tree

	Set oResult = oUtility.CreateXMLDOMObject
	oRoles.transformNodeToObject oStyle, oResult	
	RoleListDiv.InnerHTML = oResult.Xml


	' Check the appropriate items

	If sPreviousOSRoles <> "" then
		For each sCurrent in Split(sPreviousOSRoles, ",")
			Set oNode = document.GetElementByID("Role." & sCurrent)
			If not (oNode is Nothing) then
				oNode.checked = "true"
			End if
		Next
	End if

	If sPreviousOSRoleServices <> "" Then
		For each sCurrent in Split(sPreviousOSRoleServices, ",")
			Set oNode = document.GetElementByID("RoleService." & sCurrent)
			If not (oNode is Nothing) then
				oNode.checked = "true"
			End if
		Next
	End If

	If sPreviousOSFeatures <> "" Then
		For each sCurrent in Split(sPreviousOSFeatures, ",")
			Set oNode = document.GetElementByID("Feature." & sCurrent)
			If not (oNode is Nothing) then
				oNode.checked = "true"
			End if
		Next
	End If

End Function


Function CheckRoles
	
	Dim oElement

	For each oElement in RoleListDiv.all.tags("input")
		If UCase(oElement.type) = "CHECKBOX" then
			If not oElement.disabled then
				oElement.checked = true
			End if
		End if
	Next

End Function


Function UncheckRoles
	
	Dim oElement

	For each oElement in RoleListDiv.all.tags("input")
		If UCase(oElement.type) = "CHECKBOX" then
			If not oElement.disabled then
				oElement.checked = false
			End if
		End if
	Next

End Function


Function ValidateOSRoles()

	Dim sRoles
	Dim sRoleServices
	Dim sFeatures
	Dim sName
	Dim oNode
	Dim oId


	' Find all the selected roles

	sRoles = ""
	sRoleServices = ""
	sFeatures = ""

	For each sName in Array("Role", "RoleService", "Feature")
		For each oNode in oRoles.selectNodes(".//" & sName)
			Set oId = oNode.Attributes.getNamedItem("Id")
			If not (oId is Nothing) then
				Set oNode = document.GetElementByID(sName & "." & oId.Value)
				If not (oNode is Nothing) then
					If oNode.checked then
						Select Case sName
						Case "Role"
							sRoles = sRoles & oId.Value & ","
						Case "RoleService"
							sRoleServices = sRoleServices & oId.Value & ","
						Case "Feature"
							sFeatures = sFeatures & oId.Value & ","
						End Select
					End if
				End if
			End if
		Next
	Next


	' Save the selected roles to the hidden fields on the form (will be written automatically to the environment)

	If sRoles <> "" then
		OptionalOSRoles.value = Left(sRoles, Len(sRoles) - 1)
	Else
		OptionalOSRoles.value = ""
	End if
	If sRoleServices <> "" then
		OptionalOSRoleServices.value = Left(sRoleServices, Len(sRoleServices) - 1)
	Else
		OptionalOSRoleServices.value = ""
	End if
	If sFeatures <> "" then
		OptionalOSFeatures.value = Left(sFeatures, Len(sFeatures) - 1)
	Else
		OptionalOSFeatures.value = ""
	End if

	ValidateOSRoles = true
 
End Function
