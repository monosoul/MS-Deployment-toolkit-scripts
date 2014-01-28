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

'
' Will return a dictionary object containing all Friendly Names Given a GUID as string, this funtion will search all *.xml files in the deployroot for a match.
'
Function GetFriendlyNamesofGUIDs

	Dim oFiles
	Dim oFolder
	Dim oXMLFile
	Dim oXMLNode
	Dim sName
	Dim GuidList

	Set GuidList = CreateObject("Scripting.Dictionary")
	GuidList.CompareMode = vbTextCompare

	Set oFolder = oFSO.GetFolder( oEnvironment.Item("DeployRoot") & "\control" )
	If oFolder is nothing then
		oLogging.CreateEntry oUtility.ScriptName & " Unable to find DeployRoot!", LogTypeError
		Exit function
	End if

	For each oFiles in oFolder.Files

		If UCase(right(oFIles.Name, 4)) = ".XML" then
			Set oXMLFile = oUtility.CreateXMLDOMObjectEx( oFiles.Path )
			If not oXMLFile is nothing then

				for each oXMLNode in oXMLFile.selectNodes("//*/*[@guid]")

					if not oXMLNode.selectSingleNode("./Name") is nothing then
						sName = oUtility.SelectSingleNodeString(oXMLNode,"./Name")

						if not oXMLNode.selectSingleNode("./Language") is nothing then
							if oUtility.SelectSingleNodeString(oXMLNode,"./Language") <> "" then
								sName = sName & " ( " & oUtility.SelectSingleNodeString(oXMLNode,"./Language") & " )"
							end if
						end if

						if not oXMLNode.Attributes.getNamedItem("guid") is nothing then
							if oXMLNode.Attributes.getNamedItem("guid").value <> "" and sName <> "" then
								if not GuidList.Exists(oXMLNode.Attributes.getNamedItem("guid").value) then
									GuidList.Add oXMLNode.Attributes.getNamedItem("guid").value, sName
								end if
							end if
						end if
					end if

				next

			End if
		End if

	Next

	set GetFriendlyNamesofGUIDs = GuidList
End function


Function PrepareFinalScreen

	Dim GuidList
	Dim p, i, item, Buffer

	set GuidList = GetFriendlyNamesofGUIDs

	Dim re, Match

	For each p in oProperties.Keys

		If IsObject(oProperties(p)) or IsArray(oProperties(p)) then
			i = 1
			For each item in oProperties(p)
				If Item <> "" then
					oStrings.AddToList Buffer, p & i &  " = """ & item & """", vbNewLine
					i = i + 1
				End if
					
			next
		ElseIf ucase(p) = "DEFAULTDESTINATIONDISK" then
			' Skip...
		ElseIf ucase(p) = "DEFAULTDESTINATIONPARTITION" then
			' Skip...
		ElseIf ucase(p) = "DEFAULTDESTINATIONISDIRTY" then
			' Skip...
		ElseIf ucase(p) = "KEYBOARDLOCALE_EDIT" then
			' Skip...
		ElseIf ucase(p) = "USERLOCALE_EDIT" then
			' Skip...
		ElseIf oProperties(p) = "" then
			' Skip...
		ElseIf Instr(1, p, "Password" , vbTextCompare ) <> 0 then
			oStrings.AddToList Buffer, p & " = ""***********""", vbNewLine
		else
			oStrings.AddToList Buffer, p & " = """ & oProperties(p) & """", vbNewLine
		end if
	Next

	'
	' Given a text string containing GUID ID's of configuration entries on the deployment share
	'   This function will search/replace all GUID's within the text blob.
	'
	Set re = new regexp
	re.IgnoreCase = True
	re.Global = True
	re.Pattern = "\{[A-F0-9]{8}\-[A-F0-9]{4}\-[A-F0-9]{4}\-[A-F0-9]{4}\-[A-F0-9]{12}\}"

	On error resume next
	Do while re.Test( Buffer )
		For each Match in re.execute(Buffer)
			Buffer = mid(Buffer,1,Match.FirstIndex) & _
				GuidList.Item(Match.Value) & _
				mid(Buffer,Match.FirstIndex+match.Length+1)
			Exit for
		Next
	Loop
	On error goto 0

	optionalWindow1.InnerText = Buffer

End function
