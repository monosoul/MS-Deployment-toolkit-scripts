' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_LanguagePack.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Wizard pane for choosing one or more language packs for install
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


Function MarkupName(LocaleName)

	Dim oLang, width

	width = 99
	For each oLang in oXMLLanguageList.selectNodes("//LOCALEDATA/LOCALE/SISO639LANGNAME")
		If Instr(1,LocaleName,oLang.Text & "-", vbTextCompare) = 1 and len(oLang.Text) < width then

			width = len(oLang.Text)
			MarkupName = unescape(replace(oUtility.SelectSingleNodeString(oLang.ParentNode,"SENGLANGUAGE"),"\x","%u")) & " (" & _
				LocaleName & ")"
			
		End if
	Next

End Function 


Function ConstructLPQuery ( isLangPack )
	Dim Keyword
	Dim isServer
	Dim ImgBuild
	Dim SPVersion
	Dim LPQuery
	Dim LPVersion
	Dim i

	isServer  = inStr(1,oEnvironment.Item("ImageFlags"),"SERVER",vbTextCompare) <> 0
	ImgBuild  = oEnvironment.Item("ImageBuild")

	If isServer then
		LPQuery = "PackageType = 'LanguagePack' and ( ProductName = 'Microsoft-Windows-Server-LanguagePack-Package' or ProductName = 'Microsoft-Windows-Server-Refresh-LanguagePack-Package' )"
	Else
		LPQuery = "PackageType = 'LanguagePack' and ( ProductName = 'Microsoft-Windows-Client-LanguagePack-Package' or ProductName = 'Microsoft-Windows-Client-Refresh-LanguagePack-Package' )"
	End if

	If not isLangPack then
		LPQuery = "PackageType != 'LanguagePack' and (substring(ProductVersion,1,8) = '" & left(ImgBuild,8) & "' or ProductVersion = '') "
	ElseIf isServer and left(ImgBuild,4) = "6.0." then
		' All Windows Server 2008 Language Packs use Product Version 6.0.6001.18000
		LPQuery = LPQuery & " and substring(ProductVersion,1,8) = '6.0.6001' "
	ElseIf left(ImgBuild,4) = "6.0." then
		' All Windows Vista Language Packs use Product Version 6.0.6000.16386
		LPQuery = LPQuery & " and  substring(ProductVersion,1,8) = '6.0.6000' "
	Else
		LPQuery = LPQuery & " and substring(ProductVersion,1,7) = '" & left(ImgBuild,7) & "' and substring(ProductVersion,5,4) >= '" & mid(ImgBuild,5,4) & "'"
	End if

	If not isLangPack then
		' Nothing
	ElseIf left(ImgBuild,4) = "6.0." then
		LPVersion = Mid(ImgBuild,8,1)
		If IsNumeric(LPVersion) and LPVersion > 0 then
			' Exclude all Language Packs that are less than the Current OS.
			LPQuery = LPQuery & " and Keyword != 'Language Pack'"
			For i = 2 to LPVersion
				LPQuery = LPQuery & " and Keyword != 'SP" & (LPVersion - 1) & " Language Pack'"
			Next
		End if
	End if

	If UCase(oEnvironment.Item("ImageProcessor")) = "X64" then
		LPQuery = "//packages/package[ProcessorArchitecture = 'amd64' and " & LPQuery & "]"
	Else
		LPQuery = "//packages/package[ProcessorArchitecture = 'x86' and " & LPQuery & "]"
	End if

	oLogging.CreateEntry vbTab & "QUERY: " & LPQuery, LogTypeInfo
	ConstructLPQuery = LPQuery

End function

Dim g_oXMLPackageList
Dim g_sPackageDialogBox

Function CanDisplayPackageDialogBox

	Dim dXMLCollection
	Dim LocalLanguage
	Dim oItem
	Dim sInputType
	Dim sDone
	Dim oNewItem
	Dim sToAdd

	' Load and cache the application list

	If IsEmpty(g_oXMLPackageList) then

		Set g_oXMLPackageList = new ConfigFile
		g_oXMLPackageList.sFileType = "Packages"

	End if


	' Get the list of language packs

	g_oXMLPackageList.sSelectionProfile = oEnvironment.Item("WizardSelectionProfile")
	g_oXMLPackageList.sCustomSelectionProfile = oEnvironment.Item("CustomWizardSelectionProfile")
	Set g_oXMLPackageList.fnCustomFilter = GetRef("CustomPackageFilter")	
	Set dXMLCollection = g_oXMLPackageList.FindItemsEx(ConstructLPQuery(TRUE))

	If dXMLCollection.count = 0 then
		CanDisplayPackageDialogBox = False
		Exit Function
	End if

	CanDisplayPackageDialogBox = TRUE	
	oLogging.CreateEntry "CanDisplayPackageDialogBox = TRUE", LogTypeVerbose	
	

	' Ultimate and Enterprise SKU's allow for more than one Language Pack to be installed at a time.

	If IsInstallationUltimateEnterprise Then
		g_sPackageDialogBox = g_oXMLPackageList.GetHTMLEx( "CheckBox", "LanguagePacks" ) 
	Else
		' Convert property LanguagePacks back into a non-array if it's an array. 
		If isArray(property("LanguagePacks")) then
			oProperties.Item("LanguagePacks") = property("LanguagePacks")(0)
		End if

		g_sPackageDialogBox = g_oXMLPackageList.GetHTMLEx( "Radio", "LanguagePacks" ) 
	End if
	

	' Ensure that the default Language Pack has been added to the list.	

	for each oItem in oEnvironment.ListItem("ImageLanguage").Keys
		g_sPackageDialogBox = "</label>&nbsp;&nbsp;<b>(Already installed in OS)</b></div>" & vbNewLine & g_sPackageDialogBox
		g_sPackageDialogBox = "<img src='ItemIcon1.png' /><label for='DefaultLP' class=TreeItem>" & MarkupName(oItem) & g_sPackageDialogBox
		If IsInstallationUltimateEnterprise Then
			g_sPackageDialogBox = "<input name=LanguagePacks type=checkbox id='DefaultLP' value='DEFAULT' checked disabled />" & g_sPackageDialogBox
		ElseIf property("LanguagePacks") <> "" then
			g_sPackageDialogBox = "<input name=LanguagePacks type=Radio id='DefaultLP' value='DEFAULT' />" & g_sPackageDialogBox
		Else 
			g_sPackageDialogBox = "<input name=LanguagePacks type=Radio id='DefaultLP' value='DEFAULT' checked />" & g_sPackageDialogBox
		End if			
		g_sPackageDialogBox = "<div onmouseover=""javascript:this.className = 'DynamicListBoxRow-over';"" onmouseout=""javascript:this.className = 'DynamicListBoxRow';"" >" & g_sPackageDialogBox 
	next
	
End function

Dim g_LastTaskSequence

Function LanguagePack_Initialization

	If g_sPackageDialogBox = "" then
		g_LastTaskSequence = oEnvironment.Item("TaskSequenceID")
		CanDisplayPackageDialogBox

	ElseIf oEnvironment.Item("TaskSequenceID") <> "" and g_LastTaskSequence <> oEnvironment.Item("TaskSequenceID") then
		oLogging.CreateEntry "The TaskSequenceID has changed, Refresh the cache.", LogTypeInfo
		g_LastTaskSequence = oEnvironment.Item("TaskSequenceID")
		g_oXMLPackageList = empty
		CanDisplayPackageDialogBox

	End if

	PackagesListBox.InnerHTML = g_sPackageDialogBox
	PopulateElements
	
End function

Function CustomPackageFilter( sGuid, oItem )
	
	Dim oLang
	Dim bMatch

	' Check the languages that are installed and remove  
	' them from the list of languages to be installed.
	' These languages will then be added to the list later with
	' the text -(Already installed in OS)
	
	for each oLang in oEnvironment.ListItem("ImageLanguage").Keys 
		If oItem.SelectSingleNode("./Language").text = oLang then
			bMatch = TRUE
			Exit For
		End if
	next
	
	If bMatch then
		CustomPackageFilter = false
	Else	
		CustomPackageFilter = True
		oItem.SelectSingleNode("./Name").Text = MarkupName (oUtility.SelectSingleNodeString(oItem,"./Language"))
	End if
	
End function 


Function IsInstallationUltimateEnterprise

	IsInstallationUltimateEnterprise = oUtility.IsHighEndSKUEx( oEnvironment.Item("ImageFlags") )

End function

