' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_OfflineData.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Scripts for offline data migration page
' // 
' // ***************************************************************************

Option Explicit

Dim g_hasSomethingBeenChecked
Dim g_IsDisk0Bitlocked
dim g_OSDrive
Dim g_WindowsCount

Function AddDiskToTableByDll ( sCmd, EncryptedDrives )

	Dim oRow
	Dim oCol
	Dim aCmd
	Dim sDiskID
	Dim sDiskDesc
	Dim sSelected
	Dim sType
	Dim sVer, sNewVer
	Dim sArchitecture, sNewArchitecture
	Dim sDrive

	AddDiskToTableByDll = false

	oLogging.CreateEntry "Add Entry to Dialog: " & replace(sCmd,vbTab,"  "), LogTypeInfo

	aCmd = split(sCmd,vbTab)
	If aCmd(0) = "EMPTY" then
		Exit Function 
	End if

	sNewVer = Left(oEnvironment.Item("ImageBuild"), 3)
	sNewArchitecture = UCase(oEnvironment.Item("ImageProcessor"))
	sType = ""
	sSelected = ""
	sDiskDesc = "Disk " & aCmd(1) & " Partition " & aCmd(2)
	sDiskID = "Disk" & aCmd(1) & "part" & aCmd(2)

	If len(aCmd(7)) = 0 then

		oLogging.CreateEntry "Nothing to display " & sDiskID & " OK TO Skip!", LogTypeInfo
		exit function 

	ElseIf Instr(1,EncryptedDrives,left(aCmd(7),2),vbTextCompare) <> 0 then

		' This drive is BitLockered
		sDiskDesc = sDiskDesc & " <div style='display:inline; font: bold;' class='errmsg'>?? BitLocker Protected ??</div>"
		sSelected = "disabled"
		sType = ""

		g_IsDisk0Bitlocked = g_IsDisk0Bitlocked or true ' (cint(aCmd(1)) = 0)

	ElseIf ubound(aCmd) < 8 then

		oLogging.CreateEntry "Nothing to display Disabled " & sDiskID & " OK TO Skip!", LogTypeInfo
		exit function 

	ElseIf aCmd(5) < 10000 then

		oLogging.CreateEntry "Partition too small " & sDiskID & " OK TO Skip! " & aCmd(5), LogTypeInfo
		exit function 

	ElseIf aCmd(9) < 15000 then

		' Not enough Free Space
		sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
		oLogging.CreateEntry "Not enough disk space on " & sDiskID & " OK TO Skip! " & aCmd(5), LogTypeInfo
		sSelected = "Unknown"
		sType = "Drive Full"

	Else

		sDrive = Left(aCmd(7),2)
		If oFSO.FileExists(sDrive & "\Windows\System32\ntoskrnl.exe") Then

			sVer = Left(oFSO.GetFileVersion(sDrive & "\Windows\System32\ntoskrnl.exe"),3)
			If oFSO.FolderExists(sDrive & "\Program Files (x86)") then
				sArchitecture = "X64"
			Else
				sArchitecture = "X86"
			End if
			sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]" ' \Windows"

			If sArchitecture = "X64" and sNewArchitecture = "X86" then
				sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
				oLogging.CreateEntry "Existing OS is x64, new OS is x86, user state migration not supported.", LogTypeInfo
				sSelected = "disabled"
				sType = "Architecture changing from x64 to x86"
			ElseIf sVer <= sNewVer then
				sType = "\Windows " & sVer
				AddDiskToTableByDll = True
				If g_hasSomethingBeenChecked <> true and cint(aCmd(1)) = 0 then
					sSelected = "checked"
					g_hasSomethingBeenChecked = true
					g_OSDrive = sDrive
				End if
			Else
				sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
				oLogging.CreateEntry "Existing OS " & sDiskID & " is too new, " & sVer & " > " & sNewVer, LogTypeInfo
				sSelected = "disabled"
				sType = "Too new operating system"
			End if

		ElseIf oFSO.FileExists(sDrive & "\Winnt\System32\ntoskrnl.exe") Then

			sVer = Left(oFSO.GetFileVersion(sDrive & "\Winnt\System32\ntoskrnl.exe"),3)
			If oFSO.FolderExists(sDrive & "\Program Files (x86)") then
				sArchitecture = "X64"
			Else
				sArchitecture = "X86"
			End if
			sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]" ' \WinNT"

			If sArchitecture = "X64" and sNewArchitecture = "X86" then
				sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
				oLogging.CreateEntry "Existing OS is x64, new OS is x86, user state migration not supported.", LogTypeInfo
				sSelected = "disabled"
				sType = "Architecture changing from x64 to x86"
			ElseIf sVer <= sNewVer then
				sType = "\WinNT " & sVer
				AddDiskToTableByDll = True
				If g_hasSomethingBeenChecked <> true and cint(aCmd(1)) = 0 then
					sSelected = "checked"
					g_hasSomethingBeenChecked = true
					g_OSDrive = sDrive
				End if
			Else
				sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
				oLogging.CreateEntry "Existing OS " & sDiskID & " is too new, " & sVer & " > " & sNewVer, LogTypeInfo
				sSelected = "disabled"
				sType = "Too new operating system"
			End if

		Else

			sDiskDesc = sDiskDesc & " [" & aCmd(8) & "]"
			sSelected = "disabled"
			sType = "No operating system"

		End if

	End if

	' # -------------------------------------

	set oRow = document.createElement("<tr onmouseover=""javascript:this.className = 'DynamicListBoxRow-over';"" onmouseout=""javascript:this.className = 'DynamicListBoxRow';"" >")
	
	set oCol = document.createElement("TD")
	If instr(1,sSelected,"disabled",vbTextCompare) <> 0 then
		oCol.innerHTML = "<input name='OfflineMigration' id='" & sDiskID & "' type='radio' value='' language=vbscript onclick='DriveSelected' " & sSelected & " />"
	Else
		oCol.innerHTML = "<input name='OfflineMigration' id='" & sDiskID & "' type='radio' value='" & left(aCmd(7),2) & "' language=vbscript onclick='DriveSelected' " & sSelected & " />"
	End if
	oRow.appendChild oCol

	set oCol = document.createElement("TD")
	oCol.innerHTML = "<label for='" & sDiskID & "' class=TreeItem >" & sDiskDesc & "</label>"
	oRow.appendChild oCol

	set oCol = document.createElement("TD")
	oCol.innerHTML = FormatLargeSize(clng(aCmd(5)) * 1024 * 1024 )
	oRow.appendChild oCol

	set oCol = document.createElement("TD")
	oCol.innerHTML = sType
	oRow.appendChild oCol
	
	DiskTable.FirstChild.appendChild oRow

End Function


Function InitializeOfflineDataPage

	Dim i
	Dim oDisk
	Dim EncryptedDrives

	g_IsDisk0Bitlocked = false

	for each oDisk in GetObject("winmgmts:\\.\root\CIMV2\Security\MicrosoftVolumeEncryption"). _
		ExecQuery("SELECT * FROM Win32_EncryptableVolume WHERE ProtectionStatus = 2",,48)
			EncryptedDrives = EncryptedDrives & vbTab & oDisk.DriveLetter
	next

	oDisk = Empty
	on error resume next
		oDisk = oUtility.BDDUtility.HiddenPartitionsToDrives
	on error goto 0 

	g_WindowsCount = 0

	If not isEmpty(oDisk) then
		for i = 0 to ubound(oDisk)
			If ubound(split(oDisk(i),vbTab)) >= 8 then
				If AddDiskToTableByDll(oDisk(i), EncryptedDrives) then
					g_WindowsCount = g_WindowsCount + 1
				End if
			End if
		next
	End if

	If g_IsDisk0Bitlocked = true and g_hasSomethingBeenChecked <> true then

		oLogging.CreateEntry "A supported previous version of Windows was not found on this computer. Suspend BitLocker to save files and settings.", LogTypeInfo
		UDRadio1.Checked = true
		UDRadio2.disabled = true
		MSITTExtInfo.style.display = "inline"
		MSITText.InnerHTML = "A supported previous version of Windows was not found on this computer. Suspend BitLocker to save files and settings."

	ElseIf g_WindowsCount = 0 then

		oLogging.CreateEntry "A supported previous version of Windows was not found on this computer. Data and settings cannot be restored.", LogTypeInfo
		UDRadio1.Checked = true
		UDRadio2.disabled = true
		MSITTExtInfo.style.display = "inline"
		MSITText.InnerHTML = "A supported previous version of Windows was not found on this computer. Data and settings cannot be restored."

	Else

		If g_IsDisk0Bitlocked = true then
			oLogging.CreateEntry "A BitLocker-protected drive was found. Suspend BitLocker to save files and settings.", LogTypeInfo
			MSITTExtInfo.style.display = "inline"
			MSITText.InnerHTML = "A BitLocker-protected drive was found. Suspend BitLocker to save files and settings."
		End if

	End if

	If g_WindowsCount > 1 then

		for each oDisk in document.getElementsByName("OfflineMigration")
			oLogging.CreateEntry "OFfline Migration checked: [" & oProperties("OfflineMigration") & "] = [" & oDisk.Value & "]", LogTypeInfo
			If oProperties("OfflineMigration") <> "" and oProperties("OfflineMigration") = oDisk.Value then
				oDisk.checked = true
				exit for
			End if
		next

		MoreThanOneVolume.Style.Display = "inline"

		if UDRadio1.Checked then
			NoOffline
		end if 

	End if

	KeepPartitions.Disabled = true
	If UDRadio1.Checked and g_WindowsCount > 0 then
		KeepPartitions.Disabled = false
		If oProperties("doNotFormatAndPartition") = "YES" then
			KeepPartitions.checked = true
		Else
			KeepPartitions.checked = false
		End if
	End if

End function


' --------------------------

Function DriveSelected

	g_OSDrive = window.event.srcElement.Value 
	oLogging.CreateEntry "Item Selected: " & g_OSDrive, LogTypeInfo

End function

Function OnlineEnable

	Dim oDisk

	for each oDisk in document.getElementsByName("OfflineMigration")
		If oDisk.Value <> "" then
			oDisk.Disabled = false
		End if
	next

	KeepPartitions.Disabled = true

End function

Function NoOffline

	Dim oDisk

	for each oDisk in document.getElementsByName("OfflineMigration")
		oDisk.Disabled = true
	next

	If g_WindowsCount > 0 then
		KeepPartitions.Disabled = false
	End if

End function


Function RemovePropertyIfPresent ( oProperty )

	If oProperties.Exists( oProperty ) then
		oProperties.remove oProperty
	End if

End function


Function ValidateOfflineData

	If UDRadio1.Checked then

		oLogging.CreateEntry "No data to save", LogTypeInfo
		If KeepPartitions.Checked then
			oProperties("doNotFormatAndPartition") = "YES"
		Else
			RemovePropertyIfPresent "doNotFormatAndPartition"
		End if
		RemovePropertyIfPresent "UserDataLocation"
		RemovePropertyIfPresent "DestinationOSInstallType"
		RemovePropertyIfPresent "DestinationOSVariable"
		RemovePropertyIfPresent "OSDisk"
		RemovePropertyIfPresent "OSDWinPEWindir"
		RemovePropertyIfPresent "OriginalArchitecture"

	ElseIf UDRadio2.Checked then

		oLogging.CreateEntry "Migrate!", LogTypeInfo
		KeepPartitions.Disabled = true
		oProperties("doNotFormatAndPartition") = "YES"
		oProperties("UserDataLocation") = "AUTO"
		oProperties("DestinationOSInstallType") = "ByVariable"
		oProperties("DestinationOSVariable") = "OSDisk"
		oProperties("OSDisk") = g_OSDrive
		If oFSO.FolderExists(g_OSDrive & "\Windows") then
			oProperties("OSDWinPEWindir") = g_OSDrive & "\Windows"
		ElseIf oFSO.FolderExists(g_OSDrive & "\Winnt") then
			oProperties("OSDWinPEWindir") = g_OSDrive & "\Winnt"
		End if
		If oFSO.FolderExists(g_OSDrive & "\Program Files (x86)") then
			oProperties("OriginalArchitecture") = "X64"
		Else
			oProperties("OriginalArchitecture") = "X86"
		End if
					
	End if

	ValidateOfflineData = true

End Function

