' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTIDiskUtility.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:	Utility functions for disk operations
' // 
' // Usage:	 <script language="VBScript" src="ZTIDiskUtility.vbs"/>
' // 
' // ***************************************************************************

Option Explicit

Const ZTIDiskPart_Script = "4acfedd2-f865-4b32-b8dd-8e7841b235c5"


Class ZTIDisk   ' VBSCript Wrapper Class around the WMI Win32_DiskDrive class

	Private g_Win32_DiskDrive

	Property Let Disk ( iDisk )
		Dim oDisk

		TestAndFail isNumeric(iDisk), 7801, "Verify iDisk index is numeric: " & iDisk
		If iDisk <> Disk or isempty(Disk) then
			g_Win32_DiskDrive = empty
			For Each oDisk in objWMI.ExecQuery( "SELECT * FROM Win32_DiskDrive WHERE Index=" & iDisk )
				oLogging.CreateEntry "New ZTIDisk : " & oDisk.Path_ ,LogTypeInfo
				set g_Win32_DiskDrive = oDisk
				Exit Property
			next
		End if

	End property

	Property Set Disk ( oDisk )
		TestAndFail isobject(oDisk), 7802, "Verify oDisk index is an object. " & varType(oDisk)
		TestAndFail ucase(oDisk.CreationClassName) = "WIN32_DISKDRIVE", 7803, "Verify Disk is correct WMI Object." & oDisk.Path_
		oLogging.CreateEntry "New ZTIDisk : " & oDisk.Path_ ,LogTypeInfo
		set g_Win32_DiskDrive = oDisk
	End Property

	Function GetPartitionCount
		If not isempty(g_Win32_DiskDrive) then GetPartitionCount = g_Win32_DiskDrive.Partitions
	End function

	Function GetPartitions
		If not isempty(g_Win32_DiskDrive) then 
			set GetPartitions = objWMI.ExecQuery( "SELECT * FROM Win32_DiskPartition WHERE DiskIndex=" & Disk )
			oLogging.CreateEntry "GetPartitions: " & GetPartitions.Count ,LogTypeInfo
		End if
	End function

	Function GetDiskPartition( iPartition )
		set GetDiskPartition = new ZTIDiskPartition
		GetDiskPartition.SetDiskPart Disk,iPartition
	End function

	Function isOSReady( sOSBuild )
		' Check to see if this disk is "capiable" of installing an OS
		Dim oUSBDevice
		Dim oUSBController
		Dim oUSBPnPEnum
		Dim sPnPID
		If not isempty(g_Win32_DiskDrive) then
			isOSReady = g_Win32_DiskDrive.InterfaceType <> "USB" and g_Win32_DiskDrive.InterfaceType <> "1394"

			If not isOSReady Then
				Exit Function
			End if

			If len(trim(sOSBuild)) = 0 then
				' Do nothing, OS Version not specified.
			ElseIf cint(mid(ucase(sOSBuild),1,1)) > 6 then
				For Each oUSBDevice in objWMI.ExecQuery("ASSOCIATORS OF {" & g_Win32_DiskDrive.Path_ & "} WHERE AssocClass = Win32_PnPDevice")
					For Each oUSBController in objWMI.ExecQuery("ASSOCIATORS OF {" & oUSBDevice.Path_ & "} WHERE AssocClass = Win32_USBControllerDevice")
						For Each oUSBPnPEnum in objWMI.ExecQuery("ASSOCIATORS OF {" & oUSBController.Path_ & "} WHERE AssocClass = Win32_PnPDevice")
							For Each sPnPID in oUSBPnPEnum.CompatibleID
								If sPnPID = "PCI\\CC_0C0330" then
									oLogging.CreateEntry "Found USB 3.0 Ready Disk: " & g_Win32_DiskDrive.Path_ ,LogTypeInfo
									isOSReady = True ' USB 3.0
									Exit Function
								End if
							Next
						Next
					Next
				Next
			End if

		End if
	End function


	Property Get Disk
		If not isempty(g_Win32_DiskDrive) then Disk = cint(g_Win32_DiskDrive.Index)
	End Property

	Function IsRemovable
		Dim cap
		IsRemovable = False
		If not isempty(g_Win32_DiskDrive) then 
			For each Cap in g_Win32_DiskDrive.Capabilities
				If Cap = 7 then
					IsRemovable = True
				End if
			next
		End if
	End Function

	Function SizeAsString
		If not isempty(g_Win32_DiskDrive) then SizeAsString = FormatLargeSize(g_Win32_DiskDrive.Size)
	End function

	Function oWMI
		If not isempty(g_Win32_DiskDrive) then set oWMI = g_Win32_DiskDrive
	End Function

	' // instance of Win32_DiskDrive
	' // {
	' // 	BytesPerSector = 512;
	' // 	Capabilities = {3, 4};
	' // 	CapabilityDescriptions = {"Random Access", "Supports Writing"};
	' // 	Caption = "XXXXXXXXX";
	' // 	ConfigManagerErrorCode = 0;
	' // 	ConfigManagerUserConfig = FALSE;
	' // 	CreationClassName = "Win32_DiskDrive";
	' // 	Description = "Disk drive";
	' // 	DeviceID = "\\\\.\\PHYSICALDRIVE1";
	' // 	FirmwareRevision = "0002";
	' // 	Index = 1;
	' // 	InterfaceType = "IDE";
	' // 	Manufacturer = "(Standard disk drives)";
	' // 	MediaLoaded = TRUE;
	' // 	MediaType = "Fixed hard disk media";
	' // 	Model = "XXXXXXXXX";
	' // 	Name = "\\\\.\\PHYSICALDRIVE1";
	' // 	Partitions = 1;
	' // 	PNPDeviceID = "IDE\\XXXXXXXXX....";
	' // 	SCSIBus = 0;
	' // 	SCSILogicalUnit = 0;
	' // 	SCSIPort = 0;
	' // 	SCSITargetId = 1;
	' // 	SectorsPerTrack = 63;
	' // 	SerialNumber = "XXXXXXXXX";
	' // 	Signature = 1234567890;
	' // 	Size = "500105249280";
	' // 	Status = "OK";
	' // 	SystemCreationClassName = "Win32_ComputerSystem";
	' // 	SystemName = "PICKETTK";
	' // 	TotalCylinders = "60801";
	' // 	TotalHeads = 255;
	' // 	TotalSectors = "976768065";
	' // 	TotalTracks = "15504255";
	' // 	TracksPerCylinder = 255;
	' // };

End Class


Class ZTIDiskPartition   ' VBSCript Wrapper Class around the WMI Win32_DiskPartition class

	Private g_Win32_DiskPartition
	Private g_Win32_LogicalDisk

	Private Function GetLogicalDiskToPartition ( oDiskType1 )
		Dim oDisk
		TestAndFail not isEmpty(oDiskType1), 7804, "Verify First object is set."
		For Each oDisk in objWMI.ExecQuery("ASSOCIATORS OF {" & oDiskType1.Path_ & "} WHERE AssocClass = Win32_LogicalDiskToPartition")
			set GetLogicalDiskToPartition = oDisk
			oLogging.CreateEntry "New ZTIDiskPartition : " & oDiskType1.Path_  & "    " & oDisk.Path_ , LogTypeInfo
			Exit Function
		Next
		set GetLogicalDiskToPartition = nothing
	End function


	Property Let Drive ( sDrive )
		Dim oLogicalDisk
		Dim oDiskPart

		If Drive <> left(sDrive,1) & ":" then
			g_Win32_LogicalDisk = empty
			g_Win32_DiskPartition = empty

			For Each oLogicalDisk in objWMI.ExecQuery( "SELECT * FROM Win32_LogicalDisk WHERE DeviceId ='" & left(sDrive,1) & ":" & "'" )
				set g_Win32_LogicalDisk = oLogicalDisk
				set g_Win32_DiskPartition = GetLogicalDiskToPartition (g_Win32_LogicalDisk)
				Exit Property
			next
		End if
	End Property

	Property Set Drive ( oDrive )
		TestAndFail isobject(oDrive), 7805, "Verify oDisk index is an object. " & varType(oDisk)
		TestAndFail ucase(oDrive.CreationClassName) = "WIN32_LOGICALDISK", 7806, "Verify Drive is correct WMI Object." & oDrive.Path_
		set g_Win32_LogicalDisk = oDrive
		set g_Win32_DiskPartition = GetLogicalDiskToPartition (g_Win32_LogicalDisk)
	End Property

	Property Get Drive
		If IsValidObject(g_Win32_LogicalDisk) Then Drive = g_Win32_LogicalDisk.DeviceID
	End Property


	Property Get oWMIDrive( bForceMount )
		Dim i 
		' It is possible that the partition is *not* assigned a drive letter.
		If bForceMount and not IsValidObject(g_Win32_LogicalDisk) then
			If IsValidObject(g_Win32_DiskPartition) Then
				RunDiskPartSilent array("Select Disk " & Disk, "Select Partition " & Partition, "Assign", "Exit")
				i = 1 
				While i < 120 and not IsValidObject(g_Win32_LogicalDisk)
					oUtility.SafeSleep 500
					set g_Win32_LogicalDisk = GetLogicalDiskToPartition (g_Win32_DiskPartition)
					i = i + 1
				WEnd
			End if
			TestAndFail not isempty(g_Win32_LogicalDisk), 7807, "Verify Disk has been mapped."
		End if
		If not isempty(g_Win32_LogicalDisk) then 
			set oWMIDrive = g_Win32_LogicalDisk
		Else
			set oWMIDrive = nothing
		End if
	End Property


	Function SetDiskPart( iDisk, iPartition ) 
		Dim oLogicalDisk
		Dim oDiskPart

		If iDisk <> Disk or iPartition <> Partition then
			g_Win32_LogicalDisk = empty
			g_Win32_DiskPartition = empty

			For Each oDiskPart in objWMI.ExecQuery( "SELECT * FROM Win32_DiskPartition WHERE DiskIndex=" & iDisk & " and Index=" & ( iPartition - 1 ) ) 
				set g_Win32_DiskPartition = oDiskPart
				set g_Win32_LogicalDisk = GetLogicalDiskToPartition (g_Win32_DiskPartition)
				Exit Function
			next
		End if
	End function


	Property Set DiskPart ( oDisk )
		TestAndFail isobject(oDisk), 7808, "Verify oDisk index is an object. " & varType(oDisk)
		TestAndFail ucase(oDisk.CreationClassName) = "WIN32_DISKPARTITION", 7809, "Verify DiskPart is correct WMI Object." & oDisk.Path_
		set g_Win32_DiskPartition = oDisk
		set g_Win32_LogicalDisk = GetLogicalDiskToPartition (g_Win32_DiskPartition)
	End Property

	Property Get oWMIDiskPart
		If IsValidObject(g_Win32_DiskPartition) then
			set oWMIDiskPart = g_Win32_DiskPartition
		End if
	End Property


	Property Get Disk
		If IsValidObject(g_Win32_DiskPartition) Then
			Disk = cint(g_Win32_DiskPartition.DiskIndex)
		End if
	End Property

	Property Get Partition
		If IsValidObject(g_Win32_DiskPartition) Then
			Partition = cint(g_Win32_DiskPartition.Index) + 1
		End if
	End Property


	Function GetDiskObject
		set GetDiskObject = new ZTIDisk
		GetDiskObject.Disk = Disk
	End function

	Function isOSReady( sOSBuild, osFlags )

		If IsValidObject(g_Win32_DiskPartition) Then
			isOSReady = g_Win32_DiskPartition.Size / 1000 / 1000 >= GetMinimumDiskPartitionSizeMB
		End if

	End function 


	' Rule: It is possible for a Win32_DiskPartition Object to *not* have an associated win32_LogicalDisk object
	' Rule: It is possible for a win32_LogicalDisk Object to *not* have an associated Win32_DiskPartition object

	'instance of Win32_DiskPartition
	'{
	'	BlockSize = "512";
	'	Bootable = FALSE;
	'	BootPartition = FALSE;
	'	Caption = "Disk #1, Partition #0";
	'	Description = "Installable File System";
	'	DeviceID = "Disk #1, Partition #0";
	'	DiskIndex = 1;
	'	Index = 0;
	'	Name = "Disk #1, Partition #0";
	'	NumberOfBlocks = "976766976";
	'	PrimaryPartition = TRUE;
	'	Size = "500104691712";
	'	StartingOffset = "1048576";
	'	Type = "Installable File System";
	'};


	'instance of Win32_LogicalDisk
	'{
	'	Access = 0;
	'	Caption = "C:";
	'	Compressed = FALSE;
	'	Description = "Local Fixed Disk";
	'	DeviceID = "C:";
	'	DriveType = 3;
	'	FileSystem = "NTFS";
	'	FreeSpace = "6641623040";
	'	MediaType = 12;
	'	Size = "79698063360";
	'	SupportsDiskQuotas = FALSE;
	'	SupportsFileBasedCompression = TRUE;
	'	VolumeName = "OSDisk";
	'	VolumeSerialNumber = "XXXXXXXXX";
	'};

End class


' -------------------------------------------------------------------------------------------------
'
'  Wrapper routines for objects above, they all return WMI objects
'

Function AllDiskDrivesEx( sSelectionCriteria )
	set AllDiskDrivesEx = objWMI.ExecQuery( "SELECT * FROM Win32_DiskDrive " & sSelectionCriteria )
End Function

Function AllDiskDrives
	set AllDiskDrives = AllDiskDrivesEx(empty)
End Function


Function AllDiskPartEx ( sSelectionCriteria )
	set AllDiskPartEx = objWMI.ExecQuery( "SELECT * FROM Win32_DiskPartition " & sSelectionCriteria )
End function

Function AllDiskPart
	set AllDiskPart = AllDiskPartEx(empty)
End Function 


Function AllLogicalDrivesEx ( sSelectionCriteria )
	set AllLogicalDrivesEx = objWMI.ExecQuery( "SELECT * FROM Win32_LogicalDisk " & sSelectionCriteria )
End function

Function AllLogicalDrives
	set AllLogicalDrives = AllLogicalDrivesEx(empty)
End Function 


Function GetBootDriveEx ( bForced, sOSBuild, bBootFiles )
	Dim oWMIDiskPart
	Dim oDiskPart
	Dim oDisk
	Dim oLogical
	Dim sCurrentARSetting
	Dim iRetVal

	set GetBootDriveEx = nothing
	
	' Get current Autorun status
	' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun 
	' 0xFF - Disables all drive types
	If oEnv("SystemDrive") <> "X:" then
		on error resume next
		sCurrentARSetting = oShell.Regread( "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun")
		if not Err then
			oLogging.CreateEntry "The current autorun setting is - " & sCurrentARSetting, logTypeInfo
		End if
		
		oLogging.CreateEntry "Disabling Autorun", logTypeInfo
		iRetVal = oUtility.regwrite ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun", 255 )
		On Error Goto 0
	End if 
	
	oLogging.CreateEntry "Find the boot drive (if any) [" & bForced & "] [" & sOSBuild & "] [" & bBootFiles & "]" , logTypeInfo
	For each oWMIDiskPart in AllDiskPartEx ( " WHERE Bootable = TRUE or BootPartition = TRUE " )
		oLogging.CreateEntry "Found bootable drive: " & oWMIDiskPart.Path_ , logTypeVerbose
		set oDiskPart = new ZTIDiskPartition
		set oDiskPart.DiskPart = oWMIDiskPart
		set oDisk = oDiskPart.GetDiskObject
		If oDisk.isOSReady( sOSBuild ) then
			set oLogical = oDiskPart.oWMIDrive( bForced )
			If isValidObject(oLogical) then
				If not bBootFiles then
					oLogging.CreateEntry "Found bootable drive (No Boot File Test) [ " & oLogical.DeviceID & " ]: " & oLogical.Path_ , logTypeInfo
					set GetBootDriveEx = oDiskPart
				ElseIf oFSO.FileExists( oLogical.DeviceID & "\ntldr" ) or oFSO.FileExists( oLogical.DeviceID & "\bootmgr" ) or oFSO.FileExists( oLogical.DeviceID & "\EFI\Microsoft\Boot\bootmgr.efi" ) then
					oLogging.CreateEntry "Found bootable drive [ " & oLogical.DeviceID & " ]: " & oLogical.Path_ , logTypeInfo
					set GetBootDriveEx = oDiskPart
				Else
					oLogging.CreateEntry "No boot files found: " & oLogical.Path_ , logTypeInfo
				End if
			End if
		End if
	next
	
	If GetBootDriveEx is nothing then
		oLogging.CreateEntry "No boot drives found. None.", logTypeInfo
	End if
	
	'Revert Autorun status to orginal setting
	If oEnv("SystemDrive") <> "X:" then
		if isempty(sCurrentARSetting) then
			sCurrentARSetting = 0
		End if 
		
		if  not isNull(sCurrentARSetting) then 
			oLogging.CreateEntry "Reverting autorun setting to - " & sCurrentARSetting, logTypeInfo
			iRetVal = oUtility.regwrite ( "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDriveTypeAutoRun", sCurrentARSetting )
		End if
	End if
	
End Function


' -------------------------------------------------------------------------------------------------
'
'  Support Routines
'

Function IsValidObject( oTest )
	IsValidObject = False
	If isObject ( oTest ) then
		If not oTest is nothing then
			IsValidObject = True
		End if
	End if	
End function

Function RunDiskPartSilent( aCommands )

	Dim sDiskPart
	Dim sTempFile
	Dim sCommand

	sTempFile = oFileHandling.GetTempFile

	With oFSO.CreateTextFile(sTempFile,true,false)
		For each sCommand in aCommands
			oLogging.CreateEntry "DISKPART + " & sCommand, LogTypeInfo
			.WriteLine sCommand
		next
		.Close
	End With

	RunDiskPartSilent = oUtility.RunWithConsoleLoggingAndHidden("DiskPart.exe /s " & sTempFile)

End function 

Function RunDiskPart( aCommands )

	Dim sDiskPart
	Dim sTempFile
	Dim sCommand

	sTempFile = oFileHandling.GetTempFile

	With oFSO.CreateTextFile(sTempFile,true,false)
		For each sCommand in aCommands
			oLogging.CreateEntry "DISKPART + " & sCommand, LogTypeInfo
			.WriteLine sCommand
		next
		.Close
	End With

	RunDiskPart = oUtility.RunWithConsoleLogging("DiskPart.exe /s " & sTempFile)

End function 


Function FormatLargeSize( lSize )

	Dim i
	For i = 1 to len(" KMGTPEZY")
		If cdbl(lSize) < 1000 ^ i then
			FormatLargeSize = int(cdbl(lSize)/(1000^(i-1))) & " " & mid(" KMGTPEZY",i,1) & "B"
			Exit function
		End if
	next

End function


' -------------------------------------------------------------------------------------------------
'
'  BCD
'

Function RunBCDBoot()
	If oEnv("SystemDrive") = "X:" then
		RunBCDBoot = RunBCDBootEx( oUtility.GetOSTargetDriveLetter & "\windows", empty )
	Else
		RunBCDBoot = RunBCDBootEx( oFileHandling.GetWindowsFolder, empty )
	End if
End function

Function RunBCDBootEx( sWindowsTarget, sParams )

	Dim sProgram
	Dim iRC

	iRC = oUtility.FindFile("BCDBoot.exe",sProgram)
	TestAndFail iRC, 7811, "FindFile: BCDBoot.exe"
	sProgram = """" & sProgram & """"

	If oLogging.Debug then
		sProgram = sProgram & " /v"
	End if

	If OEnvironment.Item("UILanguage") = "" Then
		OEnvironment.Item("UILanguage") = oUtility.RegReadEx("HKCU\Control Panel\International\LocaleName",False)
	End If
	TestAndFail OEnvironment.Item("UILanguage") <> "", 7810, "Verify UILanguage is set."

	RunBCDBootEx = oUtility.RunWithConsoleLoggingAndHidden(sProgram & " " & sWindowsTarget  & " /l " & oEnvironment.Item("UILanguage") & " " & sParams)

End function


' -------------------------------------------------------------------------------------------------
'
'  Destination Disk/Partition/LogicalDrive routines
'

Function GetMinimumDiskPartitionSizeMB
	GetMinimumDiskPartitionSizeMB = 15 * 1000  ' 15 GB

	If oEnvironment.Item("ImageBuild") <> "" then
		If left(oEnvironment.Item("ImageBuild"),3) = "5.1" then
			GetMinimumDiskPartitionSizeMB = 1500  ' 1.5 GB
		ElseIf left(oEnvironment.Item("ImageBuild"),3) = "5.2" then
			GetMinimumDiskPartitionSizeMB = 3 * 1000   ' 3 GB
		ElseIf oEnvironment.Item("ImageFlags") <> "" then
			If inStr(1,oEnvironment.Item("ImageFlags"),"SERVER",vbTextCompare) <> 0 then
				GetMinimumDiskPartitionSizeMB = 10 * 1000 ' 32 GB
			End if
		End if
	End if

End function

Function GetFirstPossibleSystemDrive

	Dim oWMIDisk
	Dim oDisk
	Dim oWMIDiskPart
	Dim oDiskPart

	If oEnv("SystemDrive") <> "X:" then
	
		GetFirstPossibleSystemDrive = oEnv("SystemDrive") 
		oLogging.CreateEntry "Found OS Disk: " & oEnv("SystemDrive")  , logTypeInfo
		
	Else
	
		for each oWMIDisk in AllDiskDrives
			set oDisk = new ZTIDisk
			set oDisk.Disk = oWMIDisk
			TestAndFail not isempty(oDisk.oWMI), 7814, "Verify class created:  " & oWMIDisk.Path_
			If oDisk.isOSReady( oEnvironment.Item("ImageBuild") ) then

				oLogging.CreateEntry "Found Possible OS TargetDisk: " & oWMIDisk.Path_ , logTypeInfo
				for each oWMIDiskPart in oDisk.GetPartitions
					oLogging.CreateEntry "Found Possible OS Target Partition: " & oWMIDiskPart.Path_ , logTypeInfo
					set oDiskPart = new ZTIDiskPartition
					set oDiskPart.DiskPart = oWMIDiskPart
					TestAndFail not isempty(oDiskPart.oWMIDiskPart), 7815,"Verify WMI Object was accepted by ZTIDiskPartition"
					If not oDiskPart.isOSReady( oenvironment.Item("ImageBuild"), oenvironment.Item("ImageFlags") ) then
						oLogging.CreateEntry "Target Partition not big enough: " & oWMIDiskPart.Path_ , logTypeInfo
					ElseIf not isEmpty(GetFirstPossibleSystemDrive) then
						oLogging.CreateEntry "Found More than one candidate OS TargetDisk: " & oWMIDiskPart.Path_ , logTypeInfo
					ElseIf not isEmpty(oDiskPart.Drive) then
						oLogging.CreateEntry "Found Drive: " & oDiskPart.Drive , logTypeInfo
						GetFirstPossibleSystemDrive = oDiskPart.Drive
					Else
						oLogging.CreateEntry "Partition was found, however no drive was associated (not formatted?) Skip: " & oWMIDiskPart.Path_ , logTypeInfo
					End if
				next
			Else
				oLogging.CreateEntry "Found Possible NonOS Disk: " & oWMIDisk.Path_ , logTypeInfo
			End if
		next
	
	End if


	If not isEmpty(GetFirstPossibleSystemDrive) then
			oLogging.CreateEntry "Found FirstPossibleSystemDrive: " & GetFirstPossibleSystemDrive , logTypeInfo
	Else
			oLogging.CreateEntry "Did not find disk." , logTypeInfo
	End if


End function



' -------------------------------------------------------------------------------------------------
'
'  Legacy Routines
'


Function MarkActive(sDrive)

	Dim oDrive

	Set oDrive = new ZTIDiskPartition
	oDrive.Drive = sDrive
	If not isEmpty( oDrive.Drive ) then
		MarkActive = RunDiskPartSilent ( array ( "List Vol", "Select Disk " & oDrive.Disk , "Select Partition " & oDrive.Partition , "Active", "Detail Part", "Exit" ) )
	Else
		MarkActive = RunDiskPartSilent ( array ( "List Vol", "Select Vol " & sDrive, "Active",  "Detail Part", "Exit" ) )
	End if
	

End Function


Function GetDiskPartitionCount (iDisk)

	Dim oDisk

	set oDisk = new ZTIDisk
	oDisk.Disk = iDisk
	GetDiskPartitionCount = oDisk.GetPartitionCount

End Function


Function GetNotActiveDrive

	Dim oDrive
	Dim oWMIDiskPart

	For each oWMIDiskPart in AllDiskPartEx(" WHERE BootPartition = FALSE ")
		Set oDrive = new ZTIDiskPartition
		set oDrive.DiskPart = oWMIDiskPart
		If not isEmpty(oDrive.Drive) then
			GetNotActiveDrive = oDrive.Drive

			exit Function
		End If
	next

	GetNotActiveDrive = False

End Function


function GetDiskForDrive (sDriveLetter)

	Dim oDrive

	Set oDrive = new ZTIDiskPartition
	oDrive.Drive = sDriveLetter
	GetDiskForDrive = oDrive.Disk
	If isEmpty(GetDiskFOrDrive) then GetDiskForDrive = -1 ' Backwards compatibility

End Function


Function GetPartitionSize (sDriveLetter)

	Dim oDrive

	GetPartitionSize = 0
	Set oDrive = new ZTIDiskPartition
	oDrive.Drive = sDriveLetter
	If not isEmpty( oDrive.oWMIDrive(false) ) then
		GetPartitionSize = oDrive.oWMIDrive(false).Size
	End if

End Function


Function GetDiskSize (iDrive)

	Dim oDisk

	GetDiskSize = -1
	set oDisk = new ZTIDisk
	oDisk.Disk = iDrive
	If not isempty(oDisk.Disk) then
		GetDiskSize = cLng(Fix(oDisk.oWMI.Size / 1024 /1024))
	End if

End Function


Function GetBootDrive
	Dim oBootDrive

	set oBootDrive = GetBootDriveEx( false, "0.0.0.0", false )
	If not oBootDrive is nothing then 
		GetBootDrive = oBootDrive.Drive
	Else
		GetBootDrive = Failure
	End if

End function


' -------------------------------------------------------------------------------------------------
'
'  Routines that should be phased out.
'



Function GetDiskFreeSpace ( iDisk )

	oLogging.CreateEntry "ZTIDiskUtility!GetDiskFreeSpace should be deprecated, does not handle avaible space for a new partition", LogTypeDeprecated

	Dim oDisk
	Dim oWMIDiskPart
	Dim oDiskPart
	Dim iSum

	set oDisk = new ZTIDisk
	oDisk.Disk = iDisk

	For each oWMIDiskPart in oDisk.GetPartitions
		set oDiskPart = new ZTIDiskPartition
		set oDiskPart.DiskPart = oWMIDiskPart 
		iSum = iSum + clng(oDiskPart.oWMIDiskPart.Size / 1024 /1024)
	Next

	GetDiskFreeSpace = clng(oDisk.oWMI.Size / 1024 /1024) - iSum

End function


' Determine drive status as compared to the custom partition config
'
' Return:
'   0 = Partition exists and meets criteria
'   1 = Partition exists and does not meet the criteria
'   -1= Partition does not exist.

Function MatchPartitionConfiguration (iDriveIndex, iPartitionIndex, sDriveLetter, iMinSizeMB)

	Dim oDiskPart
	Dim oWMIDrv

	oLogging.CreateEntry "ZTIDiskUtility!MatchPartitionConfiguration should be deprecated. Legacy Solution.", LogTypeDeprecated

	set oDiskPart = new ZTIDiskPartition
	oDiskPart.SetDiskPart iDriveIndex, iPartitionIndex + 1
	MatchPartitionConfiguration = -1
	If not isEmpty(oDiskPart.Drive) then
		MatchPartitionConfiguration = FAILURE
		set oWMIDrv = oDiskPart.oWMIDrive(false)
		If not oWMIDrv is nothing then 
			If UCase(oDiskPart.Drive) = UCase(sDriveLetter) And oWMIDrv.DriveType = 3 THen
				If (oWMIDrv.FileSystem = "FAT" Or oWMIDrv.FileSystem = "FAT32" Or oWMIDrv.FileSystem = "NTFS") Then
					If cLng(oWMIDrv.Size / 1024 / 1024) >= cLng(iMinSizeMB) Then
						MatchPartitionConfiguration = SUCCESS
					End if
				End if
			End if
		End If
	End if

End Function


