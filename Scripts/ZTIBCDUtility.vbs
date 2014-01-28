' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTIBCDUtility.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:	Utility functions for bcd operations
' // 
' // Usage:     cscript.exe [//nologo] ztibcdutility.vbs
' //
' // Reference: http://technet.microsoft.com/en-us/magazine/2008.07.heyscriptingguy.aspx?pr=blog
' //            
' // 
' // ***************************************************************************
Option Explicit


'//---------------------------------------------------------------------------
'//
'// Support Routines
'//


'//---------------------------------------------------------------------------
'//
'// Functions that work for both WMI and BCDEdit.exe
'//

Function BCDObjectExists ( sGUID ) 
	BCDObjectExists = BCDObjectExistsEx ( "", sGUID ) 
End function 

Function BCDObjectExistsEx ( sStore, sGUID ) 
	Dim sBuffer
	Dim iRetVal

	BCDObjectExistsEx = False
	iRetVal = RunBCDEditEx("/enum " & sGUID , sStore, sBuffer )
	If iRetVal = SUCCESS then
		BCDObjectExistsEx = instr( 1, sBuffer, sGUID, vbTextCompare ) <> 0
		If not BCDObjectExistsEx then
			iRetVal = RunBCDEditEx("/enum " & sGUID & " /v", sStore, sBuffer )
			If iRetVal = SUCCESS then
				BCDObjectExistsEx = instr( 1, sBuffer, sGUID, vbTextCompare ) <> 0
			End if
		End if
	End if 
	
End function 

Function BCDBackupStore( sBackupFile )
	BCDBackupStore = RunBCDEdit( "/export """ & sBackupFile & """" ) = SUCCESS
End function

Function BCDGetCurrentGUID
	BCDGetCurrentGUID = BCDGetCurrentGUIDEx( "" ) 
End function

Function BCDGetCurrentGUIDEx ( sStore ) 
	' THis function requires WMI
	oBCDUtility.sBCDStore = sStore
	BCDGetCurrentGUIDEx = oBCDUtility.GetCurrentGUID
End function


Function GetBCDError
	oLogging.CreateEntry "GetBCDError.", LogTypeDeprecated
	GetBCDError  = ""
End function


'//---------------------------------------------------------------------------
'//
'// Functions that work for both WMI and BCDEdit.exe
'//

Function CreateRamDiskEntry( sDrive ) 
	 CreateRamDiskEntry = CreateRamDiskEntryEx( "", sDrive ) 
End function 

Function CreateRamDiskEntryEx( sStore, sDrive ) 

	If BCDObjectExists("{ramdiskoptions}") then
		oLogging.CreateEntry "{ramdiskoptions} already present.", LogTypeInfo
		exit Function 
	End if

	' Create entry.
	RunBCDEditEx "/Create {ramdiskoptions} -d ""Ramdisk Device Options""" , sStore, null 
	RunBCDEditEx "/Set {ramdiskoptions} ramdisksdidevice partition=" & sDrive , sStore, null 
	RunBCDEditEx "/Set {ramdiskoptions} ramdisksdipath \Boot\boot.sdi" , sStore, null 
	
	CreateRamDiskEntryEx = BCDObjectExistsEx(sStore, "{ramdiskoptions}")

End Function 


Const BDD_RAMDISK_GUID = "{d22e7e91-9ee7-46eb-89d7-c5859e4302f0}"

Function CreateNewBCDEntry( byRef sGUID, sDescription, sDrive, sPathToWim )
	CreateNewRamDiskEntry = CreateNewRamDiskEntryEx ( sStore, sGUID, sDescription, sDrive, sPathToWim )
End function

Function CreateNewBCDEntryEx ( sStore, byRef sGUID, sDescription, sDrive, sPathToWim )

	Dim bResult
	
	'
	' Normalize Defaults
	'
	If sDrive = "" then
		oLogging.CreateEntry "sDrive not defined for CreateNewBCDEntryEx. Defaulting to C:", LogTypeWarning
		sDrive = "C:"
	End if
	
	If sGUID = "" then
		sGUID = oStrings.GenerateRandomGUID
	End if

	'
	' Create Entry
	'	
	
	If sDescription = "" then
		bResult = RunBCDEditEx("/Create " & sGUID & " -d ""Custom Entry"" /application OSLOADER" , sStore, null)
	Else
		bResult = RunBCDEditEx("/Create " & sGUID & " -d """ & sDescription & """ /application OSLOADER" , sStore, null)
	End if
	
	'
	' Verify Store was created:
	'
	TEstAndLog BCDObjectExistsEx(sStore, sGUID ), "Create element: " & sGuid
	
	bResult = CreateRamDiskEntryEx( sStore, sDrive )
	
	'
	' Fix settings
	'
	RunBCDEditEx "/Set " & sGUID & " device ramdisk=[" & sDrive & "]" & sPathToWim & ",{ramdiskoptions}" , sStore, null 
	RunBCDEditEx "/Set " & sGUID & " osdevice ramdisk=[" & sDrive & "]" & sPathToWim & ",{ramdiskoptions}" , sStore, null
	If UCase(oEnvironment.Item("IsUEFI")) = "TRUE" then
		RunBCDEditEx "/Set " & sGUID & " path \windows\system32\boot\winload.efi" , sStore, null 
	Else
		RunBCDEditEx "/Set " & sGUID & " path \windows\system32\boot\winload.exe" , sStore, null 
	End if
	RunBCDEditEx "/Set " & sGUID & " systemroot \windows" , sStore, null 
	RunBCDEditEx "/Set " & sGUID & " detecthal yes" , sStore, null 
	RunBCDEditEx "/Set " & sGUID & " winpe yes" , sStore, null 
	' RunBCDEditEx "/Set " & sGUID & " inherit {bootloadersettings}"  , sStore, null 
	' RunBCDEditEx "/Set " & sGUID & " ems yes" , sStore, null 
	' RunBCDEditEx "/Set " & sGUID & " locale en-US" , sStore, null 
	
	CreateNewBCDEntryEx = SUCCESS 
	' programs should call BCDObjectExistsEx(sStore, sGUID ) to verify 

End Function

Function AdjustBCDDefaults ( sStore, sGUID )

	RunBCDEditEx "/timeout 0" , sStore, null 
	RunBCDEditEx "/displayorder " & sGUID & " /addfirst"  , sStore, null 
	RunBCDEditEx "/bootsequence " & sGUID               , sStore, null 
	RunBCDEditEx "/default "      & sGUID , sStore, null 
	
	AdjustBCDDefaults = SUCCESS

End function 


'//---------------------------------------------------------------------------
'//
'// BCDEdit functions
'//

Dim g_sBCDEdit

Function isBCDEditReady

	Dim iRetVal
	
	If isempty(g_sBCDEdit) then
		' Cache the BCDEdit.exe file so we don't have to search each time.
		iRetVal = oUtility.FindFile("bcdedit.exe", g_sBCDEdit)
		If iRetVal <> SUCCESS then
			oLogging.CreateEntry "Missing BCDEdit.exe.", LogTypeInfo
			g_sBCDEdit = ""
		End if
	End if

	isBCDEditReady = g_sBCDEdit <> ""

End Function

Function RunBCDEdit(sCommand )
	RunBCDEdit = RunBCDEditEx(sCommand, "", null )
End function

Function RunBCDEditEx(sCommand, sStore, ByRef sOutput )

	Dim iRetVal
	Dim sCmd
	Dim oExec
	Dim sLine
	
	If isempty(g_sBCDEdit) then
		' Cache the BCDEdit.exe file so we don't have to search each time.
		iRetVal = oUtility.FindFile("bcdedit.exe", g_sBCDEdit)
		If iRetVal <> SUCCESS then
			oLogging.CreateEntry "Missing BCDEdit.exe: " & sCommand, LogTypeInfo
			g_sBCDEdit = ""
		End if
	End if
	
	If g_sBCDEdit = "" then
		exit function
	ElseIf sStore = "" then
		sCmd = g_sBCDEdit & " " & sCommand
	Else
		sCmd = g_sBCDEdit & " /Store """ & sStore & """ " & sCommand
	End if
	
	oLogging.CreateEntry "Run Command: " & sCmd, LogTypeInfo
	set oExec = oShell.Exec(sCmd)
	
	do while oExec.Status = 0 or not oExec.StdOut.atEndOfStream
		If not oExec.StdOut.atEndOfStream then
			sLine = oExec.StdOut.ReadLine
			oLogging.CreateEntry "BCD> " & sLine, LogTypeInfo
			If not isnull(sOutput) then			
				sOutput = sOutput & vbNewLine & sLine
			End if
		End If
		If not oExec.StdErr.atEndOfStream then
			sLine = oExec.StdErr.ReadLine
			oLogging.CreateEntry "BCD> " & sLine, LogTypeError
		End If
	loop

	oLogging.CreateEntry "BCDEdit returned ErrorLevel = " & oExec.ExitCode, LogTypeInfo
	RunBCDEditEx = oExec.ExitCode

End function


'//---------------------------------------------------------------------------
'//
'// WMI Object
'//

Dim g_oBCDUtility
Function oBCDUtility
	If isempty(g_oBCDUtility) then
		set g_oBCDUtility = new BCDUtility
	End if
	set oBCDUtility = g_oBCDUtility
End function 


Class BCDUtility

	'
	' Private Global Variables
	'
	
	private g_sBCDEdit
	private g_oBCD
	private g_oBCDStore
	private g_sBCDStore
	private g_sCurrent
	
	
	'
	' Public Functions
	'
	
	Private Sub Class_Initialize
		g_sBCDStore = ""
	End Sub
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	'
	' Private Handles
	'

	Private Function oBCD
		' Get the base BCD object
		If isempty(g_oBCD) then
		
			on error resume next
			set g_oBCD = GetObject( "winmgmts:{impersonationlevel=Impersonate,(Backup,Restore)}!root/wmi:BcdStore" )
			TestAndFail not g_oBCD is nothing, 6601, "GetObject(... root/wmi:BCDStore)"	
			on error goto 0 
			
		End if
		set oBCD = g_oBCD
	End function 


	Private Function oBCDStore
		dim bResult
		If isempty(g_oBCDStore) then
			If g_sBCDStore = "" then
				oLogging.CreateEntry "BCD: Open Store: {system store}",LogTypeInfo
			Else
				oLogging.CreateEntry "BCD: Open Store: " & g_sBCDStore,LogTypeInfo
			End if 
			on error resume next
			bResult = oBCD.OpenStore( g_sBCDStore, g_oBCDStore )
			TestAndFail bResult, 6602, "BCD.OpenStore (" & g_sBCDStore & ")"
			on error goto 0 
		End if
		set oBCDStore = g_oBCDStore
	end function 


	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	'
	' Base Objects
	'

	Property Let sBCDStore( sNew )
		g_sBCDStore = sNew
		If not isempty(g_oBCDStore) then
			set g_oBCDStore = nothing
			g_oBCDStore = empty
			g_sCurrent = empty
		End if		
	End Property

	Property Get sBCDStore
		sBCDStore = g_sBCDStore
	End Property	


	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	'
	' Support Functions
	'

	Public Function TranslateGUIDToWellKnown ( sGUID ) 
		Dim oObject
	
		Select case sGUID
			case "{9dea862c-5cdd-4e70-acc1-f32b344d4795}"  TranslateGUIDToWellKnown = "{bootmgr}"
			case "{a5a30fa2-3d06-4e9f-b5f4-a01df9d1fcba}"  TranslateGUIDToWellKnown = "{fwbootmgr}"
			case "{b2721d73-1db4-4c62-bf78-c548a880142d}"  TranslateGUIDToWellKnown = "{memdiag}"
			case "{466f5a88-0af2-4f76-9038-095b170dc21c}"  TranslateGUIDToWellKnown = "{ntldr}"
			case "{fa926493-6f1c-4193-a414-58f0b2456d1e}"  TranslateGUIDToWellKnown = "{current}"
			case "{5189b25c-5558-4bf2-bca4-289b11bd29e2}"  TranslateGUIDToWellKnown = "{badmemory}"
			case "{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}"  TranslateGUIDToWellKnown = "{bootloadersettings}"
			case "{4636856e-540f-4170-a130-a84776f4c654}"  TranslateGUIDToWellKnown = "{dbgsettings}"
			case "{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}"  TranslateGUIDToWellKnown = "{emssettings}"
			case "{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}"  TranslateGUIDToWellKnown = "{globalsettings}"
			case "{1afa9c49-16ab-4a5c-901b-212802da9460}"  TranslateGUIDToWellKnown = "{resumeloadersettings}"
			case else   
				If sGUID = GetCurrentGUID then
			       TranslateGUIDToWellKnown = "{current}"
				Else
			       TranslateGUIDToWellKnown = sGUID
				End if
		End select
		
	End function

	Public Function TranslateWellKnownToGUID ( sGUID ) 
	
		Select case sGUID
			case "{bootmgr}"               TranslateWellKnownToGUID = "{9dea862c-5cdd-4e70-acc1-f32b344d4795}"
			case "{fwbootmgr}"             TranslateWellKnownToGUID = "{a5a30fa2-3d06-4e9f-b5f4-a01df9d1fcba}"
			case "{memdiag}"               TranslateWellKnownToGUID = "{b2721d73-1db4-4c62-bf78-c548a880142d}"
			case "{ntldr}"                 TranslateWellKnownToGUID = "{466f5a88-0af2-4f76-9038-095b170dc21c}"
			case "{current}"               TranslateWellKnownToGUID = "{fa926493-6f1c-4193-a414-58f0b2456d1e}"
			case "{badmemory}"             TranslateWellKnownToGUID = "{5189b25c-5558-4bf2-bca4-289b11bd29e2}"
			case "{bootloadersettings}"    TranslateWellKnownToGUID = "{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}"
			case "{dbgsettings}"           TranslateWellKnownToGUID = "{4636856e-540f-4170-a130-a84776f4c654}"
			case "{emssettings}"           TranslateWellKnownToGUID = "{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}"
			case "{globalsettings}"        TranslateWellKnownToGUID = "{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}"
			case "{resumeloadersettings}"  TranslateWellKnownToGUID = "{1afa9c49-16ab-4a5c-901b-212802da9460}"
			case else                      TranslateWellKnownToGUID = sGUID
		End select
		
	End function
	
	Public Function TranslateElementTypes ( iElementType )
	
		Select case iElementType
			case &h11000001     TranslateElementTypes = "ApplicationDevice"
			case &h12000002     TranslateElementTypes = "ApplicationPath"
			case &h12000004     TranslateElementTypes = "Description"
			case &h12000005     TranslateElementTypes = "PreferredLocale"
			case &h14000006     TranslateElementTypes = "InheritedObjects"
			case &h15000007     TranslateElementTypes = "TruncatePhysicalMemory"
			case &h14000008     TranslateElementTypes = "RecoverySequence"
			case &h16000009     TranslateElementTypes = "AutoRecoveryEnabled"
			case &h1700000a     TranslateElementTypes = "BadMemoryList"
			case &h1600000b     TranslateElementTypes = "AllowBadMemoryAccess"
			case &h1500000c     TranslateElementTypes = "FirstMegabytePolicy"
			case &h16000010     TranslateElementTypes = "DebuggerEnabled"
			case &h15000011     TranslateElementTypes = "DebuggerType"
			case &h15000012     TranslateElementTypes = "SerialDebuggerPortAddress"
			case &h15000013     TranslateElementTypes = "SerialDebuggerPort"
			case &h15000014     TranslateElementTypes = "SerialDebuggerBaudRate"
			case &h15000015     TranslateElementTypes = "1394DebuggerChannel"
			case &h12000016     TranslateElementTypes = "UsbDebuggerTargetName"
			case &h16000017     TranslateElementTypes = "DebuggerIgnoreUsermodeExceptions"
			case &h15000018     TranslateElementTypes = "DebuggerStartPolicy"
			case &h16000020     TranslateElementTypes = "EmsEnabled"
			case &h15000022     TranslateElementTypes = "EmsPort"
			case &h15000023     TranslateElementTypes = "EmsBaudRate"
			case &h12000030     TranslateElementTypes = "LoadOptionsString"
			case &h16000040     TranslateElementTypes = "DisplayAdvancedOptions"
			case &h16000041     TranslateElementTypes = "DisplayOptionsEdit"
			case &h16000046     TranslateElementTypes = "GraphicsModeDisabled"
			case &h15000047     TranslateElementTypes = "ConfigAccessPolicy"
			case &h16000049     TranslateElementTypes = "AllowPrereleaseSignatures"
			
			case &h24000001     TranslateElementTypes = "DisplayOrder"
			case &h24000002     TranslateElementTypes = "BootSequence"
			case &h23000003     TranslateElementTypes = "DefaultObject"
			case &h25000004     TranslateElementTypes = "Timeout"
			case &h26000005     TranslateElementTypes = "AttemptResume"
			case &h23000006     TranslateElementTypes = "ResumeObject"
			case &h24000010     TranslateElementTypes = "ToolsDisplayOrder"
			case &h21000022     TranslateElementTypes = "BcdDevice"
			case &h22000023     TranslateElementTypes = "BcdFilePath"

			case &h35000001     TranslateElementTypes = "RamdiskImageOffset"
			case &h35000002     TranslateElementTypes = "TftpClientPort"
			case &h31000003     TranslateElementTypes = "SdiDevice"
			case &h32000004     TranslateElementTypes = "SdiPath"
			case &h35000005     TranslateElementTypes = "RamdiskImageLength"
					
			case &h25000001     TranslateElementTypes = "PassCount"
			case &h25000003     TranslateElementTypes = "FailureCount"

			case &h21000001     TranslateElementTypes = "OSDevice"
			case &h22000002     TranslateElementTypes = "SystemRoot"
			case &h23000003     TranslateElementTypes = "AssociatedResumeObject"
			case &h26000010     TranslateElementTypes = "DetectKernelAndHal"
			case &h22000011     TranslateElementTypes = "KernelPath"
			case &h22000012     TranslateElementTypes = "HalPath"
			case &h22000013     TranslateElementTypes = "DbgTransportPath"
			case &h25000020     TranslateElementTypes = "NxPolicy"
			case &h25000021     TranslateElementTypes = "PAEPolicy"
			case &h26000022     TranslateElementTypes = "WinPEMode"
			case &h26000024     TranslateElementTypes = "DisableCrashAutoReboot"
			case &h26000025     TranslateElementTypes = "UseLastGoodSettings"
			case &h26000027     TranslateElementTypes = "AllowPrereleaseSignatures"
			case &h26000030     TranslateElementTypes = "NoLowMemory"
			case &h25000031     TranslateElementTypes = "RemoveMemory"
			case &h25000032     TranslateElementTypes = "IncreaseUserVa"
			case &h26000040     TranslateElementTypes = "UseVgaDriver"
			case &h26000041     TranslateElementTypes = "DisableBootDisplay"
			case &h26000042     TranslateElementTypes = "DisableVesaBios"
			case &h25000050     TranslateElementTypes = "ClusterModeAddressing"
			case &h26000051     TranslateElementTypes = "UsePhysicalDestination"
			case &h25000052     TranslateElementTypes = "RestrictApicCluster"
			case &h26000060     TranslateElementTypes = "UseBootProcessorOnly"
			case &h25000061     TranslateElementTypes = "NumberOfProcessors"
			case &h26000062     TranslateElementTypes = "ForceMaximumProcessors"
			case &h25000063     TranslateElementTypes = "ProcessorConfigurationFlags"
			case &h26000070     TranslateElementTypes = "UseFirmwarePciSettings"
			case &h26000071     TranslateElementTypes = "MsiPolicy"
			case &h25000080     TranslateElementTypes = "SafeBoot"
			case &h26000081     TranslateElementTypes = "SafeBootAlternateShell"
			case &h26000090     TranslateElementTypes = "BootLogInitialization"
			case &h26000091     TranslateElementTypes = "VerboseObjectLoadMode"
			case &h260000a0     TranslateElementTypes = "KernelDebuggerEnabled"
			case &h260000a1     TranslateElementTypes = "DebuggerHalBreakpoint"
			case &h260000b0     TranslateElementTypes = "EmsEnabled"
			case &h250000c1     TranslateElementTypes = "DriverLoadFailurePolicy"
			case &h250000E0     TranslateElementTypes = "BootStatusPolicy"

			case else
				TranslateElementTypes = hex(iElementType)
		End Select
	
	End function
	
	Public Function TestObject(sObjectGUID)
		Dim oObject
		on error resume next
		TestOBject = FALSE
		TestObject = oBCDStore.OpenObject(TranslateWellKnownToGUID(sObjectGUID),oObject)
		on error goto 0
	End function 

	Public Function OpenObject(sObjectGUID)
		TestAndLog oBCDStore.OpenObject(TranslateWellKnownToGUID(sObjectGUID),OpenObject), "OpenObject(" & sObjectGUID & ")"
	End function 

	Public Function GetSystemPartition
		TestAndLog oBCDStore.GetSystemPartition(GetSystemPartition), "GetSystemPartition = " & GetSystemPartition
	End function

	Public Function GetSystemDisk
		TestAndLog oBCDStore.GetSystemDisk(GetSystemDisk), "GetSystemDisk = " & GetSystemDisk
	End function
	
	Public Function GetElement( oBCDStoreObject, oItem )
		Dim bResult
		on error resume next
		bResult = oBCDStoreObject.GetElement( oItem, GetElement )
		If Err <> 0 or not bResult then
			CreateEntry oEnvironment.Substitute( "FAILURE (Err): " & FormatError(Err) & ": " & "GetElement = 0x" & hex(oItem) ), LogTypeWarning
		End if
		on error goto 0
		
	End function
	

	Public Function EnumerateObjects ( sEnumType )
		Dim bResult
		on error resume next
		bResult = oBCDStore.EnumerateObjects( sEnumType, EnumerateObjects )
		TestAndLog bResult, "Enumerate Objects"
		on error goto 0 
		If not bResult or isempty(bResult) then
			bResult = array()
		End if 
	End function 
	
	Public Function EnumerateElementTypes( oStoreObject ) 
		Dim bResult
		on error resume next
		bResult = oStoreObject.EnumerateElementTypes( EnumerateElementTypes )
		TestAndLog bResult, "EnumerateElementTypes"
		on error goto 0 
		If not bResult or isempty(bResult) then
			bResult = array()
		End if 
	End function 
	
	Public Function EnumerateElements( oStoreObject ) 
		Dim bResult
		on error resume next
		bResult = oStoreObject.EnumerateElements( EnumerateElements )
		TestAndLog bResult, "EnumerateElements"
		on error goto 0 
		If not bResult or isempty(bResult) then
			bResult = array()
		End if 
	End function 

	Function GetCurrentGUID	
		Dim oObject
		If isempty(g_sCurrent) then
			set oObject = OpenObject("{9dea862c-5cdd-4e70-acc1-f32b344d4795}")
			TestAndLog not oObject is nothing, "OpenObject({9dea862c-5cdd-4e70-acc1-f32b344d4795})"
			g_sCurrent = GetElement( oObject, &h23000003 ).ID
		End if
		GetCurrentGUID = g_sCurrent		
	End function
	

	Private Sub DumpBCDObjectStore ( oBCDStoreObject ) 
		Dim oItem
		for each oItem in EnumerateElementTypes( oBCDStoreObject )
			oLogging.CreateEntry TranslateElementTypes(oItem) & " = " & DumpBCDElement ( GetElement(oBCDStoreObject,oItem)  ), LogTypeInfo			
		next
	End Sub 
	
	Private Function DumpBCDElement( oElement )
	
		Dim oSubProp
		Dim oSubSubProp
		Dim Value
				
		for each oSubProp in oElement.Properties_
			Select case uCase(oSubProp.Name)
				case "TYPE", "OBJECTID", "STOREFILEPATH"
				
				case "STRING"    
					DumpBCDElement = oELement.String
					exit for
				case "ID"        
					DumpBCDElement = oELement.ID
					exit for
				case "BOOLEAN"   
					DumpBCDElement = oELement.Boolean
					exit for
				case "INTEGER"   
					DumpBCDElement = oELement.Integer
					exit for
				case "IDS"       
					DumpBCDElement = ostrings.ForceAsString(oELement.IDs)
					exit for
				case "INTEGERS"  
					DumpBCDElement = ostrings.ForceAsString(oELement.INtegers) 
					exit for
				case "DEVICE"    
					for each oSubSubProp in oSubProp.Value.Properties_
						Select Case uCase(oSubSubProp.Name)
						
							Case "ADDITIONALOPTIONS", "DEVICETYPE"
							Case else
								DumpBCDElement = ostrings.ForceAsString(oSubSubProp.value)
						End Select
					next
				case else

					oLogging.CreateEntry vbTab & oSubProp.Name & " = " & Value, LogTypeInfo	
			End Select
		next
		
		DumpBCDElement = TranslateGUIDToWellKnown(DumpBCDElement)
	
	End function 

	Public Function DumpBCDStore
	
		Dim oItem
		
		DumpBCDObjectStore OpenObject(TranslateWellKnownToGUID("{bootmgr}"))
		for each oItem in EnumerateObjects( &h10300006 )
			DumpBCDObjectStore oItem 
		next 
		for each oItem in EnumerateObjects( &h10200003 )
			DumpBCDObjectStore oItem 
		next 
	

	End function 

End class


