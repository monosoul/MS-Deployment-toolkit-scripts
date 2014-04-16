'#########################################################################################
'#   MICROSOFT LEGAL STATEMENT FOR SAMPLE SCRIPTS/CODE
'#########################################################################################
'#   This Sample Code is provided for the purpose of illustration only and is not 
'#   intended to be used in a production environment.
'#
'#   THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY 
'#   OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
'#   WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'#
'#   We grant You a nonexclusive, royalty-free right to use and modify the Sample Code 
'#   and to reproduce and distribute the object code form of the Sample Code, provided 
'#   that You agree: 
'#   (i)      to not use Our name, logo, or trademarks to market Your software product 
'#            in which the Sample Code is embedded; 
'#   (ii)     to include a valid copyright notice on Your software product in which 
'#            the Sample Code is embedded; and 
'#   (iii)    to indemnify, hold harmless, and defend Us and Our suppliers from and 
'#            against any claims or lawsuits, including attorneys’ fees, that arise 
'#            or result from the use or distribution of the Sample Code.
'#########################################################################################
' //***************************************************************************
' // ***** Script Header *****
' //
' // Solution:  Solution Accelerator - Microsoft Deployment Toolkit
' // File:      GetCurrentTimeZoneExit.vbs
' //
' // Purpose:   User Exit script to query current time zone information.
' //
' // Usage:     Modify CustomSettings.ini similar to this:
' //              [Settings]
' //              Priority=TimeZone, TestLegacyOS
' //              Properties=TimeZoneStandardName, TimeZoneCaption, IsLegacyOS
' //              
' //              [TimeZone]
' //              UserExit=GetCurrentTimeZoneExit.vbs
' //              TimeZoneStandardName=#GetCurrentTimeZoneWmiProperty("StandardName")#
' //              TimeZoneCaption=#GetCurrentTimeZoneWmiProperty("Caption")#
' //              TimeZoneName=#GetCurrentTimeZoneRegistryKeyName#
' //              IsLegacyOS=#ConvertBooleanToString(%OSCurrentBuild% < 5200)#
' //              
' //              [TestLegacyOS]
' //              SubSection=IsLegacyOS-%IsLegacyOS%
' //              
' //              [IsLegacyOS-True]
' //              TimeZone=#GetCurrentTimeZoneLegacyIndex#
' //
' // Customer Build Version:      1.0.0
' // Customer Script Version:     1.0.0
' //
' // Customer History:
' // 1.0.0   MDM   10/22/2008  Created script.
' //
' // ***** End Header *****
' //***************************************************************************


Function UserExit(sType, sWhen, sDetail, bSkip)

    oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs started: " & sType & " " & sWhen & " " & sDetail, LogTypeInfo

    UserExit = Success

End Function


Function GetCurrentTimeZoneWmiProperty(sProperty)

    On Error Resume Next
    
    GetCurrentTimeZoneWmiProperty = ""

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Err.Clear
    Set colTimeZone = objWMI.ExecQuery("Select * from Win32_TimeZone",,48)
    For Each objTimeZone in colTimeZone
        Execute "GetCurrentTimeZoneWmiProperty = objTimeZone." & sProperty
    Next

    If GetCurrentTimeZoneWmiProperty <> "" Then
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneWmiProperty: Output value: " & sProperty & " = " & GetCurrentTimeZoneWmiProperty, LogTypeInfo
    Else
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneWmiProperty: Error occured", LogTypeError
    End If

End Function


Function GetCurrentTimeZoneRegistryKeyName()

    GetCurrentTimeZoneRegistryKeyName = ""
    
    Const HKEY_CLASSES_ROOT   = &H80000000
    Const HKEY_CURRENT_USER   = &H80000001
    Const HKEY_LOCAL_MACHINE  = &H80000002
    Const HKEY_USERS          = &H80000003
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_DYN_DATA       = &H80000006

    Const REG_SZ        = 1
    Const REG_EXPAND_SZ = 2
    Const REG_BINARY    = 3
    Const REG_DWORD     = 4
    Const REG_MULTI_SZ  = 7
    
    Const TIME_ZONE_KEY_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
    Const TIME_ZONE_KEY_9X = "SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones"
    
    sRegValueName = "Std"
    sPropertyName = "StandardName"
    sRegValueName = "Display"
    sPropertyName = "Caption"

    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colTimeZone = objWMI.ExecQuery("Select * from Win32_TimeZone",,48)
    For Each objTimeZone in colTimeZone
        Execute "sPropertyValue = objTimeZone." & sPropertyName
    Next
    
    lRC1 = objRegistry.EnumKey(HKEY_LOCAL_MACHINE, TIME_ZONE_KEY_NT, arrSubKeys)

	On Error Resume Next
	Err.Clear
	IsSubscriptOutOfRange = arrSubKeys(0)

	If (lRC = 0) AND (Err.Number = 0) AND (NOT IsNULL(arrSubKeys)) Then
		For i = LBound(arrSubKeys) To UBound(arrSubKeys)
            'WScript.Echo arrSubKeys(i)
            lRC = objRegistry.GetStringValue(HKEY_LOCAL_MACHINE, TIME_ZONE_KEY_NT & "\" & arrSubKeys(i), sRegValueName, sValue)
            If sValue = sPropertyValue Then
                'WScript.Echo sRegValueName & " = " & sValue
                GetCurrentTimeZoneRegistryKeyName = arrSubKeys(i)
                Exit For
            End If
		Next
	End If

    If GetCurrentTimeZoneRegistryKeyName <> "" Then
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneRegistryKeyName: Output value: " & GetCurrentTimeZoneRegistryKeyName, LogTypeInfo
    Else
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneRegistryKeyName: Error occured", LogTypeError
    End If

End Function


Function GetCurrentTimeZoneLegacyIndex()

    GetCurrentTimeZoneLegacyIndex = ""
    
    Const HKEY_CLASSES_ROOT   = &H80000000
    Const HKEY_CURRENT_USER   = &H80000001
    Const HKEY_LOCAL_MACHINE  = &H80000002
    Const HKEY_USERS          = &H80000003
    Const HKEY_CURRENT_CONFIG = &H80000005
    Const HKEY_DYN_DATA       = &H80000006

    Const REG_SZ        = 1
    Const REG_EXPAND_SZ = 2
    Const REG_BINARY    = 3
    Const REG_DWORD     = 4
    Const REG_MULTI_SZ  = 7
    
    Const TIME_ZONE_KEY_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
    Const TIME_ZONE_KEY_9X = "SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones"
    
    sRegValueName = "Std"
    sPropertyName = "StandardName"
    'sRegValueName = "Display"
    'sPropertyName = "Caption"

    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colTimeZone = objWMI.ExecQuery("Select * from Win32_TimeZone",,48)
    For Each objTimeZone in colTimeZone
        Execute "sPropertyValue = objTimeZone." & sPropertyName
    Next
    
    lRC1 = objRegistry.EnumKey(HKEY_LOCAL_MACHINE, TIME_ZONE_KEY_NT, arrSubKeys)

	On Error Resume Next
	Err.Clear
	IsSubscriptOutOfRange = arrSubKeys(0)

	If (lRC = 0) AND (Err.Number = 0) AND (NOT IsNULL(arrSubKeys)) Then
		For i = LBound(arrSubKeys) To UBound(arrSubKeys)
            'WScript.Echo arrSubKeys(i)
            lRC = objRegistry.GetStringValue(HKEY_LOCAL_MACHINE, TIME_ZONE_KEY_NT & "\" & arrSubKeys(i), sRegValueName, sValue)
            If sValue = sPropertyValue Then
                'WScript.Echo sRegValueName & " = " & sValue

                lRC2 = objRegistry.GetDWORDValue(HKEY_LOCAL_MACHINE, TIME_ZONE_KEY_NT & "\" & arrSubKeys(i), "Index", lIndexValue)
                GetCurrentTimeZoneLegacyIndex = lIndexValue
                'WScript.Echo "Index = " & lIndexValue

                Exit For
            End If
		Next
	End If

    If GetCurrentTimeZoneLegacyIndex <> "" Then
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneLegacyIndex: Output value: " & GetCurrentTimeZoneLegacyIndex, LogTypeInfo
    Else
        oLogging.CreateEntry "USEREXIT:GetCurrentTimeZoneExit.vbs|GetCurrentTimeZoneLegacyIndex: Error occured", LogTypeError
    End If

End Function

