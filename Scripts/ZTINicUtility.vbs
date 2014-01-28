' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTINicUtility.wsf
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Utility functions For NIC COnfiguration
' // 
' // Usage:     <script language="VBScript" src="ZTINicUtility.vbs"/>
' // 
' // ***************************************************************************


'    Properties In Use:
'
'       OSDDiskIndex
'       OSDDiskType
'       OSDDiskType
'       Dim OSDAdapterCount
'
'        For each Adapter: OSDAdapter(N)
'
'            Dim Index
'            Dim Name
'            Dim MacAddress
'
'            Dim EnableDHCP ' ALLWAYS FALSE, otherwise assume DHCP
'
'            Dim IpAddressList()
'            Dim SubnetMask()
'            Dim Gateways()
'            Dim GatewayCostMetric
'
'            Dim DNSServerList()
'            Dim DNSSuffix
'            Dim EnableDNSRegistration ' T/F
'            Dim EnableFullDNSRegistration ' T/F
'
'            Dim EnableWINS ' T/F
'            Dim WINSServerList()
'            Dim EnableLMHOSTS ' T/F
'
'            Dim TCPIPNetBiosOptions ' T/F
'
'            Dim EnableTCPIPFiltering
'            Dim TCPFilterPortList()
'            Dim UDPFilterPortList()
'            Dim IPProtocolFilterList()
'

Option Explicit

Dim g_ObjNetworkAdapters


'
'  Given a Network Adatper from Win32_NetworkAdatperConfiguration, and an integer,
'     this rountine will extract the relevant settings from the NIC and save them to the environment.
'
Sub SaveNetworkAdapterSettings ( oAdapter, AdapterIndex, SaveWithMACAddress )

	Dim IPAddress, SubnetMask
	Dim i

	oLogging.CreateEntry "Save Network Adapter(" & AdapterIndex & ") MAC = " & SaveWithMACAddress & " " & oAdapter.Description , LogTypeInfo

	SaveAdapterSetting AdapterIndex, "EnableDHCP",				"False"

	SaveAdapterSetting AdapterIndex, "Name",				GetNicName ( oAdapter )

	If SaveWithMACAddress then
		SaveAdapterSetting AdapterIndex, "MacAddress",			oAdapter.MACAddress
	Else
		SaveAdapterSetting AdapterIndex, "MacAddress",			""
	End if

	IPAddress = oAdapter.IPAddress
	SubnetMask = oAdapter.IPSubnet

	' Filter out IPv6 Addresses, Link Local address common in Windows Vista.
	i = LBound(IPAddress)
	While i <= ubound(IPAddress)
		If isIPV6Address(IPAddress(i)) then
			RemoveElementFromArray IPAddress,i
			RemoveElementFromArray SubnetMask,i
		Else
			i = i + 1
		End if
	Wend

	' Description is not part of the spec, but we save it anyways For reference.
	SaveAdapterSetting AdapterIndex, "Description",				oAdapter.Description
	SaveAdapterSetting AdapterIndex, "IpAddressList",			IPAddress
	SaveAdapterSetting AdapterIndex, "SubnetMask",				SubnetMask
	SaveAdapterSetting AdapterIndex, "Gateways",				oAdapter.DefaultIPGateway
	' Bug: 7879 - WMI does not support "Automatic" Gateway Cost Metric. Disable.
	' SaveAdapterSetting AdapterIndex, "GatewayCostMetric",			oAdapter.GatewayCostMetric
	SaveAdapterSetting AdapterIndex, "DNSServerList",			oAdapter.DNSServerSearchOrder
	SaveAdapterSetting AdapterIndex, "DNSSuffix",				oAdapter.DNSDomain
	SaveAdapterSetting AdapterIndex, "EnableDNSRegistration",		oAdapter.DomainDNSRegistrationEnabled
	SaveAdapterSetting AdapterIndex, "EnableFullDNSRegistration",		oAdapter.FullDNSRegistrationEnabled
	SaveAdapterSetting AdapterIndex, "EnableLMHOSTS",			oAdapter.WINSEnableLMHostsLookup
	SaveAdapterSetting AdapterIndex, "TCPIPNetBiosOptions",			oAdapter.TcpipNetbiosOptions

	If not isnull(oAdapter.WINSPrimaryServer) and not isnull(oAdapter.WINSSecondaryServer) then
		SaveAdapterSetting AdapterIndex, "EnableWINS",			"True"
		SaveAdapterSetting AdapterIndex, "WINSServerList",		oAdapter.WINSPrimaryServer & "," & oAdapter.WINSSecondaryServer
	ElseIf not isnull(oAdapter.WINSPrimaryServer) then
		SaveAdapterSetting AdapterIndex, "EnableWINS",			"True"
		SaveAdapterSetting AdapterIndex, "WINSServerList",		oAdapter.WINSPrimaryServer
	ElseIf not isnull(oAdapter.WINSSecondaryServer) then
		SaveAdapterSetting AdapterIndex, "EnableWINS",			"True"
		SaveAdapterSetting AdapterIndex, "WINSServerList",		oAdapter.WINSSecondaryServer
	Else
		SaveAdapterSetting AdapterIndex, "EnableWINS",			"False"
	End if

	If oAdapter.IPFilterSecurityEnabled then
		SaveAdapterSetting AdapterIndex, "EnableIPProtocolFiltering",	"True"

		SaveAdapterSetting AdapterIndex, "TCPFilterPortList",		oAdapter.IPSecPermitTCPPorts
		SaveAdapterSetting AdapterIndex, "UDPFilterPortList",		oAdapter.IPSecPermitUDPPorts
		SaveAdapterSetting AdapterIndex, "IPProtocolFilterList",	oAdapter.IPSecPermitIPProtocols
	Else
		SaveAdapterSetting AdapterIndex, "EnableIPProtocolFiltering",	"False"
		SaveAdapterSetting AdapterIndex, "TCPFilterPortList",		""
		SaveAdapterSetting AdapterIndex, "UDPFilterPortList",		""
		SaveAdapterSetting AdapterIndex, "IPProtocolFilterList",	""
	End if

End Sub


'//----------------------------------------------------------------------------

Function LoadNetworkAdapterSettings ( oAdapter, AdapterIndex )

	' Returns

	oLogging.CreateEntry "Load Network Adapter(" & AdapterIndex & ") = " & oAdapter.Description , LogTypeInfo

	Dim i
	Dim Item
	Dim iResult
	Dim sErrorString
	Dim sCostMetric

	' -----------------------------------------------------
	' Release any acquired DHCP address
	If oAdapter.DHCPEnabled then
		If oAdapter.DHCPServer <> NULL then
			If oAdapter.DHCPServer <> "255.255.255.255" then
				oAdapter.ReleaseDHCPLeaseAll()
			End if
		End if
	End if

	' -----------------------------------------------------
	' IP Addresses and associated SubNet Masks (REQUIRED)
	Dim IPAddress, SubnetMask

	If LoadAdapterSetting( AdapterIndex, "IpAddressList" ) <> "" and LoadAdapterSetting( AdapterIndex, "SubnetMask" ) <> ""  then

		IPAddress = LoadAdapterSettingAsArray( AdapterIndex, "IpAddressList" )
		SubnetMask = LoadAdapterSettingAsArray( AdapterIndex, "SubnetMask" )

		' Filter out IPv6 Addresses
		For i = lbound(IPAddress) to ubound(IPAddress)
			If isIPV6Address(IPAddress(i)) then
				RemoveElementFromArray IPAddress,i
				RemoveElementFromArray SubnetMask,i
			End if
		next

		' Set Static Addresses
		oLogging.CreateEntry "Action: oAdapter.EnableStatic(IPAddress,SubnetMask)", LogTypeInfo
		iResult = oAdapter.EnableStatic(IPAddress,SubnetMask)
		CheckForErrorsEx iResult, sErrorString,"EnableStatic(IPAddress,SubnetMask)", false

	End if

	' -----------------------------------------------------
	' Gateway Address and associated Cost Metrics (REQUIRED)

	oLogging.CreateEntry "Action: oAdapter.SetGateways(Gateway, GatewayMetric)", LogTypeInfo
	sCostMetric = LoadAdapterSetting( AdapterIndex, "GatewayCostMetric" )
	If instr(1,sCostMetric,"automatic",vbTextCompare) <> 0 then
		' There is no way to explicitly set the gateway cost metric to "automatic" within WMI.
		' Setting to NULL should produce the nearest results.
		sCostMetric= ""
	End if
	If LoadAdapterSetting( AdapterIndex, "Gateways" ) <> "" and sCostMetric <> "" then
		iResult = oAdapter.SetGateways( LoadAdapterSettingAsArray( AdapterIndex, "Gateways" ), LoadAdapterSettingAsIntArray( AdapterIndex, "GatewayCostMetric" ) )
		CheckForErrorsEx iResult, sErrorString,"SetGateways(Gateway, GatewayMetric)", false
	ElseIf LoadAdapterSetting( AdapterIndex, "Gateways" ) <> "" then
		iResult = oAdapter.SetGateways( LoadAdapterSettingAsArray( AdapterIndex, "Gateways" ), NULL )
		CheckForErrorsEx iResult, sErrorString,"SetGateways(Gateway, GatewayMetric)", false
	End if

	' -----------------------------------------------------

	oLogging.CreateEntry "Action: oAdapter.SetTcpipNetbios", LogTypeInfo
	Item = ucase(trim(LoadAdapterSetting( AdapterIndex, "TCPIPNetBiosOptions" )))
	If Item <> "" then
		Select case(Item)
			case "FALSE", "DHCP"
				iResult = oAdapter.SetTcpipNetbios(0)
			case "TRUE"
				iResult = oAdapter.SetTcpipNetbios(1)
			case "DISABLE"
				iResult = oAdapter.SetTcpipNetbios(2)
			case "0","1","2"
				iResult = oAdapter.SetTcpipNetbios(Item)
			case else
				oLogging.CreateEntry "OSDAdapter" & AdapterIndex & "TCPIPNetBiosOptions value unknown: [" & Item & "]", LogTypeInfo
		End Select
		CheckForErrors iResult, sErrorString,"SetTcpipNetbios"
	End if

	' -----------------------------------------------------
	oLogging.CreateEntry "Action: oAdapter.SetWINSServer(1,2)", LogTypeInfo
	Item = LoadAdapterSettingAsArray( AdapterIndex, "WINSServerList" )
	If ubound(Item) >= 1 then
		If Item(0) <> "" and Item(1) <> "" then
			iResult = oAdapter.SetWINSServer( Item(0), Item(1))
			CheckForErrors iResult, sErrorString,"SetWINSServer"
		End if
	ElseIf ubound(Item) >= 0 then
		If Item(0) <> "" then
			iResult = oAdapter.SetWINSServer( Item(0), "")
			CheckForErrors iResult, sErrorString,"SetWINSServer"
		End if
	End if

	' -----------------------------------------------------

	oLogging.CreateEntry "Action: oAdapter.SetDNSServerSearchOrder()", LogTypeInfo
	If LoadAdapterSetting ( AdapterIndex, "DNSServerList" ) <> "" then
		iResult = oAdapter.SetDNSServerSearchOrder ( LoadAdapterSettingAsArray ( AdapterIndex, "DNSServerList" ) )
		CheckForErrors iResult, sErrorString,"SetDNSServerSearchOrder"
	End if

	' -----------------------------------------------------

	oLogging.CreateEntry "Action: oAdapter.SetDNSDomain()", LogTypeInfo
	If LoadAdapterSetting ( AdapterIndex, "DNSSuffix" ) <> "" then
		iResult = oAdapter.SetDNSDomain ( LoadAdapterSetting ( AdapterIndex, "DNSSuffix" ) )
		CheckForErrors iResult, sErrorString,"SetDNSDomain()"
	End if

	' -----------------------------------------------------
	Dim EnableDNS
	Dim EnableFullDNS

	oLogging.CreateEntry "Action: oAdapter.SetDynamicDNSRegistration(...)", LogTypeInfo
	EnableDNS = oAdapter.DomainDNSRegistrationEnabled
	EnableFullDNS = oAdapter.FullDNSRegistrationEnabled

	If LoadAdapterSetting( AdapterIndex, "EnableDNSRegistration" ) <> "" then
		EnableDNS = cBool(LoadAdapterSetting( AdapterIndex, "EnableDNSRegistration" ))
	End if
	If LoadAdapterSetting( AdapterIndex, "EnableFullDNSRegistration" ) <> "" then
		EnableFullDNS = cBool(LoadAdapterSetting( AdapterIndex, "EnableFullDNSRegistration" ))
	End if

	If LoadAdapterSetting( AdapterIndex, "EnableDNSRegistration" ) <> "" or LoadAdapterSetting( AdapterIndex, "EnableFullDNSRegistration" ) <> "" then
		iResult = oAdapter.SetDynamicDNSRegistration(EnableFullDNS, EnableDNS)
		CheckForErrors iResult, sErrorString,"SetDynamicDNSRegistration"
	End if

	' -----------------------------------------------------

	oLogging.CreateEntry "Action: oAdapter.EnableIPFilterSec(...) " , LogTypeInfo
	If LoadAdapterSetting( AdapterIndex, "EnableIPProtocolFiltering" ) <> "" then

		If oAdapter.IPFilterSecurityEnabled <> cBOOL(LoadAdapterSetting( AdapterIndex, "EnableIPProtocolFiltering" )) then
			oLogging.CreateEntry oAdapter.IPFilterSecurityEnabled & " <> " & cBOOL(LoadAdapterSetting( AdapterIndex, "EnableIPProtocolFiltering" )) , LogTypeInfo
			' on error resume next
			iResult = oAdapter.EnableIPFilterSec(not oAdapter.IPFilterSecurityEnabled)
			' on error goto 0
			CheckForErrors iResult, sErrorString,"EnableIPFilterSec"
		End if
		If oAdapter.IPFilterSecurityEnabled then
			iResult = oAdapter.EnableIPSec(  LoadAdapterSettingAsArray( AdapterIndex, "TCPFilterPortList" ), _
			LoadAdapterSettingAsArray( AdapterIndex, "UDPFilterPortList" ), _
			LoadAdapterSettingAsArray( AdapterIndex, "IPProtocolFilterList" ) )
			CheckForErrors iResult, sErrorString,"EnableIPSec"
		End if

	End if

	' -----------------------------------------------------

	oLogging.CreateEntry "Action: oAdapter.Name(...) " , LogTypeInfo
	If LoadAdapterSetting( AdapterIndex, "Name" ) <> "" then
		RenameNic GetNicName( oAdapter ) , LoadAdapterSetting( AdapterIndex, "Name" )
	End if

	LoadNetworkAdapterSettings = sErrorString

End Function


'//----------------------------------------------------------------------------

const NETWORK_CONNECTIONS = &H31&


Function RenameNic ( sOldName , sNewName )

	Dim IRetVal
	Dim sCmd

	oLogging.CreateEntry vbTab & "Change [" & sOldName & "] to [" & sNewName & "]", LogTypeInfo
	
	sCmd = "cmd.exe /c netsh.exe interface set interface name=""" & sOldName & """ newname=""" & sNewName & """ 1>> %LogPath%\NetSh.log 2>>&1"
	sCmd = oEnvironment.Substitute(sCmd)

	If uCase(sOldName) <> uCase(sNewName) and oEnv("SystemDrive") <> "X:" then
		iRetVal = oUtility.RunWithHeartbeat ( sCmd )
		TestAndLog iRetVal, "Change: " & sCmd
	End if

End function


Function GetNicName ( oNicCfg )

	' Given an instance of the WMI Win32_NetworkAdapterConfiguration class, will return it's Net Connection ID

	Dim oNic
	For Each oNic in objWMI.ExecQUery("ASSOCIATORS OF {" & oNicCfg.Path_ & "} WHERE AssocClass = Win32_NetworkAdapterSetting ")
		GetNicName = oNic.NetConnectionID
		Exit for
	Next

End function

Sub CheckForErrors( iResult, byref ErrorString, sDescription )
	CheckForErrorsEx iResult, ErrorString, sDescription, true
End sub

Sub CheckForErrorsEx( iResult, byref ErrorString, sDescription, bSilentMessages )

	If iResult = 0 then
		oLogging.CreateEntry "WMI Function: Adapter." & sDescription & " Success!", LogTypeInfo
	ElseIf iResult = 96 then  ' DNS notification errors are OK
		oLogging.CreateEntry "WMI Function: Adapter." & sDescription & " IGNORING: " & iResult & "   " & GetWMINetworkErrorMessage( iResult ), LogTypeInfo
	Else
		If bSilentMessages then
			oLogging.CreateEntry "WMI Function: Adapter." & sDescription & " FAILURE: " & iResult & "   " & GetWMINetworkErrorMessage( iResult ), LogTypeInfo
		Else
			oLogging.CreateEntry "WMI Function: Adapter." & sDescription & " FAILURE: " & iResult & "   " & GetWMINetworkErrorMessage( iResult ), LogTypeWarning
		ENd if

		If not isempty(ErrorString) then
			ErrorString = ErrorString & vbNewLine
		End if

		ErrorString = ErrorString & sDescription & " Failure: " & GetWMINetworkErrorMessage( iResult )

	End if

End sub


Function LoadAdapterSetting( sAdapterIndex, sName )

	LoadAdapterSetting = oEnvironment.Item( "OSDAdapter" & sAdapterIndex & sName )

End Function


Function LoadAdapterSettingAsArray( sAdapterIndex, sName )

	LoadAdapterSettingAsArray = LoadAdapterSetting( sAdapterIndex , sName )
	If LoadAdapterSettingAsArray = "" then
		LoadAdapterSettingAsArray = array ("")
	Else
		LoadAdapterSettingAsArray = split ( LoadAdapterSettingAsArray, ",")
	End if

End Function


Function LoadAdapterSettingAsIntArray( sAdapterIndex, sName )
	Dim i
	Dim StringArray
	Dim IntArray

	StringArray = LoadAdapterSettingAsArray( sAdapterIndex, sName )

	ReDim IntArray(ubound(StringArray))
	For i = 0 to uBound(StringArray)
		IntArray(i) = cint(StringArray(i))
	Next
	LoadAdapterSettingAsIntArray = IntArray

End Function


Sub CleanNetworkSettings(Index)
	Dim oItem
	Const IPVariables = "Index Name MacAddress IpAddressList SubnetMask Gateways GatewayCostMetric DNSServerList DNSSuffix EnableDNSRegistration EnableFullDNSRegistration EnableWINS WINSServerList EnableLMHOSTS TCPIPNetBiosOptions EnableTCPIPFiltering TCPFilterPortList UDPFilterPortList IPProtocolFilterList"

	For each oItem in split(IPVariables," ")
		SaveAdapterSetting  Index, oItem, ""
	Next

End sub


Sub SaveAdapterSetting( sAdapterIndex, sName, Value )

	If oEnvironment.Item( "OSDAdapter" & sAdapterIndex & sName ) = "" then
		If isArray(Value) then
			If ubound(Value) <= 1 and Value(0) = "" then
				Exit sub
			End if
		ElseIf IsNull(Value) then
			Exit sub
		ElseIf Value = "" then
			Exit sub
		End if
	End if

	If IsArray(Value) then

		' Combine Arrays into comma delimited strings
		oEnvironment.Item( "OSDAdapter" & sAdapterIndex & sName ) = join( Value, "," )

	ElseIf IsNull(Value) then
		oEnvironment.Item( "OSDAdapter" & sAdapterIndex & sName ) = ""
	Else
		oEnvironment.Item( "OSDAdapter" & sAdapterIndex & sName ) = Value
	End if

End Sub


Function isIPV6Address ( sAddress )
	isIPV6Address = instr(1,sAddress,":",vbTextCompare)
End function


Function ObjNetworkAdapters

	If isempty(g_ObjNetworkAdapters) then

		oLogging.CreateEntry "Query networking adapters...", LogTypeInfo
		Set g_ObjNetworkAdapters = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 1")
		oLogging.CreateEntry "Networking Adapters found! Count = " & g_ObjNetworkAdapters.Count, LogTypeInfo

	End if

	Set ObjNetworkAdapters = g_ObjNetworkAdapters

End Function


Sub RemoveElementFromArray( byref Arr, Index )
	Dim i

	For i = Index to ubound(Arr)-1
		Arr(i) = Arr(i+1)
	Next
	Redim preserve Arr(ubound(Arr)-1)
End sub


Function GetWMINetworkErrorMessage( iRetVal )

	Dim sErrMsg

	Select case (iRetVal)
		Case 1      sErrMsg = "Successful completion, reboot required."
		Case 64     sErrMsg = "Method not supported on this platform."
		Case 65     sErrMsg = "Unknown failure."
		Case 66     sErrMsg = "Invalid subnet mask."
		Case 67     sErrMsg = "An error occurred while processing an instance that was returned."
		Case 68     sErrMsg = "Invalid input parameter."
		Case 69     sErrMsg = "More than five gateways specified."
		Case 70     sErrMsg = "Invalid IP address."
		Case 71     sErrMsg = "Invalid gateway IP address."
		Case 72     sErrMsg = "An error occurred while accessing the registry For the requested information."
		Case 73     sErrMsg = "Invalid domain name."
		Case 74     sErrMsg = "Invalid host name."
		Case 75     sErrMsg = "No primary/secondary WINS server defined."
		Case 76     sErrMsg = "Invalid file."
		Case 77     sErrMsg = "Invalid system path."
		Case 78     sErrMsg = "File copy failed."
		Case 79     sErrMsg = "Invalid security parameter."
		Case 80     sErrMsg = "Unable to configure TCP/IP service."
		Case 81     sErrMsg = "Unable to configure DHCP service."
		Case 82     sErrMsg = "Unable to renew DHCP lease."
		Case 83     sErrMsg = "Unable to release DHCP lease."
		Case 84     sErrMsg = "IP not enabled on adapter."
		Case 85     sErrMsg = "IPX not enabled on adapter."
		Case 86     sErrMsg = "Frame/network number bounds error."
		Case 87     sErrMsg = "Invalid frame type."
		Case 88     sErrMsg = "Invalid network number."
		Case 89     sErrMsg = "Duplicate network number."
		Case 90     sErrMsg = "Parameter out of bounds."
		Case 91     sErrMsg = "Access denied."
		Case 92     sErrMsg = "Out of memory."
		Case 93     sErrMsg = "Already exists."
		Case 94     sErrMsg = "Path, file, or object not found."
		Case 95     sErrMsg = "Unable to notify service."
		Case 96     sErrMsg = "Unable to notify DNS service."
		Case 97     sErrMsg = "Interface not configurable."
		Case 98     sErrMsg = "Not all DHCP leases could be released/renewed."
		Case 100    sErrMsg = "DHCP not enabled on adapter."
		case Else   sErrMsg = "Unknown Error: " & iRetVal
	End select

	GetWMINetworkErrorMessage = sErrMsg

End function

