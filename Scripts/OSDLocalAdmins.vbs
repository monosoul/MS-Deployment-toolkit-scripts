  ' //////////////////////////////////
  ' // OSD Local Admins
  ' //////////////////////////////////////////////////////////////////////////////////////////////////////
  ' 1. Script will parse the TS variable "OSDAddAdmin" for a semicolon separated list of accounts.
  ' 2. Script will find the local administrator group name.
  ' 3. Script will add each account to the local administrator group
  ' //////////////////////////////////////////////////////////////////////////////////////////////////////

  
  Dim AdminArray
  Dim AdminGroup
  Dim TSVarAdmins
  
  wscript.echo
  wscript.echo "-------------------------------------"
  wscript.echo " Initializing TS Environment"
  wscript.echo "-------------------------------------"
  wscript.echo

  SET TSEnv = CreateObject("Microsoft.SMS.TSEnvironment")
  TSVarAdmins             = TSEnv("OSDAddAdmin")

  wscript.echo
  wscript.echo "-------------------------------------"
  wscript.echo " Finding Administrator Group Name"
  wscript.echo "-------------------------------------"
  wscript.echo
  
  Call GetAdminGroupName( AdminGroup )
 
  wscript.echo
  wscript.echo "-------------------------------------"
  wscript.echo " Parsing/Splitting Accounts"
  wscript.echo "-------------------------------------"
  wscript.echo

  AdminArray              = Split(TSVarAdmins, ";")
  
  For i = LBound(AdminArray) To UBound(AdminArray)
	Call AddUserAdmin ( Trim(AdminArray(i)), AdminGroup )
  Next

  wscript.echo
  wscript.echo "-------------------------------------"
  wscript.echo " Script End"
  wscript.echo "-------------------------------------"
  wscript.echo

  wscript.Quit (0)
  

  
  '########################################
  ' Add Administrator
  '########################################
  Sub AddUserAdmin(theUser, theGroup)
  
	Dim oWSH
	Dim oEXE
	Dim exeLine
	Dim outLine
	
	exeLine = "net localgroup " & theGroup & " /add " & Chr(34) & theUser & Chr(34)
	
	wscript.echo
	wscript.echo "-------------------------------------"
	wscript.echo " Add Administrator"
	wscript.echo "-------------------------------------"
	wscript.echo " [CMDEXE]: " & exeLine
	
	Set oWSH = CreateObject("Wscript.Shell")
	Set oEXE = oWSH.Exec(exeLine)

	outLine = oEXE.StdOut.ReadAll
	wscript.echo " [STDOUT]: " & outLine
	outLine = oEXE.StdErr.ReadAll
	wscript.echo " [STDERR]: " & outLine
	
  End Sub

  
  '########################################
  ' Get Administrator Group Name
  ' May vary from region-to-region
  '########################################
  Sub GetAdminGroupName( ByRef outName )
	Dim oWMI
	Dim oQRY

	Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
	Set oQRY = oWMI.ExecQuery  ("Select * From Win32_Group Where LocalAccount = TRUE And SID = 'S-1-5-32-544'")
  
	For Each anAccount in oQRY
		outName = anAccount.Name
		wscript.echo "Group Name Found: [" & outName & "]"
	Next
  
  End Sub
  

