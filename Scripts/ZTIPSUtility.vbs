' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTIPSUtility.vbs
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Execute a PowerShell script
' // 
' // Usage:     cscript ZTIPowerShell.wsf
' // 
' // ***************************************************************************


Function RunPowerShellScript(sScriptName, bNative)

	Dim sModules
	Dim sSource
	Dim sCmd
	Dim iRetVal
	Dim sVersion
	Dim bPrefix
	Dim sScript


	' Copy modules locally if they aren't already, to avoid .NET trust issues

	sModules = oEnvironment.Item("DeployRoot") & "\Tools\Modules"
	If Left(oEnvironment.Item("DeployRoot"), 2) = "\\" then

		sSource = sModules
		sModules = oUtility.LocalRootPath & "\Modules"
		If not oFSO.FolderExists(sModules) then
			oLogging.CreateEntry "Creating " & sModules & " folder for caching PowerShell modules locally.", LogTypeInfo
			oFSO.CreateFolder sModules
			oLogging.CreateEntry "Copying " & sSource & " folder to " & sModules, LogTypeInfo
			oFSO.CopyFolder sSource, sModules, true
		End if
		
	End if


	' Decide if we need to thunk for an x86 process

	bPrefix = false
	If UCase(oEnv("PROCESSOR_ARCHITEW6432")) = "AMD64" and (not bNative) then
		bPrefix = true
	End if


	' Determine the appropriate PowerShell host to run

	sVersion = ""
	On Error Resume Next
	Err.Clear
	sVersion = oShell.RegRead("HKLM\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine\PowerShellVersion")
	If Err then
		sVersion = oShell.RegRead("HKLM\SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\PowerShellVersion")
	End if
	On Error Goto 0
	If sVersion = "" then
		oLogging.CreateEntry "WARNING: PowerShell was not detected.", LogTypeWarning
	Else
		oLogging.CreateEntry "PowerShell version detected: " & sVersion, LogTypeInfo
	End if

	If sVersion = "3.0" then
		sCmd = """" & sModules & "\Microsoft.BDD.TaskSequenceModule\Microsoft.BDD.TaskSequencePSHost40.exe"""
		If bPrefix then
			sCmd = """" & sModules & "\Microsoft.BDD.TaskSequenceModule\Microsoft.BDD.Thunk40.exe"" " & sCmd
		End if
	Else
		sCmd = """" & sModules & "\Microsoft.BDD.TaskSequenceModule\Microsoft.BDD.TaskSequencePSHost35.exe"""
		If bPrefix then
			sCmd = """" & sModules & "\Microsoft.BDD.TaskSequenceModule\Microsoft.BDD.Thunk35.exe"" " & sCmd
		End if
	End if


	' Get the full path to the script

	iRetVal = oUtility.FindFile(sScriptName, sScript)
	If iRetVal <> Success then
		oLogging.CreateEntry "Unable to locate script " & sScriptName, LogTypeError
		RunPowerShellScript = 10901
		Exit Function
	End if


	' Add the command line parameters

	sCmd = sCmd & " """ & sScript & """ """ & oUtility.LogPath & """ " & oEnvironment.Item("Parameters")


	' Run the command line

	oLogging.CreateEntry "About to run: " & sCmd, LogTypeInfo
	iRetVal = oShell.Run(sCmd, 0, true)


	' Return the exit code

	RunPowerShellScript = iRetVal
		
End Function
