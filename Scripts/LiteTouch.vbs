' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      Litetouch.vbs
' // 
' // Version:   6.0.2149.0
' // 
' // Purpose:   Start the lite touch deployment process
' // 
' // Usage:     cscript LiteTouch.vbs [/debug:true]
' // 
' // ***************************************************************************

'//----------------------------------------------------------------------------
'//
'//  Global constant and variable declarations
'//
'//----------------------------------------------------------------------------

Option Explicit

Dim oShell
Dim oFSO
Dim iRetVal
Dim sCmd
Dim sScriptDir
Dim sArg
Dim sArgString
Dim sArchitecture
Dim oDrive

'//----------------------------------------------------------------------------
'//  Initialization
'//----------------------------------------------------------------------------

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)
For each sArg in WScript.Arguments
	sArgString = sArgString & sArg & " "
Next


'Change the Architecture type from amd64 to X64 for consistency

sArchitecture = oShell.Environment("SYSTEM").Item("Processor_Architecture")
if lcase(sArchitecture) = "amd64" then
	sArchitecture = "x64"
end if


' Clean up any existing C:\MININT directory

sArgString = " /CleanStart " & sArgString

oShell.Environment("PROCESS")("SEE_MASK_NOZONECHECKS") = 1


' Clean up any remnants of a previous task sequence

On Error Resume Next
iRetVal = oShell.Run("reg.exe delete HKCR\Microsoft.SMS.TSEnvironment /f", 0, true)
iRetVal = oShell.Run("reg.exe delete HKCR\Microsoft.SMS.TSEnvironment.1 /f", 0, true)
iRetVal = oShell.Run("reg.exe delete HKCR\Microsoft.SMS.TSProgressUI /f", 0, true)
On Error Goto 0


'//----------------------------------------------------------------------------
'//  Check to see if the prereq's have been satisfied
'//----------------------------------------------------------------------------

sCmd = "cscript.exe """ & sScriptDir & "\ZTIPrereq.vbs"""
iRetVal = oShell.Run(sCmd, 0, true)
If iRetVal <> 0 then
	oShell.Popup "This computer does not meet the prerequisites for deploying a new operating system.  (" & CStr(iRetVal) & ")", 0, "Prerequisite Error", 16
	WScript.Quit iRetVal
End if


'//----------------------------------------------------------------------------
'//  Launch LiteTouch.wsf to do the heavy lifting
'//----------------------------------------------------------------------------

sCmd = """" & sScriptDir & "\..\tools\" & sArchitecture & "\bddrun.exe"" wscript.exe """ & sScriptDir & "\LiteTouch.wsf"" " & sArgString
iRetVal = oShell.Run(sCmd, 1, true)


' All done

WScript.Quit iRetVal
