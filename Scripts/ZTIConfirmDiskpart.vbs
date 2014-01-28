' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      ZTIConfFirmDiskpart.wsf
' // 
' // Version:   6.1.2373.0
' // 
' // Purpose:   Solution Accelerator for Business Desktop Deployment
' // 
' // Usage:     cscript ZTIConfirmDiskpart.vbs 
' // 
' // ***************************************************************************


	' Hide the progress dialog

	Set oTSProgressUI = CreateObject("Microsoft.SMS.TSProgressUI") 
	oTSProgressUI.CloseProgressDialog 
	Set oTSProgressUI = Nothing

	iretval = msgbox ("The task sequence was unable to locate a logical drive" & VBCRLF &   "The hard disk will need to be partitioned and formatted." & VBCRLF &  "Click on OK to continue or Cancel to exit the task sequence", 1)


	If iRetVal <> 1 Then
		Wscript.Quit 1
	End If
		