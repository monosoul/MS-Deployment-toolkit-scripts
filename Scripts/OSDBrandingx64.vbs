'////////////////////////////////////////////////////////////////////////
' OSD Deploy Tool Branding Script
'////////////////////////////////////////////////////////////////////////
' Brand OSD variables to registry
'////////////////////////////////////////////////////////////////////////
' V2.00 6.11.2009 MICHS
'////////////////////////////////////////////////////////////////////////
' Include/Exclude
'////////////////////////////////////////////////////////////////////////
' 1. Include or exclude variables "starting with"
' 2. Use semicolon to separate multiple values
' 3. Exclude takes precedence over includes
'////////////////////////////////////////////////////////////////////////
Option Explicit

	'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
	' Constants
	'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
	Const	REG32			= "%windir%\System32\reg.exe"
	Const	REG64			= "%windir%\Sysnative\reg.exe"
	Const	REGBRANDPATH	= "HKLM\Software\WOW6432Node\Microsoft\MPSD\OSD"

	Const   tsAppVariableName      = "TsApplicationBaseVariable"
    Const   tsWindowsAppPackageAppVariableName = "TsWindowsAppPackageAppBaseVariable"
    Const   tsAppInstall           = "TsAppInstall"
    Const	includeMap		= "OSD;_SMSTSClientGUID;_SMSTSClientIdentity;USMT_;TSType;TSVersion;OldComputerName;PACKAGES;OSDBaseVariableName;DeploymentType"
	Const	excludeMap		= "OSDJoinPassword;_SMSTSReserved;OSDLocalAdminPassword"

	'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
	' Globals
	'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
	Dim REGBRAND
	Dim oWSH 				: SET oWSH = CreateObject("WScript.Shell")
	Dim oTSE				: SET oTSE = CreateObject("Microsoft.SMS.TSEnvironment")

	'[##############################################################################################################################]
	' MAIN
	'[##############################################################################################################################]

	'||||||||||||||||||||||||||||||||
	' Determine 32/64 Sysnative
	'||||||||||||||||||||||||||||||||
	Call LogArea("Environmental Setup")
	IF ( IsSysnative() = TRUE ) Then REGBRAND = REG64 Else REGBRAND = REG32

	'||||||||||||||||||||||||||||||||
	' Build Exclude/Include Arrays
	'||||||||||||||||||||||||||||||||
	Call LogArea("Mapping Inclusions and Exclusions")

    Dim applicationsPrefix : applicationsPrefix = oTSE(tsAppVariableName)
    Dim windowsAppPackageAppPrefix : windowsAppPackageAppPrefix = oTSE(tsWindowsAppPackageAppVariableName)
    Dim appInstall : appInstall = oTSE(tsAppInstall)

    'Brand Base Variable Values
    'And Variables that start with these Prefixes

    Dim incArray : incArray = Split( includeMap & ";" & tsAppVariableName & ";" & tsAppInstall & ";" & appInstall & ";" & tsWindowsAppPackageAppVariableName & ";" & windowsAppPackageAppPrefix  , ";" )
	Dim excArray : excArray = Split( excludeMap, ";" )

	'||||||||||||||||||||||||||||||||
	' Loop through TS Variables
	'||||||||||||||||||||||||||||||||
	Call LogArea("Branding Registry")
	Call BrandValue( "InstalledOn", Date )

	Dim tV
    For Each tV in oTSE.GetVariables()
		IF (MatchMaker( tV, incArray ) = TRUE) Then
			IF (MatchMaker( tV, excArray ) = FALSE ) Then
				Call BrandValue( tV, oTSE(tV) )
			End IF
		End IF
    Next

   'Brand Applications
    For Each tV in oTSE.GetVariables()
        If ( InStr(1, tV, applicationsPrefix, 1) = 1 ) Then
                    Call BrandValue( Replace(tV, applicationsPrefix, UCase(applicationsPrefix) & "0",1,-1, 1), oTSE(tV) )
		End IF
    Next

	WScript.Quit(0)

	'[##############################################################################################################################]
	' FUNCTIONS
	'[##############################################################################################################################]

	' ////////////////////////////////////////////////////
	' Brand a name and value to registry
	' ////////////////////////////////////////////////////
	Sub BrandValue( theName, theValue )

		Dim retVal : retVal = 0
		Dim runCmd : runCmd = REGBRAND & " ADD " & REGBRANDPATH & " /F /V " & theName & " /T REG_SZ /D """ & theValue & """"

		Wscript.Echo " Branding : [" & runCmd & "]"
		retVal = oWSH.Run( runCMD, 0, True )
		Wscript.Echo " Result   : [" & retVal & "]"

	End Sub

	' ////////////////////////////////////////////////////
	' Match "StartsWith" against an array of values
	' ////////////////////////////////////////////////////
	Function MatchMaker(theItem, theArray)
		Dim retVal : retVal = FALSE

		Dim anItem
		For Each anItem in theArray
			If ( Len(anItem)=0 ) Then Exit For
			' ||||||||||||||||||||||||||||||||
			'  - StartsWith is position 1
			'  - Case/Text Insensitive is 1
			' ||||||||||||||||||||||||||||||||
			If ( InStr(1, theItem, anItem, 1) = 1 ) Then
				retVal = TRUE
				Exit For
			End If
		Next

		MatchMaker = retVal

	End Function

	' ////////////////////////////////////////////////////
	' Detects if 32-bit environment on 64-bit OS
	' ////////////////////////////////////////////////////
	Function IsSysnative()

		Dim	PARCH1 : PARCH1 = UCASE( oWSH.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") )
		Dim	PARCH2 : PARCH2 = UCASE( oWSH.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%") )

		wscript.echo "%PROCESSOR_ARCHITECTURE% = [" & PARCH1 & "]"
		wscript.echo "%PROCESSOR_ARCHITEW6432% = [" & PARCH2 & "]"

		IF ( (PARCH1 = "X86") AND (PARCH2 = "AMD64") ) Then IsSysnative=TRUE _
		ELSE IsSysnative = FALSE

		wscript.echo "32-BIT Environment on a 64-BIT OS: [" & IsSysnative & "]"

	End Function

	' ////////////////////////////////////////////////////
	' Log Area
	' ////////////////////////////////////////////////////
	Sub LogArea( theText )

		Wscript.Echo
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo " " & theText
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo

	End Sub
'' SIG '' Begin signature block
'' SIG '' MIIaWgYJKoZIhvcNAQcCoIIaSzCCGkcCAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFM3KdrSqJuAj
'' SIG '' kwtlbQNllcJ16BjwoIIVNjCCBKkwggORoAMCAQICEzMA
'' SIG '' AACIWQ48UR/iamcAAQAAAIgwDQYJKoZIhvcNAQEFBQAw
'' SIG '' eTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EwHhcNMTIwNzI2
'' SIG '' MjA1MDQxWhcNMTMxMDI2MjA1MDQxWjCBgzELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjENMAsGA1UECxMETU9QUjEeMBwGA1UE
'' SIG '' AxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAs3R00II8h6ea
'' SIG '' 1I6yBEKAlyUu5EHOk2M2XxPytHiYgMYofsyKE+89N4w7
'' SIG '' CaDYFMVcXtipHX8BwbOYG1B37P7qfEXPf+EhDsWEyp8P
'' SIG '' a7MJOLd0xFcevvBIqHla3w6bHJqovMhStQxpj4TOcVV7
'' SIG '' /wkgv0B3NyEwdFuV33fLoOXBchIGPfLIVWyvwftqFifI
'' SIG '' 9bNh49nOGw8e9OTNTDRsPkcR5wIrXxR6BAf11z2L22d9
'' SIG '' Vz41622NAUCNGoeW4g93TIm6OJz7jgKR2yIP5dA2qbg3
'' SIG '' RdAq/JaNwWBxM6WIsfbCBDCHW8PXL7J5EdiLZWKiihFm
'' SIG '' XX5/BXpzih96heXNKBDRPQIDAQABo4IBHTCCARkwEwYD
'' SIG '' VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFCZbPltd
'' SIG '' ll/i93eIf15FU1ioLlu4MA4GA1UdDwEB/wQEAwIHgDAf
'' SIG '' BgNVHSMEGDAWgBTLEejK0rQWWAHJNy4zFha5TJoKHzBW
'' SIG '' BgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNDb2RT
'' SIG '' aWdQQ0FfMDgtMzEtMjAxMC5jcmwwWgYIKwYBBQUHAQEE
'' SIG '' TjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NlcnRzL01pY0NvZFNpZ1BDQV8w
'' SIG '' OC0zMS0yMDEwLmNydDANBgkqhkiG9w0BAQUFAAOCAQEA
'' SIG '' D95ASYiR0TE3o0Q4abJqK9SR+2iFrli7HgyPVvqZ18qX
'' SIG '' J0zohY55aSzkvZY/5XBml5UwZSmtxsqs9Q95qGe/afQP
'' SIG '' l+MKD7/ulnYpsiLQM8b/i0mtrrL9vyXq7ydQwOsZ+Bpk
'' SIG '' aqDhF1mv8c/sgaiJ6LHSFAbjam10UmTalpQqXGlrH+0F
'' SIG '' mRrc6GWqiBsVlRrTpFGW/VWV+GONnxQMsZ5/SgT/w2at
'' SIG '' Cq+upN5j+vDqw7Oy64fbxTittnPSeGTq7CFbazvWRCL0
'' SIG '' gVKlK0MpiwyhKnGCQsurG37Upaet9973RprOQznoKlPt
'' SIG '' z0Dkd4hCv0cW4KU2au+nGo06PTME9iUgIzCCBLowggOi
'' SIG '' oAMCAQICCmECkkoAAAAAACAwDQYJKoZIhvcNAQEFBQAw
'' SIG '' dzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBMB4XDTEyMDEwOTIy
'' SIG '' MjU1OVoXDTEzMDQwOTIyMjU1OVowgbMxCzAJBgNVBAYT
'' SIG '' AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
'' SIG '' EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
'' SIG '' cG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsT
'' SIG '' Hm5DaXBoZXIgRFNFIEVTTjpCOEVDLTMwQTQtNzE0NDEl
'' SIG '' MCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
'' SIG '' dmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
'' SIG '' ggEBAM1jw/eitUfZ+TmUU6xrj6Z5OCH00W49FTgWwXMs
'' SIG '' mY/74Dxb4aJMi7Kri7TySse5k1DRJvWHU7B6dfNHDxcr
'' SIG '' Zyxk62DnSozgi17EVmk3OioEXRcByL+pt9PJq6ORqIHj
'' SIG '' Py232OTEeAB5Oc/9x2TiIxJ4ngx2J0mPmqwOdOMGVVVJ
'' SIG '' yO2hfHBFYX6ycRYe4cFBudLSMulSJPM2UATX3W88SdUL
'' SIG '' 1HZA/GVlE36VUTrV/7iap1drSxXlN1gf3AANxa7q34FH
'' SIG '' +fBSrubPWqzgFEqmcZSA+v2wIzBg6YNgrA4kHv8R8uel
'' SIG '' VWKV7p9/ninWzUsKdoPwQwTfBkkg8lNaRLBRejkCAwEA
'' SIG '' AaOCAQkwggEFMB0GA1UdDgQWBBTNGaxhTZRnK/avlHVZ
'' SIG '' 2/BYAIOhOjAfBgNVHSMEGDAWgBQjNPjZUkZwCu1A+3b7
'' SIG '' syuwwzWzDzBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
'' SIG '' Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0
'' SIG '' cy9NaWNyb3NvZnRUaW1lU3RhbXBQQ0EuY3JsMFgGCCsG
'' SIG '' AQUFBwEBBEwwSjBIBggrBgEFBQcwAoY8aHR0cDovL3d3
'' SIG '' dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3Nv
'' SIG '' ZnRUaW1lU3RhbXBQQ0EuY3J0MBMGA1UdJQQMMAoGCCsG
'' SIG '' AQUFBwMIMA0GCSqGSIb3DQEBBQUAA4IBAQBRHNbfNh3c
'' SIG '' gLwCp8aZ3xbIkAZpFZoyufNkENKK82IpG3mPymCps13E
'' SIG '' 5BYtNYxEm/H0XGGkQa6ai7pQ0Wp5arNijJ1NUVALqY7U
'' SIG '' v6IQwEfVTnVSiR4/lmqPLkAUBnLuP3BZkl2F7YOZ+oKE
'' SIG '' nuQDASETqyfWzHFJ5dod/288CU7VjWboDMl/7jEUAjdf
'' SIG '' e2nsiT5FfyVE5x8a1sUaw0rk4fGEmOdP+amYpxhG7IRs
'' SIG '' 7KkDCv18elIdnGukqA+YkqSSeFwreON9ssfZtnB931tz
'' SIG '' U7+q1GZQS/DJO5WF5cFKZZ0lWFC7IFSReTobB1xqVyiv
'' SIG '' Mcef58Md7kf9J9d/z3TcZcU/MIIFvDCCA6SgAwIBAgIK
'' SIG '' YTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYK
'' SIG '' CZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJ
'' SIG '' bWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9v
'' SIG '' dCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
'' SIG '' MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29k
'' SIG '' ZSBTaWduaW5nIFBDQTCCASIwDQYJKoZIhvcNAQEBBQAD
'' SIG '' ggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBCmXZTbD4b
'' SIG '' 1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1
'' SIG '' VwqJyq4gSfTwaKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNE
'' SIG '' f9eB2/O98jakyVxF3K+tPeAoaJcap6Vyc1bxF5Tk/TWU
'' SIG '' cqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2
'' SIG '' L1TdFDBZ+NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049
'' SIG '' oDI9kM2hOAaFXE5WgigqBTK3S9dPY+fSLWLxRT3nrAgA
'' SIG '' 9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzB
'' SIG '' rAlfA9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMB
'' SIG '' Af8wHQYDVR0OBBYEFMsR6MrStBZYAck3LjMWFrlMmgof
'' SIG '' MAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
'' SIG '' MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+m
'' SIG '' PLzYLTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAf
'' SIG '' BgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnkpDBQ
'' SIG '' BgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9taWNyb3Nv
'' SIG '' ZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
'' SIG '' CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
'' SIG '' b20vcGtpL2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNy
'' SIG '' dDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+fyZGr+tvQLEy
'' SIG '' tWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAh
'' SIG '' Xaw9L0y6oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6Wc
'' SIG '' IC36C1DEVs0t40rSvHDnqA2iA6VW4LiKS1fylUKc8fPv
'' SIG '' 7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WR
'' SIG '' Tsgb0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZE
'' SIG '' olHN+emjWFbdmwJFRC9f9Nqu1IIybvyklRPk62nnqaIs
'' SIG '' vsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
'' SIG '' NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxP
'' SIG '' bQ4TTj18KUicctHzbMrB7HCjV5JXfZSNoBtIA1r3z6Nn
'' SIG '' CnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDordEN5k
'' SIG '' 9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7
'' SIG '' B2BHZWBATrBC7E7ts3Z52Ao0CW0cgDEf4g5U3eWh++VH
'' SIG '' EK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jshrg1cnPCi
'' SIG '' roZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONx
'' SIG '' PdcAfmJH0c6IybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3
'' SIG '' UAjWwz0wggYHMIID76ADAgECAgphFmg0AAAAAAAcMA0G
'' SIG '' CSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNj
'' SIG '' b20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTAr
'' SIG '' BgNVBAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmljYXRl
'' SIG '' IEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0
'' SIG '' MDMxMzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
'' SIG '' EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
'' SIG '' HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
'' SIG '' BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCC
'' SIG '' ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJ+h
'' SIG '' bLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn0Uyt
'' SIG '' dDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4E
'' SIG '' mPCJzB/LMySHnfL0Zxws/HvniB3q506jocEjU8qN+kXP
'' SIG '' CdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4nrIZPVVIM
'' SIG '' 5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0
'' SIG '' RZCfSABKR2YRJylmqJfk0waBSqL5hKcRRxQJgp+E7VV4
'' SIG '' /gGaHVAIhQAQMEbtt94jRrvELVSfrx54QTF3zJvfO4OT
'' SIG '' oWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAasw
'' SIG '' ggGnMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0
'' SIG '' +NlSRnAK7UD7dvuzK7DDNbMPMAsGA1UdDwQEAwIBhjAQ
'' SIG '' BgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQO
'' SIG '' rIJgQFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmS
'' SIG '' JomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1p
'' SIG '' Y3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
'' SIG '' Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxz
'' SIG '' WPQHEy5lMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9j
'' SIG '' cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
'' SIG '' L21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcB
'' SIG '' AQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly93d3cubWlj
'' SIG '' cm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0Um9v
'' SIG '' dENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0G
'' SIG '' CSqGSIb3DQEBBQUAA4ICAQAQl4rDXANENt3ptK132855
'' SIG '' UU0BsS50cVttDBOrzr57j7gu1BKijG1iuFcCy04gE1CZ
'' SIG '' 3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji
'' SIG '' 8FMV3U+rkuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZ
'' SIG '' Lg33B+JwvBhOnY5rCnKVuKE5nGctxVEO6mJcPxaYiyA/
'' SIG '' 4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tO
'' SIG '' i3/FNSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLa
'' SIG '' FJj1PLlmWLMtL+f5hYbMUVbonXCUbKw5TNT2eb+qGHpi
'' SIG '' Ke+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
'' SIG '' NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCe
'' SIG '' FTBm6EISXhrIniIh0EPpK+m79EjMLNTYMoBMJipIJF9a
'' SIG '' 6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2JoXZh
'' SIG '' tG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0t
'' SIG '' r1mPuOQh5bWwymO0eFQF1EEuUKyUsKV4q7OglnUa2ZKH
'' SIG '' E3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng9wFlb4kL
'' SIG '' fchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj
'' SIG '' /TGCBJAwggSMAgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25p
'' SIG '' bmcgUENBAhMzAAAAiFkOPFEf4mpnAAEAAACIMAkGBSsO
'' SIG '' AwIaBQCggbIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' IwYJKoZIhvcNAQkEMRYEFILOYI/P8JuuQ1fniAdlqGbs
'' SIG '' jRXgMFIGCisGAQQBgjcCAQwxRDBCoCSAIgBNAEQAVAAg
'' SIG '' AFUARABJAHYAMwAgAFQAbwBvAGwAawBpAHShGoAYaHR0
'' SIG '' cDovL3d3dy5taWNyb3NvZnQuY29tMA0GCSqGSIb3DQEB
'' SIG '' AQUABIIBABecYP13bFRV+c/w7EePFBgeTBF+/Ar72LAi
'' SIG '' cswewyVdOCo7pMUUlyU0QmyRcFDvbHGSY6AcpiLrAhkj
'' SIG '' RUCU1GqT2b8A7gyEzziC+ULUVI2rR3+O+N4jImehOo5W
'' SIG '' j/e2hpuoG68n7whf7ncYXEbnOR3zkWuxPVlsevomHrJ7
'' SIG '' ewtXPSxD7786J5kXsskKagc6bS38H60CTyIblkfki6yX
'' SIG '' HifuZLpf5WoJaVLpu/hL4i4LvZrq09B6lDpekZ9cEwcE
'' SIG '' CmHaBSWeT+95jYHKMEqhUb2rLbIl+Ta6yCafh9xdBJGL
'' SIG '' 9yZpgn8bgFvyBSONAfXO5Mdz9jwajidmpE4qp22D81qh
'' SIG '' ggIfMIICGwYJKoZIhvcNAQkGMYICDDCCAggCAQEwgYUw
'' SIG '' dzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBAgphApJKAAAAAAAg
'' SIG '' MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
'' SIG '' hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xMjA4MzAxOTU3
'' SIG '' MDJaMCMGCSqGSIb3DQEJBDEWBBSHJuM1p6R6KDNceVIU
'' SIG '' JywGAkayETANBgkqhkiG9w0BAQUFAASCAQB8+DP3PNH5
'' SIG '' wRQNbUFL5IS+D6smNFI91YW35AGb+7LeIvfmJYIYKtdz
'' SIG '' JV5ZfeoAbekGMmSkC8YZx5ruTl8RBTM01LwZdWFh186/
'' SIG '' MA58/qiiZ0WZ7ilgjSccyvOEgduWCvmXndJcFoAlLOo7
'' SIG '' CVZAWANGd46W+22499cZ+wmAmcKm3R/5KuetEpvzDZ5r
'' SIG '' i0CuYKOYKAGgtsYB+I2SgB6SX6089lW483looC4FTQes
'' SIG '' yI/I4C0pLiGq9FParWrwVeDy31hbutIN1qFzn/jFgCTT
'' SIG '' qc0Kv6eGISZ3/cfkawHZ9I5siUO2b/py3rZ4Jn9M94ZA
'' SIG '' DtnXk3ktZyhdnhaVT9hcA+QM
'' SIG '' End signature block
