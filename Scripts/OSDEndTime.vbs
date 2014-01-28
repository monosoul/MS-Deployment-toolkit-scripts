Option Explicit

Dim	oTSE
Dim	oWSH

  On Error Resume Next

  Set oTSE = CreateObject("Microsoft.SMS.TSEnvironment") 
  Set oWSH = CreateObject("WScript.Shell")

' ||||||||||||||||||||||||||||||||||||||||||
' || Get UTC offset
' ||||||||||||||||||||||||||||||||||||||||||
  Dim	timeBias
  Dim	timeOffset

  timeOffset = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias" 
  timeBias   = oWSH.RegRead(timeOffset)
  
' ||||||||||||||||||||||||||||||||||||||||||
' || Convert 'now' to UTC
' ||||||||||||||||||||||||||||||||||||||||||
  Dim	timeNow
  Dim	timeUTC

  timeNow   = Now() 
  timeUTC   = DateAdd( "n", timeBias, timeNow) 

' ||||||||||||||||||||||||||||||||||||||||||
' || Standardize for Branding / OSDResults
' ||||||||||||||||||||||||||||||||||||||||||
  Dim	timeBranding
  timeBranding =     Year(timeUTC)       & "-" &_
        Right( "0" & Month(timeUTC), 2)  & "-" &_
        Right( "0" & Day(timeUTC), 2)    & " " &_
        Right( "0" & Hour(timeUTC), 2)   & ":" &_
        Right( "0" & Minute(timeUTC), 2) & ":" &_
        Right( "0" & Second(timeUTC), 2) & "Z"

  'wscript.echo " Current  : [" & timeNow & "]"
  'wscript.echo " UTC      : [" & timeUTC & "]"
  'wscript.echo " Branding : [" & timeBranding & "]"

  oTSE("OSDEndTime") = timeBranding