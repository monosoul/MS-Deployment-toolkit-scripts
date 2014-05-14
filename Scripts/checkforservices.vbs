const HKEY_LOCAL_MACHINE = &H80000002
const strComputer = "."
Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True
Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")
Set objFSO = CreateObject("Scripting.FileSystemObject")

SvcName = WScript.Arguments(0)
RootKey = "SYSTEM\CurrentControlSet\Services\"
SrvKey = RootKey & SvcName

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
If oReg.EnumKey(HKEY_LOCAL_MACHINE, SrvKey, arrSubKeys) = 0 Then
  If (oShell.RegRead("HKEY_LOCAL_MACHINE\" & SrvKey & "\Start") = 2) Then
    Set objFileOut = objFSO.OpenTextFile(SysDrive & "\backedup_shares\" & SvcName, ForWriting, True)
	objFileOut.Write(SvcName)
	objFileOut.Close
  End If
End If
