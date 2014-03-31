Const ForReading = 1
Const ForWriting = 2
Const bWaitOnReturn = True

Set oShell = CreateObject("WScript.Shell")
SysDrive=oShell.ExpandEnvironmentStrings("%SystemDrive%")

'Экспортируем шаблоны
oShell.run "dirquota template export /file:" & SysDrive & "\backedup_shares\quota_templates.xml",0,bWaitOnReturn

'Получаем список квот
oShell.run "cmd /c ""dirquota q l > " & SysDrive & "\backedup_shares\all_quotas.txt""",0,bWaitOnReturn
