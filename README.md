Set WshShell = WScript.CreateObject("WScript.Shell") While True WScript.Sleep 600000 WShell.SendKeys "{UP}" WShell.SendKeys "{DOWN}" Wend
