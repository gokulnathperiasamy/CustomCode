set wsc = CreateObject("WScript.Shell")
Do
    WScript.Sleep(1*60*1000)
    wsc.SendKeys("{SCROLLLOCK 2}")
Loop