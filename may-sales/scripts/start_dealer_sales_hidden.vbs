Set WshShell = CreateObject("WScript.Shell")
Set WMI = GetObject("winmgmts:\\.\root\cimv2")
Set Procs = WMI.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='python.exe' AND CommandLine LIKE '%dealer_sales_http.py%'")
If Procs.Count > 0 Then
    WScript.Quit 0
End If

WshShell.Run """C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"" ""D:\Services\robot-services\may-sales\src\dealer_sales_http.py"" --port 8003", 0, False
