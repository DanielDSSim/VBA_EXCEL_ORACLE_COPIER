set shell = WScript.CreateObject("WScript.Shell")

col = InputBox("How many col of data to input?")
row = InputBox("How many row of data to input?")
x = MsgBox("Open destination file, click on first box and press OK", 0)
x = MsgBox("Open source file, click on first box and press OK", 0)

for i=1 to row 
	shell.SendKeys "^c"
	WScript.sleep 1000
	shell.SendKeys "%{TAB}"
	WScript.sleep 1000
	shell.SendKeys "^v"

	for j=2 to col
		WScript.sleep 1000
		shell.SendKeys "{TAB}"
		WScript.sleep 1000
		shell.SendKeys "%{TAB}"
		WScript.sleep 2000
		shell.SendKeys "{RIGHT}"
		WScript.sleep 1000
		shell.SendKeys "^c"
		WScript.sleep 1000
		shell.SendKeys "%{TAB}"
		WScript.sleep 2000
		shell.SendKeys "^v"
	next
	shell.SendKeys "%{TAB}"
	WScript.sleep 2000
	shell.SendKeys "^{LEFT}"
	shell.SendKeys "{DOWN}"
next 

WScript.Echo "Script ended"