Set wso=CreateObject("WindowScriptingObject")
x = wso.ActiveWindow
msgbox x, , "vbs"
msgbox wso.windowtext(x), , "vbs"
