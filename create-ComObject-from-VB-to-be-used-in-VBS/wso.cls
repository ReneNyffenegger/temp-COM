' vi: ft=basic
'
'     https://social.msdn.microsoft.com/Forums/en-US/df0248cb-612f-4e2f-9665-11c68c401458/this-is-how-to-call-win32-api-calls-in-vbscript-and-jscript-sample-windows-api-functions?forum=scripting
' 
'     "C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /target:library /out:wso.dll wso.cls /verbose
'
'     "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm" /codebase wso.dll /tlb:wso.tlb          /v    rem <<-- this didn't work in Zicke because of missing admiminstrative rights
'     "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm" /codebase wso.dll              /regfile /v    rem <<-- this workd
'
'     reg import wso.reg
'
Imports System
    Imports System.Runtime.InteropServices
    Imports Microsoft.Win32
   
    Namespace WindowScriptingObject
   
        <Guid("7448E08D-ED0F-4E23-B528-91937BB41756"), _
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> _
       Public Interface _WindowScriptingObject
            <DispId(1)> Function ActiveWindow() As Integer
            <DispId(2)> Function WindowText(ByVal hWnd As Integer) As String
        End Interface
   
        <Guid("B146BF9E-78FC-4DB0-ABFE-9FF026B43E4D"), _
         ClassInterface(ClassInterfaceType.None), _
         ProgId("WindowScriptingObject")> Public Class WindowScriptingObject
            Implements _WindowScriptingObject
   
            Public WindowScriptingObject()
   
            Public Declare Auto Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow"() As Integer
            Public Declare Auto Function GetWindowText Lib "user32.dll" (ByVal hwnd As Int32, <Out()> ByVal lpString As System.Text.StringBuilder, ByVal cch As Int32) As Int32
            Public Function ActiveWindow() As Integer Implements _WindowScriptingObject.ActiveWindow
     ActiveWindow=GetForegroundWindow()
   
            End Function
   
            Public Function WindowText(hwnd as Integer) As String Implements _WindowScriptingObject.WindowText
     on error resume next
     Dim b As New System.Text.StringBuilder(ChrW(0), 512)
                    Dim ret = GetWindowText(hWnd, b, b.Capacity)
     WindowText = b.tostring
            End Function
   
   
        End Class
   
    End Namespace'
