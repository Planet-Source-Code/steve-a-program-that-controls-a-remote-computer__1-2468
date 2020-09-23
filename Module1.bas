Attribute VB_Name = "Module1"

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean

Global Const EWX_SHUTDOWN = 1
Global Const EWX_FORCE = 4
Global Const EWX_REBOOT = 2
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
            

'CD Door close command
Sub CDClose()
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
'CD Door open command
Sub CDOpen()
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub
'Form Top most command
Sub TopMost(Mee As Form)
Dim lResult As Long
    lResult = SetWindowPos(Mee.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
'Shut down computer command
Sub ShutDown()
 ExitWindowsEx EWX_SHUTDOWN, 0
End Sub


