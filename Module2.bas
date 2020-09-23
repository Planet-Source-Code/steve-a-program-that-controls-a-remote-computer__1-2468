Attribute VB_Name = "Module2"
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long


Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (LpVersionInformation As OSVERSIONINFO) As Long
    Public Const VK_MENU = &H12
    Public Const KEYEVENTF_KEYUP = &H2


Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
    End Type
    

Public Sub GetWindowSnapShot(Mode As Long, ThisImage As PictureBox)

    
    ' mode = 0 -> Screen snapshot
    ' mode = 1 -> Window snapshot
    
    Dim altscan%, NT As Boolean, nmode As Long
    
    NT = IsNT


    If Not NT Then
        If Mode = 0& Then Mode = 1& Else Mode = 0&
    End If

    


    If NT And Mode = 0 Then
        keybd_event vbKeySnapshot, 0&, 0&, 0&
    Else
        altscan = MapVirtualKey(VK_MENU, 0)
        keybd_event VK_MENU, altscan, 0, 0


        DoEvents
            keybd_event vbKeySnapshot, Mode, 0&, 0&
        End If



        DoEvents
            ThisImage = Clipboard.GetData(vbCFBitmap)
            keybd_event VK_MENU, altscan, KEYEVENTF_KEYUP, 0
        End Sub



Public Function IsNT() As Boolean

    Dim verinfo As OSVERSIONINFO
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    If (GetVersionEx(verinfo)) = 0 Then Exit Function
    If verinfo.dwPlatformId = 2 Then IsNT = True
End Function

