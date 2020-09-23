VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9225
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   2820
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Your IP:"
            TextSave        =   "Your IP:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "4:47 PM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "7/11/99"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800000&
      Caption         =   "Connection Info"
      ForeColor       =   &H0000FFFF&
      Height          =   2655
      Left            =   5760
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin RichTextLib.RichTextBox rt 
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4048
         _Version        =   327680
         BackColor       =   8421504
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmmain.frx":08CA
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800000&
      Caption         =   "Functions"
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
      Begin VB.CommandButton Command5 
         Caption         =   "&About"
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Use"
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00808080&
         Height          =   840
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Fun Stuff"
      ForeColor       =   &H0000FFFF&
      Height          =   2655
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command6 
         Caption         =   "&About"
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Use"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00808080&
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "IP of computer to control"
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton Command7 
         Caption         =   "&Save"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Connect"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00808080&
         Height          =   645
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock Win 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem Text1.Text
Text1.Text = ""
End Sub

Private Sub Command2_Click()
On Error GoTo er
Win.RemoteHost = List1.Text
Win.SendData "Connect"
ip = List1.Text
Exit Sub
er:
MsgBox "An error has occurred.  Most likely you forgot to highlight an address", vbCritical, "Error"
Exit Sub
End Sub

Private Sub Command3_Click()

Dim command As String
On Error GoTo er
    command = List2.Text
    With Win
     .SendData command
    End With
    Exit Sub
er:
Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo er
Select Case List3.Text
Case "Open a website"
mf = InputBox("Enter the url of the website you would like to open:", "Enter url")
Win.SendData "Web " & mf
Case "Open a program"
mf = InputBox("Enter the location of the file you would like to open:", "Enter File Location")
Win.SendData "Ope " & mf
Case "Annoying loop"
mf = InputBox("Enter the message you would like displayed:", "Enter the message")
l = InputBox("Enter how many times you want the loop to repeat:", "Number of repeats")
For X = 0 To l
Win.SendData "Mes " & mf
Next X
Case "Send a message"
Form2.Show
End Select
er:
Exit Sub
End Sub

Private Sub Command5_Click()
MsgBox "These commands do require some input from you." & Chr(10) & Chr(10) & "Any command in here, does require you to input some information.  Such as a url, or a filename, and other such things.", vbQuestion
End Sub

Private Sub Command6_Click()
MsgBox "These commands don't require any input from you." & Chr(10) & Chr(10) & "Any commands in here simply require you to select them, and press ok.  The program won't require any inputed information.", vbQuestion
End Sub

Private Sub Command7_Click()
SaveListBox List1, App.Path & "ip.lst"
End Sub

Private Sub Form_Load()

 With Winsock2
     .Protocol = sckUDPProtocol
     .RemotePort = 693
     .Bind
    End With
    With Win
     .Protocol = sckUDPProtocol
     .RemotePort = 692
     .Bind
    End With

    List2.AddItem "Open Cdrom"
    List2.AddItem "Close Cdrom"
    List3.AddItem "Open a website"
    List3.AddItem "Open a program"
    List3.AddItem "Annoying loop"
    List3.AddItem "Send a message"
    StatusBar1.Panels(1).Text = StatusBar1.Panels(1).Text & " " & Win.LocalIP
LoadListBox List1, App.Path & "\ip.lst"
End Sub

Private Sub Win_DataArrival(ByVal bytesTotal As Long)

Dim data As String
Win.GetData data

rt.SelText = Win.RemoteHostIP & ": " & data & vbCrLf

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case Left(data, 3)
Case "Meg"
ms = Right(data, Len(data) - 4)
Form3.Text1.Text = data
If Form3.Visible = False Then
Form3.Show
Else
Exit Sub
End If
End Select
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
