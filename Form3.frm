VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4815
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Send"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6000
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ip = Form1.Winsock1.RemoteHostIP
Me.Height = 6630
Text2.Visible = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Winsock2.RemoteHost = ip
Form1.Winsock2.SendData "Meg " & ip & " says: " & Text2.Text & vbCrLf
Text1.SelText = ip & " says: " & Text2.Text & vbCrLf
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

