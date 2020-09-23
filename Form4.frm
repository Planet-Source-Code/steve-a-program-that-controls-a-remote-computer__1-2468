VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning Messenger"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "&Send"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ip = Form1.Win.RemoteHostIP
Me.Height = 5840
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Winsock2.RemoteHost = ip
Form1.Winsock2.SendData "Meg " & Form1.Win.LocalIP & " says: " & Text1.Text & vbCrLf
Text1.SelText = Form1.Win.LocalIP & " says: " & Text1.Text & vbCrLf
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text2.SetFocus
End Sub

