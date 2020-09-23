VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lightning Messanger"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Send"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      ForeColor       =   &H0000FFFF&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er

Form1.Win.SendData "Meg " & Form1.Win.LocalIP & " says: " & Text1.Text & vbCrLf


Form3.Text1.Text = ""
Form3.Text1.SelText = Form1.Win.LocalIP & " says: " & Text1.Text & vbCrLf
Unload Me
Form3.Show
Exit Sub
er:
MsgBox "An error has occurred, most likely you have forgotten to connect", vbCritical, "Error"

Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
GRAD Me, vbRed, "Down", &H80&, &H80FF&
End Sub

