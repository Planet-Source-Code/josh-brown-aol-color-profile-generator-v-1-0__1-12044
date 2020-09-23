VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Font"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form5"
   ScaleHeight     =   1065
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Font Name"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the name of the font you would like on your profile."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Text1.Text = "< font . face=" & Text1.Text & " >" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub
