VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Font Size"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2355
   LinkTopic       =   "Form4"
   ScaleHeight     =   735
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number between 8-60"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Text1.Text = "< font . ptsize=" & Text1.Text & " >" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub
