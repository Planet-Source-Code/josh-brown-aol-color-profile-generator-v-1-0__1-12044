VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Color"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3180
   LinkTopic       =   "Form2"
   ScaleHeight     =   1860
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C000C0&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "These are simple colors, if you want more, go to: http://www.ecohost.com/images/hex2.htm"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Picture1_Click()
  Form1.Text1.Text = "< body . bgcolor=#ff0000>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture2_Click()
  Form1.Text1.Text = "< body . bgcolor=#ffa500>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture3_Click()
  Form1.Text1.Text = "< body . bgcolor=#ffff00>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture4_Click()
  Form1.Text1.Text = "< body . bgcolor=#008000>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture5_Click()
  Form1.Text1.Text = "< body . bgcolor=#000080>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture6_Click()
  Form1.Text1.Text = "< body . bgcolor=#ee82ee>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture7_Click()
  Form1.Text1.Text = "< body . bgcolor=#000000>" & Form1.Text1.Text
  Me.Hide
End Sub

Private Sub Picture8_Click()
  Form1.Text1.Text = "< body . bgcolor=#f8f8ff>" & Form1.Text1.Text
  Me.Hide
End Sub
