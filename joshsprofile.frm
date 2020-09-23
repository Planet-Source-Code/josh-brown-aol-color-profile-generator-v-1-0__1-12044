VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Josh's AOL Color Profile Generator V. 1.0"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Start Over"
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Font"
      Height          =   495
      Left            =   3240
      TabIndex        =   19
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Font Size"
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Backround Color"
      Height          =   495
      Left            =   1320
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Font Color"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Use AOL Profile editor for this one"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "Fill in your info before putting in color and font effects"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label8 
      Caption         =   "Personal Quote:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Occupation:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Computers Used:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Hobbies:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Marital Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Sex:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Member Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form2.Show
End Sub

Private Sub Command2_Click()
  Form3.Show
End Sub

Private Sub Command3_Click()
  Form4.Show
End Sub

Private Sub Command4_Click()
  Form5.Show
End Sub

Private Sub Command5_Click()
  Text1.Text = ""
  Text2.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text7.Text = ""
  Text8.Text = ""
End Sub


Private Sub Form_Load()
  MsgBox "Copy and Paste each line into your AOL profile.", , "Josh's Color Profile Generator V. 1.0"
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub
