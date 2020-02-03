VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   6210
   ClientTop       =   2310
   ClientWidth     =   6435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   7305
   ScaleWidth      =   6435
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000003&
      Caption         =   "ÇÌÑÇ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      MaskColor       =   &H000000FF&
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000003&
      Caption         =   "ÏÑÈÇÑå í"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  End
End Sub

Private Sub Command2_Click()
   Form1.Hide
   Form2.Show
   'Form2.Height = 9495
   'Form2.Width = 6525
    Form2.Top = 1935
   Form2.Left = 6165
   Form2.BorderStyle = 1
    Form2.Caption = "Enrollment student"
End Sub

Private Sub Command3_Click()
Form1.Hide
Form3.Show

End Sub

Private Sub Form_Load()
  
   Form1.Height = 7725
   Form1.Width = 6525
    Form1.Top = 1935
   Form1.Left = 6165
   Form1.BorderStyle = 1
    Form1.Caption = "Enrollment student"
End Sub

