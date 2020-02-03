VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   6330
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9075
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0CCA
   ScaleHeight     =   6330
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   6000
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List3 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   5040
      Top             =   1080
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "’›ÕÂ ﬁ»·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1900
      Width           =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   4525
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   3650
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   2775
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderStyle     =   2  'Dash
      BorderWidth     =   3
      X1              =   3000
      X2              =   7440
      Y1              =   5500
      Y2              =   5500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      X1              =   3000
      X2              =   3000
      Y1              =   2000
      Y2              =   5500
   End
   Begin VB.Menu khati 
      Caption         =   "khati"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type std
    stdkode As Integer
    firstname As String * 25
    family As String * 15
    flag As Boolean
End Type
Private Type mark
     stdcode As Integer
     unitname As String * 25
     mark As Single
     flag As Boolean
End Type
Private Sub Command1_Click()
  Form6.Hide
  Form3.Show
End Sub

Private Sub Form_Load()
Line2.X2 = Form6.Width
'List1.Height = Form6.Height


End Sub

Private Sub khati_Click()
      List1.Clear
      List2.Clear
      List3.Clear
      List4.Clear
      Dim jaber As std
      Dim javad As mark
      Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
      n = 3500
      m = 4000
      For t = 1 To FileLen("d:\jb.dat") \ Len(jaber)
              z = 0
              j = 0
              Get #1, , jaber
              If jaber.flag = True Then
                       d = jaber.stdkode
                       List1.AddItem d
                       List2.AddItem jaber.firstname
                       List3.AddItem jaber.family
                       Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
                       For g = 1 To FileLen("d:\jv.dat") \ Len(javad)
                               Get #2, , javad
                               If d = javad.stdcode Then
                                      z = z + 1
                                      j = j + javad.mark
                               End If
                      Next
                      If j = 0 And z = 0 Then
                          MsgBox "please insert mark", vbInformation
                            z = 1
                            j = 1
                      End If
                      List4.AddItem j \ z
                      Close #2
                      
                      u = j \ z
                      q = 2000
                      i = 2000
                      If t < 2 Then
                         For v = 20 To 1 Step -1
                          If v = u Then Line (n, q)-(n, q) Else q = q + 175
                         Next
                      End If
                      For X = 20 To 1 Step -1
                            If X = u Then Line -(n, i), QBColor(12) Else i = i + 175
                      Next
                       n = n + 1000
                       m = m + 1000
            End If
   Next
   Close #1
End Sub

Private Sub Timer1_Timer()
    Command1.Top = 0
    Command1.Left = Form6.Width - 2000
    Line2.X2 = Form6.Width
End Sub
