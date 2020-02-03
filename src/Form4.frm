VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6330
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9075
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0CCA
   ScaleHeight     =   6330
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000001&
      Height          =   840
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List3 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   1415
   End
   Begin VB.ListBox List2 
      BackColor       =   &H80000004&
      Height          =   840
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   1415
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
      TabIndex        =   6
      Top             =   120
      Width           =   1650
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   5760
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
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
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
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
      Left            =   2760
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
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
      Height          =   1215
      Left            =   2760
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
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
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   1995
      Width           =   1080
   End
   Begin VB.Menu Averagechart 
      Caption         =   "Average chart"
   End
End
Attribute VB_Name = "Form4"
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
Private Sub Averagechart_Click()
      List1.Clear
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
                      h = j \ z
                      List4.AddItem h
                      Close #2
                       Line (3000, 2000)-(3000, 5500)
                      Line (3000, 5500)-(Form4.Width, 5500)
                      u = j \ z
                      i = 2000
                      k = 4
                      If u = 20 Then k = 5
                      For c = 20 To 1 Step -1
                               If c = k Then k = 0
                               If u = c Then Line (n, i)-(m, 5500), QBColor(c - k), BF Else i = i + 175
                      Next
                      n = n + 1000
                      m = m + 1000
            End If
   Next
   Close #1
End Sub

Private Sub Command1_Click()
     Form4.Hide
     Form3.Show
End Sub

Private Sub Form_Load()
     Form4.Caption = "Enrollment student"
End Sub

Private Sub Timer1_Timer()
      Command1.Top = 0
      Command1.Left = Form4.Width - 2000
End Sub
