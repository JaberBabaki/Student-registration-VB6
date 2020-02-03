VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9150
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0CCA
   ScaleHeight     =   9150
   ScaleWidth      =   10830
   Begin VB.ListBox List5 
      Height          =   5580
      Left            =   8280
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   5580
      Left            =   6120
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   5580
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   5580
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   600
      Top             =   5160
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
      Height          =   495
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8520
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   5580
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄œ·"
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   120
      X2              =   10800
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "‰„—Â"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ œ—” "
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "—„“ ⁄»Ê—"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu Result 
      Caption         =   "Result"
   End
End
Attribute VB_Name = "Form5"
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
Form5.Hide
Form3.Show
End Sub

Private Sub Form_Load()
     Form5.Caption = "Enrollment student"
     Line1.X2 = Form5.Width
End Sub

Private Sub Result_Click()
     List1.Clear
     List2.Clear
     List3.Clear
     List4.Clear
     List5.Clear
     Dim jaber As std
     Dim javad As mark
     Dim a As Integer
     Dim b As Integer
     Dim c As Integer
     Dim d As Integer
     a = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
     b = 0
     Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
     For e = 1 To FileLen("d:\jb.dat") \ Len(jaber)
             c = 0
             d = 0
             Get #1, , jaber
             If a = jaber.stdkode And jaber.flag = True Then
                     Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
                     For f = 1 To FileLen("d:\jv.dat") \ Len(javad)
                              Get #2, , javad
                              If a = javad.stdcode Then
                                    c = c + 1
                                    d = d + javad.mark
                                    List1.AddItem c
                                    List2.AddItem javad.stdcode
                                    List3.AddItem javad.unitname
                                    List4.AddItem javad.mark
                              End If
                     Next
                     If c = 0 And d = 0 Then
                            MsgBox "please insert mark", vbInformation
                            c = 1
                            d = 1
                      End If
                     List5.AddItem d / c
                     
                     Close #2
           Else
                     b = b + 1
           End If
   Next
   If b = FileLen("d:\jb.dat") \ Len(jaber) Then MsgBox "this std has not  been registred or std delete", vbInformation + vbOKOnly, "«Œÿ«—"
   
   Close #1
  
End Sub

Private Sub Timer1_Timer()
     Command1.Left = 0
     Command1.Top = Form5.Height - 1500
End Sub

