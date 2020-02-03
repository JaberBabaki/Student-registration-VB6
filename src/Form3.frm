VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   9090
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   12510
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0CCA
   ScaleHeight     =   9090
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "’›ÕÂ «’·Ì"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Menu Dataentry 
      Caption         =   "Data entry"
      Begin VB.Menu Register 
         Caption         =   "Register"
         Shortcut        =   ^R
      End
      Begin VB.Menu Markentry 
         Caption         =   "Mark entry"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu View 
      Caption         =   "View"
      Begin VB.Menu Result 
         Caption         =   "Result"
      End
      Begin VB.Menu Averagechart 
         Caption         =   "Average chart"
      End
      Begin VB.Menu khati 
         Caption         =   "khati"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Attribute 
         Caption         =   "Attribute"
      End
      Begin VB.Menu Marks 
         Caption         =   "Marks"
      End
   End
   Begin VB.Menu Delete 
      Caption         =   "Delete"
      Begin VB.Menu Phisical 
         Caption         =   "Phisical"
      End
      Begin VB.Menu Logical 
         Caption         =   "Logical"
      End
      Begin VB.Menu Restore 
         Caption         =   "Restore"
      End
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Attribute_Click()
     Dim jaber As std
     Dim javad As mark
     Dim m As Integer
     Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
     h = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
     m = 0
     For w = 1 To FileLen("d:\jb.dat") \ Len(jaber)
         Get #1, , jaber
         If h = jaber.stdkode And jaber.flag = True Then
             jaber.firstname = InputBox("·ÿ›« ‰«„ —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ ‰«„", , 5000, 2000)
             jaber.family = InputBox("·ÿ›« ‰«„ Œ«‰Ê«œêÌ —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ ‰«„ Œ«‰Ê«œêÌ", , 5000, 2000)
             Put #1, w, jaber
         Else
             m = m + 1
         End If
     Next
     If m = FileLen("d:\jb.dat") \ Len(jaber) Then MsgBox "this std has  been registred", vbInformation + vbOKOnly, "«Œÿ«—"
     Close #1
End Sub

Private Sub Averagechart_Click()
Form3.Hide
Form4.Show
End Sub

Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Form3.Caption = "Enrollment student"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
      PopupMenu Dataentry
End If
End Sub

Private Sub khati_Click()
Form3.Hide
Form6.Show
End Sub

Private Sub Logical_Click()
      Dim jaber As std
      Dim javad As mark
      Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
      Y = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
      d = 0
      For i = 1 To FileLen("d:\jb.dat") \ Len(jaber)
              Get #1, , jaber
              If Y = jaber.stdkode Then
                     jaber.flag = False
                     Put #1, i, jaber
                     Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
                     For e = 1 To FileLen("d:\jv.dat") \ Len(javad)
                               Get #2, , javad
                               If Y = javad.stdcode Then
                                      javad.flag = False
                                      Put #2, e, javad
                               End If
                     Next
                     Close #2
              Else
                     d = d + 1
              End If
     Next
     If d = FileLen("d:\jv.dat") \ Len(jaber) Then MsgBox "Eraseing Information", vbInformation + vbOKOnly, "«Œÿ«—"
     Close #1
     
End Sub

Private Sub Markentry_Click()
     Dim jaber As std
     Dim javad As mark
     Dim a As Integer
     Dim b As Integer
     Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
     a = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
     b = 0
     Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
     For c = 1 To FileLen("d:\jb.dat") \ Len(jaber)
             Get #1, , jaber
             If a = jaber.stdkode And jaber.flag = True Then
                     javad.stdcode = a
                     javad.unitname = InputBox("·ÿ›« ‰«„ œ— —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ ‰«„ œ—”", , 5000, 2000)
                     javad.mark = InputBox("·ÿ›« ‰„—Â —«Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ ‰„—Â ", , 5000, 2000)
                     javad.flag = True
                     Seek #2, FileLen("d:\jv.dat") \ Len(javad) + 1
                     Put #2, , javad
             Else
                     b = b + 1
             End If
     Next
     If b = FileLen("d:\jb.dat") \ Len(jaber) Then MsgBox "this std has not been registred", vbInformation + vbOKOnly, "«Œÿ«—"
     Close #2
     Close #1
End Sub

Private Sub Marks_Click()
      Dim jaber As std
      Dim javad As mark
      Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
      b = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
      s = 0
      For w = 1 To FileLen("d:\jb.dat") \ Len(jaber)
              Get #1, , jaber
              If b = jaber.stdkode And jaber.flag = True Then
                      Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
                      For r = 1 To FileLen("d:\jv.dat") \ Len(javad)
                                   Get #2, , javad
                                   If b = javad.stdcode Then
                                          javad.unitname = InputBox("·ÿ›« ‰«„ œ—” ÃœÌœ —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ ‰«„ œ—” ÃœÌœ", , 5000, 2000)
                                          javad.mark = InputBox("·ÿ›« ‰„—Â ÃœÌœ —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ ‰„— ÃœÌœ", , 5000, 2000)
                                          Put #2, r, javad
                                   End If
                      Next
              Else
                     s = s + 1
              End If
              Close #2
              
      Next
      If s = FileLen("d:\jb.dat") \ Len(jaber) Then MsgBox "this std has not been registred", vbInformation + vbOKOnly, "«Œÿ«—"
      Close #1
End Sub

Private Sub Phisical_Click()
      Dim jaber As std
      Dim javad As mark
      b = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
      Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
      For i = 1 To FileLen("d:\jb.dat") \ Len(jaber)
          Get #1, , jaber
          If b = jaber.stdkode Then
                 j = MsgBox("are you erasing data?", vbYesNo + vbQuestion, "Erase Data")
          Else
                 s = s + 1
          End If
      Next
      If s = FileLen("d:\jb.dat") \ Len(jaber) Then MsgBox "this std has  been registred", vbInformation + vbOKOnly, "«Œÿ«—"
      Close #1
      g = False
      If j = vbYes Then
             Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
             Open ("d:\jb1.dat") For Random As #3 Len = Len(jaber)
             For r = 1 To FileLen("d:\jb.dat") \ Len(jaber)
                     Get #1, , jaber
                     If b <> jaber.stdkode Then
                             Put #3, , jaber
                     Else
                             g = True
                     End If
            Next
            Close #1
            Close #3
            If g = True Then
                  Kill "d:\jb.dat"
                  Name "d:\jb1.dat" As "d:\jb.dat"
           End If
           Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
           Open ("d:\jv1.dat") For Random As #4 Len = Len(javad)
           For r = 1 To FileLen("d:\jv.dat") \ Len(javad)
                   Get #2, , javad
                   If b <> javad.stdcode Then
                           Put #4, , javad
                   Else
                           k = True
                   End If
          Next
          Close #2
          Close #4
          If k = True Then
                 Kill "d:\jv.dat"
                 Name "d:\jv1.dat" As "d:\jv.dat"
                 MsgBox "Eraseing Information", vbInformation
          End If
     End If
End Sub

Private Sub Register_Click()
     Dim jaber As std
     Dim a As Boolean
     Dim b As Integer
     Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
     a = False
     b = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
     For c = 1 To FileLen("d:\jb.dat") \ Len(jaber)
          Get #1, , jaber
          If b = jaber.stdkode Then a = True
     Next
     If a = True Then
         MsgBox "this std has  been registred", vbInformation + vbOKOnly, "«Œÿ«—"
     Else
         jaber.firstname = InputBox("·ÿ›« ‰«„ —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ‰«„", , 5000, 2000)
         jaber.stdkode = b
         jaber.family = InputBox("·ÿ›« ‰«„ Œ«‰Ê«œêÌ —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ ‰«„ Œ«‰Ê«œêÌ", , 5000, 2000)
         jaber.flag = True
         Seek #1, FileLen("d:\jb.dat") \ Len(jaber) + 1
         Put #1, , jaber
     End If
     Close
     
End Sub

Private Sub Restore_Click()
      Dim jaber As std
      Dim javad As mark
      Open ("d:\jb.dat") For Random As #1 Len = Len(jaber)
      Y = InputBox("·ÿ›« —„“ ⁄»Ê— —« Ê«—œ ò‰Ìœ", "ò«œ— Ê—Êœ —„“", , 5000, 2000)
      d = 0
      For i = 1 To FileLen("d:\jb.dat") \ Len(jaber)
              Get #1, , jaber
              If Y = jaber.stdkode Then
                     jaber.flag = True
                     Put #1, i, jaber
                     Open ("d:\jv.dat") For Random As #2 Len = Len(javad)
                     For e = 1 To FileLen("d:\jv.dat") \ Len(javad)
                               Get #2, , javad
                               If Y = javad.stdcode Then
                                      javad.flag = True
                                      Put #2, e, javad
                               End If
                     Next
                     Close #2
              Else
                     d = d + 1
              End If
     Next
     If d = FileLen("d:\jv.dat") \ Len(jaber) Then MsgBox "this std has not  been registred", vbInformation + vbOKOnly, "«Œÿ«—"
     Close #1
End Sub

Private Sub Result_Click()
     Form3.Hide
     Form5.Show
End Sub
