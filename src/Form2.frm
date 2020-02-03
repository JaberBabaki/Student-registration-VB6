VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0CCA
   ScaleHeight     =   8955
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "’›ÕÂ ﬁ»·"
      Height          =   615
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form1.Show
End Sub

