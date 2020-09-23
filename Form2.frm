VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   3675
      Left            =   -360
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   7395
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call pGetTasks
End Sub
Private Sub pGetTasks()
    Call fEnumWindows(List1)
    List1.ListIndex = 0
End Sub
