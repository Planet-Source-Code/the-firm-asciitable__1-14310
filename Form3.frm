VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6690
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   -240
      ScaleHeight     =   3675
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         MouseIcon       =   "Form3.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         MouseIcon       =   "Form3.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6360
      MouseIcon       =   "Form3.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
        Picture1.Picture = Form2.Image1.Picture
        Unload Form2
        Me.show
        DoEvents
        Me.Cls
        Me.Font.Size = 12
        
        m_bDoEffect = True
        TextEffect "AsciiTable Version 1.0", 5, 0, False, 50, -1, &HFFFF&, &H808080       'RGB(&H80, 0, 0)

        Me.Font.Size = 9
        m_bDoEffect = True
        TextEffect "Programmed by sOul-", 5, 17, False, 100, -1, &HE0E0E0, &H808080      'RGB(&H80, 0, 0)
  
   

        Label3.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFFFF
End Sub

Private Sub Label1_Click()
Call ExecuteLink("http://users.belgacom.net/soul/index.html")
End Sub

Private Sub Label2_Click()
Call ExecuteLink("mailto:angeldust@be.packardbell.org")
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFFFF
Label2.ForeColor = &HFFFFFF
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFF&
Label2.ForeColor = &HFFFFFF
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFF&
Label1.ForeColor = &HFFFFFF
End Sub
Public Sub TextEffect( _
ByVal sText As String, _
ByVal lX As Long, ByVal lY As Long, _
Optional ByVal bLoop As Boolean = False, _
Optional ByVal lStartSpacing As Long = 128, _
Optional ByVal lEndSpacing As Long = -1, _
Optional ByVal oColor As OLE_COLOR = vbWindowText, _
Optional ByVal bcolor As OLE_COLOR)


Dim i As Long
Dim X As Long
Dim lLen As Long
Dim lHDC As Long
Dim hBrush As Long
Static tR As RECT
Dim iDir As Long
Dim bNotFirstTime As Boolean
Dim lTime As Long
Dim lIter As Long
Dim bSlowDown As Boolean
Dim lCOlor As Long
Dim bDoIt As Boolean


    iDir = -1
    i = lStartSpacing
    tR.Left = lX: tR.Top = lY: tR.Right = lX: tR.Bottom = lY
    OleTranslateColor oColor, 0, lCOlor
   
    
    hBrush = CreateSolidBrush(bcolor)    'GetSysColor(COLOR_BTNFACE))
    lLen = Len(sText)
    lHDC = Me.hdc
    SetTextColor lHDC, lCOlor
    bDoIt = True
    
    Do While m_bDoEffect And bDoIt
        lTime = timeGetTime
        If (i < -3) And Not (bLoop) And Not (bSlowDown) Then
            bSlowDown = True
            iDir = 1
            lIter = (i + 4)
        End If
        If (i > 128) Then iDir = -1
        If Not (bLoop) And iDir = 1 Then
            If (i = lEndSpacing) Then
                ' Stop
                bDoIt = False
            Else
                lIter = lIter - 1
                If (lIter <= 0) Then
                    i = i + iDir
                    lIter = (i + 4)
                End If
            End If
        Else
            i = i + iDir
        End If
        FillRect lHDC, tR, hBrush
        X = 32 - (i * lLen)
        SetTextCharacterExtra lHDC, i
        DrawText lHDC, sText, lLen, tR, DT_CALCRECT
        tR.Right = tR.Right + 4
        If (tR.Right > Me.ScaleWidth \ Screen.TwipsPerPixelX) Then tR.Right = Me.ScaleWidth \ Screen.TwipsPerPixelX
        DrawText lHDC, sText, lLen, tR, DT_LEFT
        Me.Refresh
        Do
           
            DoEvents
        Loop While (timeGetTime - lTime) < 20
    Loop
    DeleteObject hBrush

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    m_bDoEffect = False
End Sub

Private Sub Label3_Click()
Form1.show
Unload Form3
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HC000&
End Sub

