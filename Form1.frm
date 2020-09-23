VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Ascii"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00000000&
      Height          =   6420
      Left            =   0
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   6360
      ScaleWidth      =   2550
      TabIndex        =   19
      Top             =   0
      Width           =   2610
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   1440
         Top             =   1080
      End
      Begin VB.PictureBox Pictray 
         Height          =   495
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSRDC.MSRDC MSRDC1 
         Height          =   330
         Left            =   600
         Top             =   5520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         _Version        =   393216
         Options         =   0
         CursorDriver    =   0
         BOFAction       =   0
         EOFAction       =   0
         RecordsetType   =   1
         LockType        =   3
         QueryType       =   0
         Prompt          =   3
         Appearance      =   1
         QueryTimeout    =   30
         RowsetSize      =   100
         LoginTimeout    =   15
         KeysetSize      =   0
         MaxRows         =   0
         ErrorThreshold  =   -1
         BatchSize       =   15
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         ReadOnly        =   0   'False
         Appearance      =   -1  'True
         DataSourceName  =   "Ascii"
         RecordSource    =   "SELECT * FROM [AsciiTable] ORDER BY [DEC]"
         UserName        =   "Admin"
         Password        =   "vannelle"
         Connect         =   ""
         LogMessages     =   ""
         Caption         =   "MSRDC1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image ImTray 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":B9D7
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   6420
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Bindings        =   "Form1.frx":BCE1
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   258
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         AllowBigSelection=   0   'False
         TextStyleFixed  =   1
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   3360
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   -75
         Width           =   3255
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   4
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   12
            ToolTipText     =   "Search for CODE value"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   11
            ToolTipText     =   "Search for CHARACTER value"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   10
            ToolTipText     =   "Search for HEXDICIMAL value"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   720
            MaxLength       =   3
            TabIndex        =   9
            ToolTipText     =   "Search for DECIMAL value"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   2
            TabIndex        =   8
            ToolTipText     =   "Search for CONTROL value"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "CTRLVAL"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   390
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "DECVAL"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   390
            Index           =   1
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "HEXVAL"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   390
            Index           =   2
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "CHARVAL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   13.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   390
            Index           =   3
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            DataField       =   "CODEVAL"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   390
            Index           =   4
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "CODE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "CHAR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "HEX"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "DEC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "CTRL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Menu f 
      Caption         =   "File"
      Begin VB.Menu OnTop 
         Caption         =   "OnTop"
      End
      Begin VB.Menu min 
         Caption         =   "Tray"
      End
      Begin VB.Menu streepken 
         Caption         =   "-"
      End
      Begin VB.Menu Stop 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu show 
         Caption         =   "Show"
      End
      Begin VB.Menu streepken2 
         Caption         =   "-"
      End
      Begin VB.Menu TheEnd 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' TrayIcon
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA



Private Sub About_Click()
Me.Hide
Load Form3
End Sub

Private Sub Command1_Click()
Text3.Text = StrConv(Text3.Text, vbUnicode)
End Sub

Private Sub Form_Initialize()

Dim i As Long
Flex.ColWidth(5) = 0
For i = 0 To 4

Flex.ColAlignment(i) = 3
Flex.ColWidth(i) = 1000
DoEvents

Next i

Me.show: DoEvents
Timer1.Enabled = True

End Sub
Private Sub cmdCreate_Click()
MsgBox "clicked"
End Sub
Private Sub Form_Load()

Call SetNumber(Text2(1), True)
Call SetUcase(Text2(4), True)
Call SetUcase(Text2(2), True)
RemoveMenus Me, True, False, True, False, True, True, True



End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width <> 8745 Then Me.Width = 8745
If Me.Height <> 7230 Then Me.Height = 7230
End Sub

Private Sub min_Click()
    Dim Cancel
    Cancel = Not ReadyToClose
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = Pictray.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = ImTray.Picture
    TrayI.szTip = "Ascii -By sOul- " & Chr(169) & " 2000" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, TrayI
    Me.Hide
End Sub

Private Sub OnTop_Click()

If OnTop.Checked = False Then
OnTop.Checked = True
    res% = SetWindowPos(Form1.hWnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS)
Else
OnTop.Checked = False
    res% = SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, _
    0, 0, 0, 0, FLAGS)
End If

End Sub

Private Sub Pictray_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / Screen.TwipsPerPixelX
If msg = WM_RBUTTONDOWN Or msg = WM_LBUTTONDOWN Then
        Me.PopupMenu TrayMenu
End If
End Sub

Private Sub show_Click()
Me.show
End Sub

Private Sub Stop_Click()
MSRDC1.Connection.Close

If FileExist(App.Path & "\AsciiTable.mdb") = True Then
    Kill App.Path & "\AsciiTable.mdb"
End If
Call DeleteAccessDSN("Ascii")
Shell_NotifyIcon NIM_DELETE, TrayI

Unload Me
End Sub

Private Sub Text1_Change(Index As Integer)

If Index = 3 Then

End If
End Sub

Private Sub Text2_Change(Index As Integer)

Select Case Index

Case 0
    Call SearchInFlex(0, Text2(Index).Text)
Case 1
    Call SearchInFlex(1, Text2(Index).Text)
Case 2
    Call SearchInFlex(2, Text2(Index).Text)
Case 3
    Call SearchInFlex(3, Text2(Index).Text)
Case 4
    Call SearchInFlex(4, Text2(Index).Text)

End Select

End Sub

Private Sub Text2_LostFocus(Index As Integer)
Text2(Index).Text = ""
End Sub

Private Sub TheEnd_Click()
Stop_Click
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

On Error GoTo Blank
StatusBar1.SimpleText = (Mid(Flex.TextArray(((1 * Flex.Row) * Flex.Cols) + 5), 1, 1)) & _
LCase(Mid(Flex.TextArray(((1 * Flex.Row) * Flex.Cols) + 5), 2, Len(Flex.TextArray(((1 * Flex.Row) * Flex.Cols) + 5)) - 1))
GoTo Good
Blank:
StatusBar1.SimpleText = ""
Good:

Text1(0).Text = Flex.Text
For i = 1 To 4
    If i = 3 Then
    Text1(i).Text = Chr(Flex.TextArray(((1 * Flex.Row) * Flex.Cols) + 1))
    Else
    Text1(i).Text = Flex.TextArray(((1 * Flex.Row) * Flex.Cols) + i)
    End If
Next i

End Sub
Private Sub SearchInFlex(Col As Integer, SearchString As String)
Dim C As Integer: Dim A As String

If SearchString = "" Then Call SelectRow(1)

 For C = 1 To Flex.Rows - 1
    A = Flex.TextMatrix(C, Col)
    
    If A = SearchString Then
   
      Call SelectRow(C)

    End If
 
 Next C

End Sub
Private Sub SelectRow(C As Integer)

With Flex
        .TopRow = C
        .AllowBigSelection = False
        .SelectionMode = flexSelectionByRow
        .Redraw = False
        .Row = C
        .Col = 0
        .ColSel = .Cols - 1
        .Redraw = True
End With

End Sub

Private Sub Timer2_Timer()
   If IsWindowVisible(hWnd) Then
     Shell_NotifyIcon NIM_DELETE, TrayI
   End If
End Sub


