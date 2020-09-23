Attribute VB_Name = "TextBoxFilter"
Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Private Const GWL_USERDATA = (-21)
Private Const GWL_STYLE = (-16)

' Edit Control Styles
Const ES_AUTOHSCROLL = &H80&
Const ES_AUTOVSCROLL = &H40&
Const ES_CENTER = &H1&
Const ES_LEFT = &H0&
Const ES_LOWERCASE = &H10&
Const ES_MULTILINE = &H4&
Const ES_NUMBER = &H2000&
Const ES_NOHIDESEL = &H100&
Const ES_OEMCONVERT = &H400&
Const ES_PASSWORD = &H20&
Const ES_READONLY = &H800&
Const ES_RIGHT = &H2&
Const ES_UNICODE = &H0&
Const ES_UPPERCASE = &H8&
Const ES_WANTRETURN = &H1000&

' Edit Control Notification Codes
Const EN_SETFOCUS = &H100
Const EN_KILLFOCUS = &H200
Const EN_CHANGE = &H300
Const EN_UPDATE = &H400
Const EN_ERRSPACE = &H500
Const EN_MAXTEXT = &H501
Const EN_HSCROLL = &H601
Const EN_VSCROLL = &H602

' Edit Control Messages
Const EM_GETSEL = &HB0
Const EM_SETSEL = &HB1
Const EM_GETRECT = &HB2
Const EM_SETRECT = &HB3
Const EM_SETRECTNP = &HB4
Const EM_SCROLL = &HB5
Const EM_LINESCROLL = &HB6
Const EM_SCROLLCARET = &HB7
Const EM_GETMODIFY = &HB8
Const EM_SETMODIFY = &HB9
Const EM_GETLINECOUNT = &HBA
Const EM_LINEINDEX = &HBB
Const EM_SETHANDLE = &HBC
Const EM_GETHANDLE = &HBD
Const EM_GETTHUMB = &HBE
Const EM_LINELENGTH = &HC1
Const EM_REPLACESEL = &HC2
Const EM_GETLINE = &HC4
Const EM_LIMITTEXT = &HC5
Const EM_CANUNDO = &HC6
Const EM_UNDO = &HC7
Const EM_FMTLINES = &HC8
Const EM_LINEFROMCHAR = &HC9
Const EM_SETTABSTOPS = &HCB
Const EM_SETPASSWORDCHAR = &HCC
Const EM_EMPTYUNDOBUFFER = &HCD
Const EM_GETFIRSTVISIBLELINE = &HCE
Const EM_SETREADONLY = &HCF
Const EM_SETWORDBREAKPROC = &HD0
Const EM_GETWORDBREAKPROC = &HD1
Const EM_GETPASSWORDCHAR = &HD2

Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
Dim curstyle As Long
Dim newstyle As Long

curstyle = GetWindowLong(NumberText.hWnd, GWL_STYLE)

If Flag Then
   curstyle = curstyle Or ES_NUMBER
Else
   curstyle = curstyle And (Not ES_NUMBER)
End If

newstyle = SetWindowLong(NumberText.hWnd, GWL_STYLE, curstyle)
NumberText.Refresh
End Sub

Public Sub SetUcase(NumberText As TextBox, Flag As Boolean)
Dim curstyle As Long
Dim newstyle As Long

curstyle = GetWindowLong(NumberText.hWnd, GWL_STYLE)

If Flag Then
   curstyle = curstyle Or ES_UPPERCASE
Else
   curstyle = curstyle And (Not ES_UPPERCASE)
End If

newstyle = SetWindowLong(NumberText.hWnd, GWL_STYLE, curstyle)
NumberText.Refresh
End Sub

