Attribute VB_Name = "RemoveMenu"
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public ReadyToClose As Boolean

Public Sub RemoveMenus(Frm As Form, remove_restore As Boolean, remove_move As Boolean, _
    remove_size As Boolean, remove_minimize As Boolean, remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)

Dim hMenu As Long
Dim hwnd As Long: hwnd = Frm.hwnd
hMenu = GetSystemMenu(hwnd, False)

If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION

End Sub

