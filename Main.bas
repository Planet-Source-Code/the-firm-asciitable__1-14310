Attribute VB_Name = "MainMod"

Dim MyDsn As String
Sub Main()
On Error GoTo ErrorHandler
Dim Data As String

'Check if this program is already running!

If App.PrevInstance = True Then
    MsgBox ("The program is already running!")
    For i = 0 To Form2.List1.ListCount
         If Form2.List1.ListCount = 0 Then Exit Sub
        If Form2.List1.List(i) = "Ascii" Then
            Call pSetForegroundWindow(Form2.List1.ItemData(i))
        End If
    Next i
    End
End If
Form2.show
DoEvents

'Delete Database if Exist
If FileExist(App.Path & "\AsciiTable.mdb") = True Then
    Kill App.Path & "\AsciiTable.mdb"
End If

'Delete old DNS-settings
Call DeleteAccessDSN("Ascii")

'Download Database to App.path
Call DownloadFile(101, App.Path & "\AsciiTable.mdb")
        
MyDsn = "Ascii"
A = checkWantedAccessDSN(MyDsn) ' Dns_module
    
    If A = True Then

         'If Exist

    Else

        ' If Not, Make it....!!

        Call createAccessDSN(App.Path & "\AsciiTable.mdb", "Ascii")

    End If
Load Form1
Unload Form2
GoTo Einde
ErrorHandler:
MsgBox Err.Number & Chr(13) & Err.Description, vbCritical, "Error"

Einde:
End Sub
Function FileExist(FileName As String) As Boolean

If Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = "" Then
 FileExist = False
Else
 FileExist = True
End If

End Function
Function DownloadFile(Nr As Integer, DestFileName As String) As Boolean
Dim mFreeFile As Integer: mFreeFile = FreeFile
Dim TmpDir As String: TmpDir = App.Path & "\tmp~"
Dim Buffer As String
Dim Data As Variant, B As Long
Data = LoadResData(Nr, "Custom")


Open TmpDir For Binary As mFreeFile
        Put mFreeFile, , Data
Close mFreeFile

B = FileLen(TmpDir)
Buffer = String(B - 12, " ")

Open TmpDir For Binary As mFreeFile
    Seek mFreeFile, 13
    Get mFreeFile, , Buffer
Close mFreeFile

Kill TmpDir

Open DestFileName For Binary As mFreeFile
    Put mFreeFile, , Buffer
Close mFreeFile

End Function
