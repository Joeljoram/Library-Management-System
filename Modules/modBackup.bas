Attribute VB_Name = "modBackup"
Option Explicit

Public Function DoesFileExist(PathName As String) As Boolean

If Dir$(PathName) <> vbNullString Then
    'file exists
    DoesFileExist = True
Else
    'file doesn't exist
    DoesFileExist = False
End If

End Function

Public Function MDbackupdatabases() As Long
Dim strPath, strBackup As String
Dim sbakfile As String
Dim FSO As FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")
'Call DoesFileExist

If DoesFileExist(App.Path & "\Backup\data-" & Format(Date, "dd-mm-yyyy") & ".mdb") Then
MsgBox "The File Exists", vbCritical, "File Exist"
frmBackUp.Timer2.Enabled = False
Exit Function
Else
strPath = App.Path & "\Database\data.mdb"
strBackup = App.Path & "\Backup\data-" & Format(Date, "dd-mm-yyyy") & ".mdb"
FSO.CopyFile strPath, strBackup
frmBackUp.Timer2.Enabled = True
End If


End Function



