Attribute VB_Name = "modCon"
'Option Explicit
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const MF_BYPOSITION = &H400&

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
ByVal bRevert As Long) As Long
Global dbCon As ADODB.Connection
Global dbRec As ADODB.Recordset
Global dbRec2 As ADODB.Recordset
Global dbRec3 As ADODB.Recordset
Public UserLog As String
Public Squery As String





Public Sub Connected()
Set dbCon = New ADODB.Connection
dbCon.CursorLocation = adUseClient
dbCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\data.mdb;Persist Security Info=False"
dbCon.Open

End Sub


'Dim OpenDB As Database
'Dim OpenRec As Recordset
'Dim sRole As String



'Private Sub cmdOK_Click()





'Set OpenRec = OpenDB.OpenRecordset("UserAccount")
'While Not OpenRec.EOF
'If OpenRec!UserName = txtUsername.Text And OpenRec!Password = txtPassword.Text Then
'sRole = OpenRec!Role
'End If
'OpenRec.MoveNext
'Wend
'If sRole = "Administrator" Then
'Unload Me
'Form1.Show
'Form1.Command1.Enabled = False


'Else
'Unload Me
'Form1.Show
'Form1.Command1.Enabled = True
'End If
'End Sub

'Private Sub Form_Load()
'Module1.Connected
 'Set OpenDB = OpenDatabase("C:\Users\harvey\Desktop\dblogin2003.mdb")
 'End Sub




Public Sub DisableClose(hwnd As Long)
Dim hMenu As Long
hMenu = GetSystemMenu(hwnd, 0)
RemoveMenu hMenu, 6, MF_BYPOSITION
RemoveMenu hMenu, 5, MF_BYPOSITION
End Sub
'Then, to disable the Close button on a form, use
'DisableClose (Form.hWnd)
