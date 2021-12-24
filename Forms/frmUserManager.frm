VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account Manager"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserManager.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Account Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5040
      TabIndex        =   17
      Top             =   1440
      Width           =   7815
      Begin MSComctlLib.ListView lvAccount 
         Height          =   3135
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Database Record"
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Password"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Confirm Password"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "User Previllege"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   3550
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Count: "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3550
         Width           =   915
      End
   End
   Begin prjLMS.jcbutton btnNew 
      Height          =   495
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "New Record"
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "New"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   53
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Input Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4815
      Begin VB.TextBox txtName 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   24
         ToolTipText     =   "Type your Complete Name"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtConfirm 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "l"
         TabIndex        =   23
         ToolTipText     =   "Type the Confirm Password"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "l"
         TabIndex        =   22
         ToolTipText     =   "Type your Password"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   21
         ToolTipText     =   "Type your Username"
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show Password"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cboUsertype 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Select your User Previllege"
         Top             =   1800
         Width           =   2895
      End
      Begin prjLMS.jcbutton btnSave 
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         ToolTipText     =   "Save Data"
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Save"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnDelete 
         Height          =   495
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "Delete Data"
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Delete"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnClear 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Clear fields"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Clear"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnClose 
         Height          =   495
         Left            =   3120
         TabIndex        =   15
         ToolTipText     =   "Close Dialog"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Close"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnUpdate 
         Height          =   495
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Update Data"
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         ButtonStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15199212
         Caption         =   "Update"
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Previllege :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1395
      End
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   360
      Picture         =   "frmUserManager.frx":059A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can add, change, and delete Username and Password "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   4245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Account Manager"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   3660
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmUserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnClear_Click()
txtUsername.Text = ""
txtPassword.Text = ""
txtName.Text = ""
txtConfirm.Text = ""
cboUsertype.Text = ""
btnSave.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False
btnNew.Enabled = True
End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub btnDelete_Click()
    MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from UserAccount where ID=" & Text1.Text & "", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .Delete
        Call RefreshUserAccount
        .Close
    End With
    Set dbRec = Nothing
btnNew.Enabled = True
btnDelete.Enabled = False
btnUpdate.Enabled = False

txtUsername.Text = ""
txtPassword.Text = ""
txtConfirm.Text = ""
txtName.Text = ""
cboUsertype.Text = ""

txtUsername.Enabled = False
txtPassword.Enabled = False
txtConfirm.Enabled = False
txtName.Enabled = False
cboUsertype.Enabled = False

End Sub

Private Sub btnNew_Click()
txtUsername.Enabled = True
txtPassword.Enabled = True
txtConfirm.Enabled = True
txtName.Enabled = True
cboUsertype.Enabled = True

txtUsername.SetFocus
btnNew.Enabled = False
btnSave.Enabled = True
'btnDelete.Enabled = True
'btnUpdate.Enabled = True
End Sub

Private Sub btnSave_Click()
If txtUsername.Text = "" Or txtPassword.Text = "" Or txtConfirm.Text = "" Or txtName.Text = "" Or cboUsertype.Text = "" Then
    MsgBox "Some of your fields is empty. Please complete the information", vbExclamation + vbOKOnly, "Warning"
    txtUsername.SetFocus
ElseIf txtPassword.Text <> txtConfirm.Text Then
    MsgBox "Your password did not match, Please re-enter.", vbExclamation + vbOKOnly, "Warning"
    txtPassword.Text = ""
    txtConfirm.Text = ""
    txtPassword.SetFocus
Else
    MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from UserAccount", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .AddNew
        .Fields("Username") = txtUsername.Text
        .Fields("Password") = txtPassword.Text
        .Fields("ConfirmPassword") = txtConfirm.Text
        .Fields("CompleteName") = txtName.Text
        .Fields("UserType") = cboUsertype.Text
        .Update
        Call RefreshUserAccount
    End With
    dbCon.Close
    'dbRec.Close
    Set dbRec = Nothing
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtConfirm.Text = ""
    txtName.Text = ""
    cboUsertype.Text = ""
    txtUsername.Enabled = False
    txtPassword.Enabled = False
    txtConfirm.Enabled = False
    txtName.Enabled = False
    cboUsertype.Enabled = False
    btnSave.Enabled = False
    btnNew.Enabled = True
End If
End Sub

Private Sub btnUpdate_Click()
    MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from UserAccount where ID=" & Text1.Text & "", dbCon, adOpenDynamic, adLockOptimistic
    With dbRec
        !UserName = txtUsername.Text
        !Password = txtPassword.Text
        !ConfirmPassword = txtConfirm.Text
        !CompleteName = txtName.Text
        !UserType = cboUsertype.Text
        dbRec.Update
        Call RefreshUserAccount
    End With
    Set dbRec = Nothing
btnNew.Enabled = True
btnDelete.Enabled = False
btnUpdate.Enabled = False

txtUsername.Text = ""
txtPassword.Text = ""
txtConfirm.Text = ""
txtName.Text = ""
cboUsertype.Text = ""

txtUsername.Enabled = False
txtPassword.Enabled = False
txtConfirm.Enabled = False
txtName.Enabled = False
cboUsertype.Enabled = False

End Sub

Private Sub chkShow_Click()
If chkShow.Value = 1 Then
    txtPassword.PasswordChar = ""
    txtConfirm.PasswordChar = ""
    txtPassword.Font = "MS Sans Serif"
    txtConfirm.Font = "MS Sans Serif"
Else
    txtPassword.Font = "Wingdings"
    txtConfirm.Font = "Wingdings"
    txtPassword.PasswordChar = "l"
    txtConfirm.PasswordChar = "l"
End If
End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

btnSave.Enabled = False
btnDelete.Enabled = False
btnUpdate.Enabled = False

Call RefreshUserAccount


Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserPrev", dbCon, 3, 3
Do While Not dbRec.EOF
    cboUsertype.AddItem dbRec!UserPrev
dbRec.MoveNext
Loop

txtUsername.Enabled = False
txtPassword.Enabled = False
txtConfirm.Enabled = False
txtName.Enabled = False
cboUsertype.Enabled = False

'AltLVBackground lvAccount, vbWhite, &H80000002
DisableClose Me.hwnd
End Sub

Private Sub RefreshUserAccount()
Dim dbRec As ADODB.Recordset

lvAccount.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserAccount Order by Username", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvAccount.ListItems.Add , , !ID, 1, 1
                lvAccount.ListItems(lvAccount.ListItems.Count).SubItems(1) = "" & !UserName
                lvAccount.ListItems(lvAccount.ListItems.Count).SubItems(2) = "" & !Password
                lvAccount.ListItems(lvAccount.ListItems.Count).SubItems(3) = "" & !ConfirmPassword
                lvAccount.ListItems(lvAccount.ListItems.Count).SubItems(4) = "" & !CompleteName
                lvAccount.ListItems(lvAccount.ListItems.Count).SubItems(5) = "" & !UserType
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing
Label9.Caption = lvAccount.ListItems.Count
End Sub

Private Sub lvAccount_Click()
btnUpdate.Enabled = True
btnDelete.Enabled = True
btnNew.Enabled = False

On Error Resume Next
Text1.Text = lvAccount.SelectedItem
txtUsername.Text = lvAccount.SelectedItem.SubItems(1)
txtPassword.Text = lvAccount.SelectedItem.SubItems(2)
txtConfirm.Text = lvAccount.SelectedItem.SubItems(3)
txtName.Text = lvAccount.SelectedItem.SubItems(4)
cboUsertype.Text = lvAccount.SelectedItem.SubItems(5)
txtUsername.Enabled = True
txtPassword.Enabled = True
txtConfirm.Enabled = True
txtName.Enabled = True
cboUsertype.Enabled = True

End Sub

Private Sub lvAccount_LostFocus()
'btnUpdate.Enabled = False
'btnDelete.Enabled = False
'btnNew.Enabled = True
'txtUsername.Text = ""
'txtPassword.Text = ""
'txtConfirm.Text = ""
'txtName.Text = ""
'cboUsertype.Text = ""

End Sub

Private Sub AltLVBackground(lv As ListView, _
    ByVal BackColorOne As OLE_COLOR, _
    ByVal BackColorTwo As OLE_COLOR)
'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
Dim lh      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
    With lvAccount
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = Me.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lh = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lh * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lh)-(.ScaleWidth, lh * 2), BackColorTwo, BF
                Set lvAccount.Picture = .Image
            End With
            Set picAlt = Nothing
            Me.Controls.Remove "picAlt"
            lvAccount.Parent.ScaleMode = lSM
        End If
    End With
End Sub

