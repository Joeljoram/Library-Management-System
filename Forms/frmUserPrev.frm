VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPrev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Previllege"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin prjLMS.jcbutton btnClear 
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "New Record"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.TextBox txtUsertype 
      Height          =   360
      Left            =   120
      MaxLength       =   20
      TabIndex        =   4
      ToolTipText     =   "Type your User Previllege"
      Top             =   3285
      Width           =   5655
   End
   Begin prjLMS.jcbutton btnNew 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "New Record"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSComctlLib.ListView lvUsertype 
      Height          =   2655
      Left            =   40
      TabIndex        =   3
      ToolTipText     =   "Database Record"
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Previllege"
         Object.Width           =   2540
      EndProperty
   End
   Begin prjLMS.jcbutton btnSave 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Save Data"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
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
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Close Dialog"
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
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
   Begin prjLMS.jcbutton btnDelete 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "Delete Data"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
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
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Add && Delete User Previllege "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4680
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
            Picture         =   "frmUserPrev.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "frmUserPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()

txtUsertype.Text = ""
txtUsertype.SetFocus
btnDelete.Enabled = False
btnSave.Enabled = False
btnNew.Enabled = True

End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show

End Sub

Private Sub btnDelete_Click()
On Error Resume Next
If MsgBox("Are you sure you want to delete this record: " & txtUsertype.Text & " ?", vbYesNo + vbQuestion, "Delete") = vbYes Then
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from UserPrev where UserID=" & Text1.Text, dbCon, adOpenDynamic, adLockOptimistic
    With dbRec
        .Fields("UserID") = Text1.Text
        .Fields("UserPrev") = txtUsertype.Text
        .Delete
        .Update
        Call RefreshUserPrev
    End With
    MsgBox "Record Successfully Deleted.", vbInformation, "Delete Success"
    btnNew.Enabled = True
    btnDelete.Enabled = False
    Set dbRec = Nothing
Else
Text1.Text = ""
txtUsertype.Text = ""
Exit Sub
End If

End Sub

Private Sub btnNew_Click()
txtUsertype.SetFocus
btnNew.Enabled = False
btnSave.Enabled = True


End Sub

Private Sub btnSave_Click()

If IsNumeric(txtUsertype.Text) = True Then
    MsgBox "User Previllege must be a word value.", vbExclamation + vbOKOnly, "Warning"
    txtUsertype.Text = ""
    txtUsertype.SetFocus
ElseIf txtUsertype.Text = "" Then
    MsgBox "The box is empty.", vbExclamation + vbOKOnly, "Warning"
    txtUsertype.SetFocus
Else
    MsgBox "Record Successfully Saved!", vbInformation + vbOKOnly, "Success Saved!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from UserPrev", dbCon, 3, 3
    With dbRec
        .AddNew
        .Fields("UserPrev") = txtUsertype.Text
        .Update
        Call RefreshUserPrev
    End With
    Set dbRec = Nothing
    btnSave.Enabled = False
    btnNew.Enabled = True
    txtUsertype.Text = ""
End If

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

btnSave.Enabled = False
btnDelete.Enabled = False

modCon.Connected
Call RefreshUserPrev
DisableClose Me.hwnd
End Sub

Public Sub RefreshUserPrev()
lvUsertype.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserPrev Order by UserPrev", dbCon, 3, 3
    With dbRec
            Do While Not .EOF
                lvUsertype.ListItems.Add , , !UserID, 1, 1
                lvUsertype.ListItems(lvUsertype.ListItems.Count).SubItems(1) = "" & !UserPrev
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing

End Sub


Private Sub lvUsertype_Click()
On Error Resume Next
    Text1.Text = lvUsertype.SelectedItem
    txtUsertype.Text = lvUsertype.SelectedItem.SubItems(1)
    'txtUsertype.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnNew.Enabled = False
End Sub

Private Sub lvUsertype_LostFocus()
btnNew.Enabled = True
btnSave.Enabled = False
btnDelete.Enabled = False
'txtUsertype.Enabled = False
txtUsertype.Text = ""

End Sub



