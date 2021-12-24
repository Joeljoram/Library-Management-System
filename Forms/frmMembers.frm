VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMembers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Membership Registration"
   ClientHeight    =   6240
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   4080
      TabIndex        =   32
      Top             =   6480
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   5640
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
            Picture         =   "frmMembers.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Member Fill Up Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10215
      Begin VB.Frame Frame2 
         Caption         =   "Member Database Record"
         Height          =   2775
         Left            =   5640
         TabIndex        =   30
         Top             =   1200
         Width           =   4455
         Begin MSComctlLib.ListView lvMember 
            Height          =   2415
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4260
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
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Stud / Emp ID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Address"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Date Of Birth"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Phone No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Email"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Date Created"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Gender"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Course"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Section"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         ToolTipText     =   "Type your email address. eg. youremail@yahoo.com"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         ToolTipText     =   "Type your Phone Number. eg. 091987654321"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtSection 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         ToolTipText     =   "Type your Year and Section. eg. 2nd - A1"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.ComboBox cboCourse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         ToolTipText     =   "Selct your course"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         ToolTipText     =   "Type your complete address"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         ToolTipText     =   "Type your Complete Name. eg. Juan Dela Cruz"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtStudID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         ToolTipText     =   "Type your Student or Employee ID. eg. 00345324"
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox cboGender 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   15
         ToolTipText     =   "Select your gender"
         Top             =   360
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtDc 
         Height          =   255
         Left            =   7560
         TabIndex        =   14
         ToolTipText     =   "Select your date of birth"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   107085825
         CurrentDate     =   41844
      End
      Begin MSComCtl2.DTPicker dtDob 
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         ToolTipText     =   "Select the date of registration"
         Top             =   3720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         Format          =   107085825
         CurrentDate     =   41844
      End
      Begin VB.Label Label12 
         Caption         =   "Date Of Birth                  :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Email                               :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Phone No.                      :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Date Created                 :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Gender                           :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Student/Employee ID     :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Address                          :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Year and Section                           :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Course                           :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Complete Name              :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnSave 
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   5640
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
      Left            =   4680
      TabIndex        =   25
      Top             =   5640
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
      Left            =   7560
      TabIndex        =   26
      Top             =   5640
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
      Left            =   9000
      TabIndex        =   27
      Top             =   5640
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
      Left            =   6120
      TabIndex        =   28
      Top             =   5640
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
   Begin prjLMS.jcbutton btnNew 
      Height          =   495
      Left            =   1800
      TabIndex        =   29
      Top             =   5640
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
   Begin VB.Image Image1 
      Height          =   735
      Left            =   600
      Picture         =   "frmMembers.frx":059A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Membership Registration"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage the administration of membership registration."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   3780
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClear_Click()
    txtStudID.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtEmail.Text = ""
    cboGender.Text = ""
    cboCourse.Text = ""
    txtSection.Text = ""
    dtDob.Value = Date
    dtDc.Value = Date
    btnSave.Enabled = False
    btnUpdate.Enabled = False
    btnDelete.Enabled = False
    btnNew.Enabled = True
    txtStudID.Enabled = False
txtName.Enabled = False
txtAddress.Enabled = False
dtDob.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False
dtDc.Enabled = False
cboGender.Enabled = False
cboCourse.Enabled = False
txtSection.Enabled = False
End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub btnDelete_Click()
    MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Member where ID=" & Text1.Text & "", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .Delete
        Call RefreshMember
        .Close
    End With
    Set dbRec = Nothing
    btnNew.Enabled = True
    btnDelete.Enabled = False
    btnUpdate.Enabled = False
    txtStudID.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtEmail.Text = ""
    cboGender.Text = ""
    cboCourse.Text = ""
    txtSection.Text = ""
    dtDob.Value = Date
    dtDc.Value = Date
txtName.Enabled = False
txtAddress.Enabled = False
dtDob.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False
dtDc.Enabled = False
cboGender.Enabled = False
cboCourse.Enabled = False
txtSection.Enabled = False
End Sub

Private Sub btnNew_Click()
btnNew.Enabled = False
btnSave.Enabled = True
txtStudID.Enabled = True
txtName.Enabled = True
txtAddress.Enabled = True
dtDob.Enabled = True
txtPhone.Enabled = True
txtEmail.Enabled = True
dtDc.Enabled = True
cboGender.Enabled = True
cboCourse.Enabled = True
txtSection.Enabled = True
txtStudID.SetFocus
End Sub

Private Sub btnSave_Click()
If txtStudID.Text = "" Or txtName.Text = "" Or txtAddress.Text = "" Or cboCourse.Text = "" Or txtSection.Text = "" Or txtPhone.Text = "" Or txtEmail.Text = "" Or cboGender.Text = "" Or dtDc.Value = 0 Or dtDob.Value = 0 Then
    MsgBox "Some of your fields is empty. Please complete the information", vbExclamation + vbOKOnly, "Warning"
    txtStudID.SetFocus
Else
    MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Member", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .AddNew
        .Fields("StudentID") = txtStudID.Text
        .Fields("Name") = txtName.Text
        .Fields("Address") = txtAddress.Text
        .Fields("DOB") = dtDob.Value
        .Fields("PhoneNo") = txtPhone.Text
        .Fields("Email") = txtEmail.Text
        .Fields("DateCreated") = dtDc.Value
        .Fields("Gender") = cboGender.Text
        .Fields("Course") = cboCourse.Text
        .Fields("Section") = txtSection.Text
        .Update
        Call RefreshMember
    End With
    dbRec.Close
    Set dbRec = Nothing
    txtStudID.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtEmail.Text = ""
    dtDob.Value = Date
    dtDc.Value = Date
    cboGender.Text = ""
    cboCourse.Text = ""
    txtSection.Text = ""
    btnSave.Enabled = False
    btnNew.Enabled = True
End If

End Sub

Private Sub btnUpdate_Click()
    MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Member where ID=" & Text1.Text & "", dbCon, adOpenDynamic, adLockOptimistic
    With dbRec
        !StudentID = txtStudID.Text
        !Name = txtName.Text
        !Address = txtAddress.Text
        !DOB = dtDob.Value
        !PhoneNo = txtPhone.Text
        !Email = txtEmail.Text
        !DateCreated = dtDc.Value
        !Gender = cboGender.Text
        !Course = cboCourse.Text
        !Section = txtSection.Text
        dbRec.Update
        Call RefreshMember
    End With
    Set dbRec = Nothing
    btnNew.Enabled = True
    btnDelete.Enabled = False
    btnUpdate.Enabled = False
    txtStudID.Text = ""
    txtName.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtEmail.Text = ""
    cboGender.Text = ""
    cboCourse.Text = ""
    txtSection.Text = ""
    dtDob.Value = Date
    dtDc.Value = Date
txtStudID.Enabled = False
txtName.Enabled = False
txtAddress.Enabled = False
dtDob.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False
dtDc.Enabled = False
cboGender.Enabled = False
cboCourse.Enabled = False
txtSection.Enabled = False
End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2


Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Courses", dbCon, 3, 3
Do While Not dbRec.EOF
    cboCourse.AddItem dbRec!CourseName
dbRec.MoveNext
Loop


btnSave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False

cboGender.AddItem "Male"
cboGender.AddItem "Female"

Call RefreshMember
DisableClose Me.hwnd

txtStudID.Enabled = False
txtName.Enabled = False
txtAddress.Enabled = False
dtDob.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False
dtDc.Enabled = False
cboGender.Enabled = False
cboCourse.Enabled = False
txtSection.Enabled = False
End Sub

Private Sub RefreshMember()
Dim dbRec As ADODB.Recordset

lvMember.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member Order by Name", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvMember.ListItems.Add , , !ID, 1, 1
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(1) = "" & !StudentID
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(2) = "" & !Name
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(3) = "" & !Address
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(4) = "" & !DOB
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(5) = "" & !PhoneNo
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(6) = "" & !Email
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(7) = "" & !DateCreated
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(8) = "" & !Gender
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(9) = "" & !Course
                lvMember.ListItems(lvMember.ListItems.Count).SubItems(10) = "" & !Section
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing
                
End Sub



Private Sub lvMember_Click()
btnDelete.Enabled = True
btnUpdate.Enabled = True
txtStudID.Enabled = True
txtName.Enabled = True
txtAddress.Enabled = True
dtDob.Enabled = True
txtPhone.Enabled = True
txtEmail.Enabled = True
dtDc.Enabled = True
cboGender.Enabled = True
cboCourse.Enabled = True
txtSection.Enabled = True
btnNew.Enabled = False
On Error Resume Next
Text1.Text = lvMember.SelectedItem
txtStudID.Text = lvMember.SelectedItem.SubItems(1)
txtName.Text = lvMember.SelectedItem.SubItems(2)
txtAddress.Text = lvMember.SelectedItem.SubItems(3)
dtDob.Value = lvMember.SelectedItem.SubItems(4)
txtPhone.Text = lvMember.SelectedItem.SubItems(5)
txtEmail.Text = lvMember.SelectedItem.SubItems(6)
dtDc.Value = lvMember.SelectedItem.SubItems(7)
cboGender.Text = lvMember.SelectedItem.SubItems(8)
cboCourse.Text = lvMember.SelectedItem.SubItems(9)
txtSection.Text = lvMember.SelectedItem.SubItems(10)
End Sub
