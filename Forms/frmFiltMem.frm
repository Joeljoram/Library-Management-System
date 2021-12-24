VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltMem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Member"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvMemRep 
      Height          =   3255
      Left            =   30
      TabIndex        =   11
      Top             =   1740
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "StudentID"
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
         Text            =   "Phone No."
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   5280
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
            Picture         =   "frmFiltMem.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtMemFilter 
      Height          =   375
      Left            =   40
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -120
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnFilter 
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Caption         =   "Search"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   5160
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
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnPrint 
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Print Individual"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   5280
      Width           =   375
   End
   Begin prjLMS.jcbutton btnPrintCourse 
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Print by Course"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintSection 
      Height          =   495
      Left            =   5040
      TabIndex        =   14
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Print by Section"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintGender 
      Height          =   495
      Left            =   6600
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Print by Gender"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Count: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1440
      TabIndex        =   8
      Top             =   5280
      Width           =   120
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   480
      Picture         =   "frmFiltMem.frx":059A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Member Report"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   3360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can filter member information by Name, Course, Date Created && Gender"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   5340
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
Attribute VB_Name = "frmFiltMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub btnFilter_Click()
Set dbRec = New ADODB.Recordset

With dbRec
'On Error GoTo myerrsearch
If cboFilter.ListIndex = 0 Then
dbRec.Open "Select * from Member", dbCon, 3, 3
        If .RecordCount >= 1 Then
            lvMemRep.ListItems.Clear
                Do While Not .EOF
                    lvMemRep.ListItems.Add , , !ID, 1, 1
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtMemFilter.Text = ""
            txtMemFilter.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 1 Then
dbRec.Open "Select * from Member where Name like'%" & txtMemFilter.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvMemRep.ListItems.Clear
                Do While Not .EOF
                    lvMemRep.ListItems.Add , , !ID, 1, 1
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtMemFilter.Text = ""
            txtMemFilter.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 2 Then
dbRec.Open "Select * from Member where Course like'%" & txtMemFilter.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvMemRep.ListItems.Clear
                Do While Not .EOF
                    lvMemRep.ListItems.Add , , !ID, 1, 1
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtMemFilter.Text = ""
            txtMemFilter.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 3 Then
dbRec.Open "Select * from Member where [Section] like'" & txtMemFilter.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvMemRep.ListItems.Clear
                Do While Not .EOF
                    lvMemRep.ListItems.Add , , !ID, 1, 1
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtMemFilter.Text = ""
            txtMemFilter.SetFocus
            Exit Sub
        End If
        .Close
Else
dbRec.Open "Select * from Member where Gender like'%" & txtMemFilter.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvMemRep.ListItems.Clear
                Do While Not .EOF
                    lvMemRep.ListItems.Add , , !ID, 1, 1
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                    lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtMemFilter.Text = ""
            txtMemFilter.SetFocus
            Exit Sub
        End If
        .Close
End If

End With
'Exit Sub:
'myerrsearch:
'MsgBox Err.Description, vbCritical, "Error"
Set dbRec = Nothing

End Sub

Private Sub btnPrint_Click()
If Text1.Text = "" Then
MsgBox "Please select a record in the database to print.", vbCritical, "Warning"
Else
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member where ID=" & CInt(Text1.Text) & " ", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltMemPrintInd
            Set rptFiltMemPrintInd.DataSource = dbRec
                .Sections("Section1").Controls("Text3").DataField = "ID"
                .Sections("Section1").Controls("Text2").DataField = "StudentID"
                .Sections("Section1").Controls("Text1").DataField = "Name"
                .Sections("Section1").Controls("Text7").DataField = "Address"
                .Sections("Section1").Controls("Text6").DataField = "DOB"
                .Sections("Section1").Controls("Text10").DataField = "PhoneNo"
                .Sections("Section1").Controls("Text5").DataField = "Email"
                .Sections("Section1").Controls("Text8").DataField = "DateCreated"
                .Sections("Section1").Controls("Text4").DataField = "Gender"
                .Sections("Section1").Controls("Text9").DataField = "Course"
                .Sections("Section1").Controls("Text11").DataField = "Section"
                .Show 1
            Set dbRec = Nothing
        End With
    End If
End If
End Sub


Private Sub btnPrintCourse_Click()
frmFiltMemCourse.Show 1
End Sub

Private Sub btnPrintSection_Click()
frmFiltMemSec.Show 1
End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
cboFilter.Clear
cboFilter.AddItem "All"
cboFilter.AddItem "Name"
cboFilter.AddItem "Course"
cboFilter.AddItem "Section"
cboFilter.AddItem "Gender"
cboFilter.ListIndex = 0
'btnPrint.Enabled = False

Call RefreshMember
DisableClose Me.hwnd


End Sub


Private Sub RefreshMember()
Dim dbRec As ADODB.Recordset

lvMemRep.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member Order by Name", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvMemRep.ListItems.Add , , !ID, 1, 1
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(1) = "" & !StudentID
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(2) = "" & !Name
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(3) = "" & !Address
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(4) = "" & !DOB
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(5) = "" & !PhoneNo
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(6) = "" & !Email
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(7) = "" & !DateCreated
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(8) = "" & !Gender
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(9) = "" & !Course
                lvMemRep.ListItems(lvMemRep.ListItems.Count).SubItems(10) = "" & !Section
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing
Label2.Caption = lvMemRep.ListItems.Count
                
End Sub

Private Sub lvMemRep_Click()
btnPrint.Enabled = True
On Error Resume Next
Text1.Text = lvMemRep.SelectedItem
End Sub

Private Sub lvMemRep_LostFocus()
'btnPrint.Enabled = False
txtMemFilter.SetFocus

End Sub

Private Sub txtMemFilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnFilter_Click
End If
End Sub
