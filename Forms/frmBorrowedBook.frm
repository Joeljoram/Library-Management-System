VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltBorrowedBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrowed Books"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvBorrow 
      Height          =   3240
      Left            =   60
      TabIndex        =   11
      Top             =   1755
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   5715
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Borrowed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Book Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Borrowers Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Book Quantity"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
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
            Picture         =   "frmBorrowedBook.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   5120
      TabIndex        =   5
      Top             =   1320
      Width           =   1930
   End
   Begin VB.TextBox txtBorrow 
      Height          =   375
      Left            =   80
      TabIndex        =   4
      Top             =   1320
      Width           =   4990
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnFilter 
      Height          =   375
      Left            =   7080
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
      Caption         =   "Filter"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   7080
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
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnPrint 
      Height          =   495
      Left            =   3960
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
      Caption         =   "Print by Date"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintInd 
      Height          =   495
      Left            =   2280
      TabIndex        =   12
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
      Left            =   3120
      TabIndex        =   13
      Top             =   5280
      Width           =   615
   End
   Begin prjLMS.jcbutton btnPrintAll 
      Height          =   495
      Left            =   5640
      TabIndex        =   14
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
      Caption         =   "Print All"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   600
      Picture         =   "frmBorrowedBook.frx":059A
      Top             =   360
      Width           =   525
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
      Left            =   240
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
      Left            =   1560
      TabIndex        =   8
      Top             =   5280
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Borrowed Books Report"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   4680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can filter the borrowed books by Name, ISBN, Author and Status."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmFiltBorrowedBook"
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
dbRec.Open "Select * from Borrow where BorrowDate Or ExpirationDate or Title or MembersName like'%" & txtBorrow.Text & "%'", dbCon, 3, 3

'Call SearchData
        If .RecordCount >= 1 Then
            lvBorrow.ListItems.Clear
                Do While Not .EOF
                    lvBorrow.ListItems.Add , , !ID, 1, 1
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtBorrow.Text = ""
            txtBorrow.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 1 Then
dbRec.Open "Select * from Borrow where BorrowDate like'%" & txtBorrow.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBorrow.ListItems.Clear
                Do While Not .EOF
                    lvBorrow.ListItems.Add , , !ID, 1, 1
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtBorrow.Text = ""
            txtBorrow.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 2 Then
dbRec.Open "Select * from Borrow where ExpirationDate like'%" & txtBorrow.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBorrow.ListItems.Clear
                Do While Not .EOF
                    lvBorrow.ListItems.Add , , !ID, 1, 1
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtBorrow.Text = ""
            txtBorrow.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 3 Then
dbRec.Open "Select * from Borrow where Title like'%" & txtBorrow.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBorrow.ListItems.Clear
                Do While Not .EOF
                    lvBorrow.ListItems.Add , , !ID, 1, 1
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtBorrow.Text = ""
            txtBorrow.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 4 Then
dbRec.Open "Select * from Borrow where MembersName like'%" & txtBorrow.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBorrow.ListItems.Clear
                Do While Not .EOF
                    lvBorrow.ListItems.Add , , !ID, 1, 1
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                    lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtBorrow.Text = ""
            txtBorrow.SetFocus
            Exit Sub
        End If
        .Close
End If
End With

Set dbRec = Nothing
End Sub

Private Sub btnPrint_Click()
frmFiltBorrowedDate.Show 1
End Sub

Private Sub btnPrintAll_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow", dbCon, 3, 3
Set rptFiltBorrow.DataSource = dbRec
Set dbRec = Nothing
rptFiltBorrow.Show 1

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
modCon.Connected

cboFilter.Clear
cboFilter.AddItem "All"
cboFilter.AddItem "Date Borrowed"
cboFilter.AddItem "Due Date"
cboFilter.AddItem "Title"
cboFilter.AddItem "Borrower's Name"
cboFilter.ListIndex = 0

Call RefreshBorrow
DisableClose Me.hwnd

End Sub


Private Sub RefreshBorrow()
Dim dbRec As ADODB.Recordset

lvBorrow.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow Order by BorrowDate", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvBorrow.ListItems.Add , , !ID, 1, 1
                lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(3) = "" & !Title
                lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                lvBorrow.ListItems(lvBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing

Label2.Caption = lvBorrow.ListItems.Count
End Sub

Private Sub btnPrintInd_Click()
If Text1.Text = "" Then
MsgBox "Please select a record in the database to print.", vbCritical, "Warning"
Else
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow where ID=" & CInt(Text1.Text) & " ", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltBorrowInd
            Set rptFiltBorrowInd.DataSource = dbRec
                .Sections("Section1").Controls("Text3").DataField = "ID"
                .Sections("Section1").Controls("Text2").DataField = "BorrowDate"
                .Sections("Section1").Controls("Text1").DataField = "ExpirationDate"
                .Sections("Section1").Controls("Text7").DataField = "Title"
                .Sections("Section1").Controls("Text6").DataField = "MembersName"
                .Sections("Section1").Controls("Text4").DataField = "BookQty"
                .Show 1
            Set dbRec = Nothing
        End With
    End If
End If

End Sub

Private Sub lvBorrow_Click()
    On Error Resume Next
    Text1.Text = lvBorrow.SelectedItem
    btnPrint.Enabled = True
End Sub


Private Sub txtBorrow_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnFilter_Click
End If
End Sub
