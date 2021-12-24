VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltReturnedBooks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returned Books"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
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
            Picture         =   "frmFiltReturnedBooks.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvReturnedBooks 
      Height          =   3225
      Left            =   30
      TabIndex        =   11
      Top             =   1750
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5689
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
         Text            =   "Return Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Borrowers Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Penalty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Returned Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtReturned 
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   8535
      _extentx        =   15055
      _extenty        =   53
   End
   Begin prjLMS.jcbutton btnFilter 
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      buttonstyle     =   2
      font            =   "frmFiltReturnedBooks.frx":059A
      backcolor       =   15199212
      caption         =   "Filter"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmFiltReturnedBooks.frx":05C2
      backcolor       =   15199212
      caption         =   "Close"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   8655
      _extentx        =   15266
      _extenty        =   53
   End
   Begin prjLMS.jcbutton btnPrintInd 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmFiltReturnedBooks.frx":05EA
      backcolor       =   15199212
      caption         =   "Print Individual"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrint 
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   5160
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmFiltReturnedBooks.frx":0612
      backcolor       =   15199212
      caption         =   "Print by Date"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintAll 
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmFiltReturnedBooks.frx":063A
      backcolor       =   15199212
      caption         =   "Print All"
      usemaskcolor    =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   720
      Picture         =   "frmFiltReturnedBooks.frx":0662
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can filter the returned books by Name, ISBN, Author and Status."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   720
      Width           =   4860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Returned Books Report"
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
      TabIndex        =   9
      Top             =   240
      Width           =   4620
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
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Shape Titl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   -2880
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmFiltReturnedBooks"
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
dbRec.Open "Select * from Return", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvReturnedBooks.ListItems.Clear
                Do While Not .EOF
                    lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtReturned.Text = ""
            txtReturned.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 1 Then
dbRec.Open "Select * from Return where ReturnDate like'%" & txtReturned.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvReturnedBooks.ListItems.Clear
                Do While Not .EOF
                    lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtReturned.Text = ""
            txtReturned.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 2 Then
dbRec.Open "Select * from Return where Title like'%" & txtReturned.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvReturnedBooks.ListItems.Clear
                Do While Not .EOF
                    lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtReturned.Text = ""
            txtReturned.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 3 Then
dbRec.Open "Select * from Return where Returnee like'%" & txtReturned.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvReturnedBooks.ListItems.Clear
                Do While Not .EOF
                    lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtReturned.Text = ""
            txtReturned.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboFilter.ListIndex = 4 Then
dbRec.Open "Select * from Return where PenaltyAmount like'%" & txtReturned.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvReturnedBooks.ListItems.Clear
                Do While Not .EOF
                    lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                    lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtReturned.Text = ""
            txtReturned.SetFocus
            Exit Sub
        End If
        .Close
End If
End With
Set dbRec = Nothing
End Sub

Private Sub btnPrint_Click()
frmReturnedBooksDate.Show 1
End Sub

Private Sub btnPrintAll_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Return", dbCon, 3, 3
Set rptFiltReturnAll.DataSource = dbRec
Set dbRec = Nothing
rptFiltReturnAll.Show 1
End Sub

Private Sub btnPrintInd_Click()
If Text1.Text = "" Then
MsgBox "Please select a record in the database to print.", vbCritical, "Warning"
Else
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Return where ID=" & CInt(Text1.Text) & " ", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltReturnInd
            Set rptFiltReturnInd.DataSource = dbRec
                .Sections("Section1").Controls("Text3").DataField = "ID"
                .Sections("Section1").Controls("Text2").DataField = "ReturnDate"
                .Sections("Section1").Controls("Text1").DataField = "Title"
                .Sections("Section1").Controls("Text7").DataField = "Returnee"
                .Sections("Section1").Controls("Text4").DataField = "PenaltyAmount"
                .Sections("Section1").Controls("Text5").DataField = "ReturnBook"
                .Show 1
            Set dbRec = Nothing
        End With
    End If
End If
End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

cboFilter.Clear
cboFilter.AddItem "All"
cboFilter.AddItem "Return Date"
cboFilter.AddItem "Title"
cboFilter.AddItem "Borrower's Name"
cboFilter.AddItem "Penalty Amount"
cboFilter.ListIndex = 0

Call RefreshReturn
DisableClose Me.hwnd

End Sub


Private Sub RefreshReturn()
Dim dbRec As ADODB.Recordset

lvReturnedBooks.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Return Order by ReturnDate", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvReturnedBooks.ListItems.Add , , !ID, 1, 1
                lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(1) = "" & !ReturnDate
                lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(2) = "" & !Title
                lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(3) = "" & !Returnee
                lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
                lvReturnedBooks.ListItems(lvReturnedBooks.ListItems.Count).SubItems(5) = "" & !ReturnBook
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing
Label2.Caption = lvReturnedBooks.ListItems.Count
End Sub


Private Sub lvReturnedBooks_Click()
Text1.Text = lvReturnedBooks.SelectedItem.Text
End Sub

Private Sub txtReturned_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnFilter_Click
End If

End Sub
