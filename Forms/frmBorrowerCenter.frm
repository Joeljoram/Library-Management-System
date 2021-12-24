VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBorrowerCenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrower Center"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   6360
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   6120
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
            Picture         =   "frmBorrowerCenter.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Borrowing Details"
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
      Top             =   1440
      Width           =   10455
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   9840
         MaxLength       =   1
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cboMember 
         Height          =   315
         Left            =   6600
         TabIndex        =   12
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboTitle 
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         Top             =   480
         Width           =   2415
      End
      Begin MSComctlLib.ListView lvBookBorrow 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
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
            Text            =   "Date Borrowed"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Expiration Date"
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
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTBorrowed 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   107610113
         CurrentDate     =   38291
      End
      Begin MSComCtl2.DTPicker DTExpired 
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Format          =   107610113
         CurrentDate     =   38291
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty     :"
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
         Left            =   9120
         TabIndex        =   19
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title                         :"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date                 :"
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Borrowed        :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrower's Name     :"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   1500
      End
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnBorrow 
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   5760
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
      Caption         =   "Borrow"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   5760
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
   Begin prjLMS.jcbutton btnCancel 
      Height          =   495
      Left            =   6360
      TabIndex        =   15
      Top             =   5760
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
      Caption         =   "Cancel Borrow"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
   End
   Begin prjLMS.jcbutton btnClear 
      Height          =   495
      Left            =   7800
      TabIndex        =   18
      Top             =   5760
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
   Begin VB.Image Image4 
      Height          =   720
      Left            =   840
      Picture         =   "frmBorrowerCenter.frx":059A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Borrowing Center"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   3690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can use this to borrow books and other reference."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmBorrowerCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBorrow_Click()
    


If DTBorrowed.Value = 0 Or DTExpired.Value = 0 Or cboTitle.Text = "" Or cboMember.Text = "" Then
    MsgBox "Some of your fields is empty. Please complete the information", vbExclamation + vbOKOnly, "Warning"
Else
If MsgBox("Are you sure you want to borrow this book?", vbInformation + vbYesNo, "Warning") = vbYes Then
    
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Borrow", dbCon, 3, 3
    With dbRec
        .AddNew
        .Fields("BorrowDate") = DTBorrowed.Value
        .Fields("ExpirationDate") = DTExpired.Value
        .Fields("Title") = cboTitle.Text
        .Fields("MembersName") = cboMember.Text
        .Fields("BookQty") = txtQty.Text
        .Update
        .Close
        Call RefreshBorrow
    End With
    Set dbRec = Nothing
    btnBorrow.Enabled = False
    MsgBox "Book Successfully Borrowed!", vbInformation + vbOKOnly, "Success!"
Else
    Exit Sub
End If
End If
End Sub

Private Sub btnCancel_Click()
If MsgBox("Are you sure you want to cancel this book?", vbYesNo + vbQuestion, "Delete") = vbYes Then
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Borrow where ID=" & Text1.Text, dbCon, adOpenDynamic, adLockOptimistic
    With dbRec
        .Delete
        Call RefreshBorrow
    End With
    MsgBox "Book Successfully Cancelled.", vbInformation, "Success!"
    Set dbRec = Nothing
Else
Text1.Text = ""
Exit Sub
End If

End Sub

Private Sub btnClear_Click()
btnBorrow.Enabled = True
btnCancel.Enabled = False
cboTitle.Text = ""
cboMember.Text = ""
DTBorrowed.Value = Date
DTExpired.Value = Date
End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DTBorrowed.Value = Date
DTExpired.Value = Date
btnCancel.Enabled = False
modCon.Connected
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from LBook Order by LTitle", dbCon, 3, 3
Do While Not dbRec.EOF
    cboTitle.AddItem dbRec!LTitle
dbRec.MoveNext
Loop

Set dbRec2 = New ADODB.Recordset
dbRec2.Open "Select * from Member Order by Name", dbCon, 3, 3
Do While Not dbRec2.EOF
    cboMember.AddItem dbRec2!Name
dbRec2.MoveNext
Loop

Call RefreshBorrow
DisableClose Me.hwnd
End Sub

Private Sub RefreshBorrow()
lvBookBorrow.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow Order by ExpirationDate", dbCon, 3, 3
    With dbRec
            Do While Not .EOF
                lvBookBorrow.ListItems.Add , , !ID, 1, 1
                lvBookBorrow.ListItems(lvBookBorrow.ListItems.Count).SubItems(1) = "" & !BorrowDate
                lvBookBorrow.ListItems(lvBookBorrow.ListItems.Count).SubItems(2) = "" & !ExpirationDate
                lvBookBorrow.ListItems(lvBookBorrow.ListItems.Count).SubItems(3) = "" & !Title
                lvBookBorrow.ListItems(lvBookBorrow.ListItems.Count).SubItems(4) = "" & !MembersName
                lvBookBorrow.ListItems(lvBookBorrow.ListItems.Count).SubItems(5) = "" & !BookQty
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing


End Sub

Private Sub lvBookBorrow_Click()
On Error Resume Next
Text1.Text = lvBookBorrow.SelectedItem
DTBorrowed.Value = lvBookBorrow.SelectedItem.SubItems(1)
DTExpired.Value = lvBookBorrow.SelectedItem.SubItems(2)
cboTitle.Text = lvBookBorrow.SelectedItem.SubItems(3)
cboMember.Text = lvBookBorrow.SelectedItem.SubItems(4)
txtQty.Text = lvBookBorrow.SelectedItem.SubItems(5)

btnBorrow.Enabled = False
btnCancel.Enabled = True
End Sub

Private Sub lvBookBorrow_LostFocus()
btnBorrow.Enabled = True
End Sub
