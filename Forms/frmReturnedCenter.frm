VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturnedCenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Returned Center"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin prjLMS.jcbutton btnCancel 
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      ToolTipText     =   "Cancel Book"
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
      Caption         =   "Cancel Return"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Returned Details"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   10455
      Begin VB.ComboBox cboAmount 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         ToolTipText     =   "Select the Penalty"
         Top             =   800
         Width           =   2895
      End
      Begin VB.ComboBox cboTitle 
         Height          =   315
         Left            =   7440
         TabIndex        =   3
         ToolTipText     =   "Select book title"
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cboMember 
         Height          =   315
         Left            =   7440
         TabIndex        =   2
         ToolTipText     =   "Select borrowers name"
         Top             =   840
         Width           =   2895
      End
      Begin MSComctlLib.ListView lvBookReturn 
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Database Record"
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date Returned"
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
            Text            =   "Penalty Amount"
            Object.Width           =   3422
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTReturned 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "Select date the book returned"
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   107085825
         CurrentDate     =   38291
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
         Left            =   5880
         TabIndex        =   9
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Returned            :"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penalty Amount          :"
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
         Width           =   1650
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
         Left            =   5880
         TabIndex        =   6
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   0
      Top             =   5880
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
            Picture         =   "frmReturnedCenter.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnReturn 
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      ToolTipText     =   "Return book"
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
      Caption         =   "Return"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      ToolTipText     =   "Close dialog"
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
   Begin prjLMS.jcbutton btnClear 
      Height          =   495
      Left            =   7800
      TabIndex        =   13
      ToolTipText     =   "Clear fields"
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can use this to track the returned books and other reference"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   720
      Width           =   4590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Returned Center"
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
      TabIndex        =   14
      Top             =   240
      Width           =   3555
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   840
      Picture         =   "frmReturnedCenter.frx":059A
      Top             =   240
      Width           =   720
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
Attribute VB_Name = "frmReturnedCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
If MsgBox("Are you sure to cancel the return of this book?", vbExclamation + vbYesNo, "Cancel Return Book") = vbYes Then
        Set dbRec = New ADODB.Recordset
        dbRec.Open "Select * from Return where ID=" & Text1.Text, dbCon, adOpenDynamic, adLockOptimistic
        With dbRec
            .Delete
            Call RefreshReturn
        End With
        MsgBox "Returned Book Successfully Cancelled.", vbInformation, "Success!"
        Set dbRec = Nothing
        btnCancel.Enabled = False
Else
    Exit Sub
End If
End Sub

Private Sub btnClear_Click()
btnReturn.Enabled = True
btnCancel.Enabled = False
cboAmount.Text = ""
cboTitle.Text = ""
cboMember.Text = ""

End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub btnReturn_Click()
If DTReturned.Value = 0 Or cboAmount.Text = "" Or cboMember.Text = "" Then
    MsgBox "Some of your fields is empty. Please complete the information", vbExclamation + vbOKOnly, "Warning"
Else
If MsgBox("Are you sure you want to return this book?", vbInformation + vbYesNo, "Warning") = vbYes Then
    
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from Return", dbCon, 3, 3
    With dbRec
        .AddNew
        .Fields("ReturnDate") = DTReturned.Value
        .Fields("Title") = cboTitle.Text
        .Fields("Returnee") = cboMember.Text
        .Fields("PenaltyAmount") = cboAmount.Text
        .Update
        .Close
        Call RefreshReturn
    End With
    Set dbRec = Nothing
    btnReturn.Enabled = False
    MsgBox "Book Successfully Returned!", vbInformation + vbOKOnly, "Success!"
Else
    Exit Sub
End If
End If

End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

cboAmount.Clear
cboAmount.AddItem "No Penalty"
cboAmount.AddItem "1st Offense - 5.00"
cboAmount.AddItem "2nd Offense - 10.00"
cboAmount.AddItem "3rd Offense - 15.00"
cboAmount.AddItem "4th Offense - 20.00"
cboAmount.ListIndex = 0

DTReturned.Value = Date
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

Call RefreshReturn
DisableClose Me.hwnd
btnCancel.Enabled = False
End Sub

Private Sub RefreshReturn()
lvBookReturn.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Return Order by Returnee", dbCon, 3, 3
    With dbRec
            Do While Not .EOF
                lvBookReturn.ListItems.Add , , !ID, 1, 1
                lvBookReturn.ListItems(lvBookReturn.ListItems.Count).SubItems(1) = "" & !ReturnDate
                lvBookReturn.ListItems(lvBookReturn.ListItems.Count).SubItems(2) = "" & !Title
                lvBookReturn.ListItems(lvBookReturn.ListItems.Count).SubItems(3) = "" & !Returnee
                lvBookReturn.ListItems(lvBookReturn.ListItems.Count).SubItems(4) = "" & !PenaltyAmount
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing

End Sub


Private Sub lvBookReturn_Click()
On Error Resume Next
Text1.Text = lvBookReturn.SelectedItem
DTReturned.Value = lvBookReturn.SelectedItem.SubItems(1)
cboTitle.Text = lvBookReturn.SelectedItem.SubItems(2)
cboMember.Text = lvBookReturn.SelectedItem.SubItems(3)
cboAmount.Text = lvBookReturn.SelectedItem.SubItems(4)
btnCancel.Enabled = True
btnReturn.Enabled = False
End Sub

Private Sub lvBookReturn_LostFocus()
btnCancel.Enabled = False
btnReturn.Enabled = True
End Sub
