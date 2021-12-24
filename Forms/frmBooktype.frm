VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBooktype 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Types of Books"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin prjLMS.jcbutton btnClear 
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   3840
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
            Picture         =   "frmBooktype.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvType 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Book Types"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtBook 
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   3405
      Width           =   5655
   End
   Begin prjLMS.jcbutton btnNew 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3840
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
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
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
      Caption         =   "Close"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnSave 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3840
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
   Begin prjLMS.jcbutton btnDelete 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3840
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
      TabIndex        =   6
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7858
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add && Delete Type of Books"
            Object.Tag             =   ""
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
End
Attribute VB_Name = "frmBooktype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
txtBook.Text = ""
btnDelete.Enabled = False
btnSave.Enabled = False
btnNew.Enabled = True
txtBook.SetFocus
End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub btnDelete_Click()

If MsgBox("Are you sure you want to delete this record: " & txtBook.Text & " ?", vbYesNo + vbQuestion, "Delete") = vbYes Then
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from BookType where ID=" & Text1.Text, dbCon, adOpenDynamic, adLockOptimistic
    With dbRec
        .Delete
        Call RefreshBook
    End With
    MsgBox "Record Successfully Deleted.", vbInformation, "Delete Success"
    btnNew.Enabled = True
    btnDelete.Enabled = False
    Set dbRec = Nothing
    Text1.Text = ""
    txtBook.Text = ""
Else
Text1.Text = ""
txtBook.Text = ""
Exit Sub
End If
End Sub

Private Sub btnNew_Click()
btnNew.Enabled = False
btnSave.Enabled = True
txtBook.Enabled = True
txtBook.SetFocus
End Sub

Private Sub btnSave_Click()
If IsNumeric(txtBook.Text) = True Then
    MsgBox "Book Type must be a word value.", vbExclamation + vbOKOnly, "Warning"
    txtBook.Text = ""
    txtBook.SetFocus
ElseIf txtBook.Text = "" Then
    MsgBox "The box is empty.", vbExclamation + vbOKOnly, "Warning"
    txtBook.SetFocus
Else
    MsgBox "Record Successfully Saved!", vbInformation + vbOKOnly, "Success Saved!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from BookType", dbCon, 3, 3
    With dbRec
        .AddNew
        .Fields("TypesOfBook") = txtBook.Text
        .Update
        .Close
        Call RefreshBook
    End With
    Set dbRec = Nothing
    btnSave.Enabled = False
    btnNew.Enabled = True
    txtBook.Text = ""
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

btnSave.Enabled = False
btnDelete.Enabled = False
txtBook.Enabled = False

modCon.Connected
Call RefreshBook
DisableClose Me.hwnd

End Sub


Public Sub RefreshBook()
lvType.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from BookType Order by TypesOfBook", dbCon, 3, 3
    With dbRec
            Do While Not .EOF
                lvType.ListItems.Add , , !ID, 1, 1
                lvType.ListItems(lvType.ListItems.Count).SubItems(1) = "" & !TypesOfBook
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing


End Sub


Private Sub lvType_Click()
On Error Resume Next
    Text1.Text = lvType.SelectedItem
    txtBook.Text = lvType.SelectedItem.SubItems(1)
    txtBook.Enabled = True
    btnDelete.Enabled = True
    btnSave.Enabled = False
    btnNew.Enabled = False
End Sub


