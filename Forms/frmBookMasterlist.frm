VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBookMasterlist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Masterlist Record"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   7320
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
            Picture         =   "frmBookMasterlist.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   4680
      TabIndex        =   0
      Top             =   2280
      Width           =   5175
      Begin VB.TextBox txtISBN 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtAuthor 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtBookCode 
         Height          =   285
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboBookType 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox txtPublisher 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtPage 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox txtEdition 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1920
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   7320
         ScaleHeight     =   2235
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   7320
         TabIndex        =   1
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
      End
      Begin prjLMS.jcbutton btnUpload 
         Height          =   2295
         Left            =   9360
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4048
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
         Caption         =   "Upload"
         UseMaskCOlor    =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   3000
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   38291
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edition :"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN :"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BookCode :"
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
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
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
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author :"
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
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher :"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price :"
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
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Page :"
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
         Left            =   360
         TabIndex        =   17
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Published Date :"
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
         Left            =   360
         TabIndex        =   16
         Top             =   3000
         Width           =   1170
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Type :"
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
         Left            =   360
         TabIndex        =   15
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directory :"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin MSComctlLib.ListView lvBookRec 
      Height          =   5055
      Left            =   70
      TabIndex        =   41
      Top             =   1920
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   8916
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
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Book Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ISBN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Edition"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Publisher"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Published Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Pages"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Book Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   80
      TabIndex        =   28
      Top             =   1320
      Width           =   10780
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   25
      Top             =   1360
      Width           =   2415
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   26
      Top             =   6960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnSearch 
      Height          =   495
      Left            =   13440
      TabIndex        =   27
      Top             =   1320
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
      Caption         =   "Search"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   29
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnSave 
      Height          =   495
      Left            =   7800
      TabIndex        =   30
      Top             =   7080
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
      Left            =   10680
      TabIndex        =   31
      Top             =   7080
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
      Left            =   12120
      TabIndex        =   32
      Top             =   7080
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
      Left            =   13560
      TabIndex        =   33
      Top             =   7080
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
      Left            =   9240
      TabIndex        =   34
      Top             =   7080
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
      Left            =   6360
      TabIndex        =   35
      Top             =   7080
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   40
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view and search all the book records."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   39
      Top             =   720
      Width           =   3300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Stock Masterlist"
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
      TabIndex        =   38
      Top             =   240
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1560
      TabIndex        =   37
      Top             =   7080
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record Count: "
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
      TabIndex        =   36
      Top             =   7080
      Width           =   1395
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   840
      Picture         =   "frmBookMasterlist.frx":059A
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
Attribute VB_Name = "frmBookMasterlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
txtSearch.Text = ""
txtBookCode.Text = ""
txtTitle.Text = ""
txtAuthor.Text = ""
txtISBN.Text = ""
txtEdition.Text = ""
txtPrice.Text = ""
txtPublisher.Text = ""
DTDate.Value = Date
txtPage.Text = ""
txtDirectory.Text = ""
cboBookType.Text = ""
Frame1.Visible = False
btnSave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False
btnNew.Enabled = True
txtSearch.Enabled = True
cboSearch.Enabled = True
btnSearch.Enabled = True
Call RefreshBookMaster
End Sub

Private Sub btnClose_Click()
Unload Me
Load MainForm
MainForm.Show

End Sub

Private Sub btnDelete_Click()
    MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from LBook where LBookCode='" & txtBookCode.Text & "'", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .Delete
        Call RefreshBookMaster
        .Close
    End With
    'dbCon.Close
    Set dbRec = Nothing
    txtBookCode.Text = ""
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtISBN.Text = ""
    txtEdition.Text = ""
    txtPrice.Text = ""
    txtPublisher.Text = ""
    DTDate.Value = Date
    txtPage.Text = ""
    txtDirectory.Text = ""
    cboBookType.Text = ""
    Frame1.Visible = False

    txtSearch.Enabled = True
    cboSearch.Enabled = True
    btnSearch.Enabled = True

    txtSearch.Text = ""
    txtSearch.SetFocus
    btnSave.Enabled = False
    btnUpdate.Enabled = False
    btnDelete.Enabled = False
    btnNew.Enabled = True
End Sub

Private Sub btnNew_Click()
Frame1.Visible = True
btnSave.Enabled = True
txtBookCode.SetFocus
txtSearch.Enabled = False
cboSearch.Enabled = False
btnSearch.Enabled = False
btnNew.Enabled = False



End Sub

Private Sub btnSave_Click()

If txtBookCode.Text = "" Or txtTitle.Text = "" Or txtAuthor.Text = "" Or txtISBN.Text = "" Or txtEdition.Text = "" Or txtPrice.Text = "" Or txtPublisher.Text = "" Or txtPage.Text = "" Or cboBookType.Text = "" Then
    MsgBox "Some of your fields is empty. Please complete the information", vbExclamation + vbOKOnly, "Warning"
        txtBookCode.SetFocus
Else
    MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from LBook", dbCon, adOpenKeyset, adLockPessimistic
    With dbRec
        .AddNew
        .Fields("LBookCode") = txtBookCode.Text
        .Fields("LTitle") = txtTitle.Text
        .Fields("LAuthor") = txtAuthor.Text
        .Fields("LISBN") = txtISBN.Text
        .Fields("LEdition") = txtEdition.Text
        .Fields("LPrice") = txtPrice.Text
        .Fields("LPublisher") = txtPublisher.Text
        .Fields("LPublishedDate") = DTDate.Value
        .Fields("LPages") = txtPage.Text
        .Fields("LBooktype") = cboBookType.Text
        .Update
        Call RefreshBookMaster

    End With
    dbCon.Close
    'dbRec.Close
    Set dbRec = Nothing
    'Set dbCon = Nothing
    txtBookCode.Text = ""
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtISBN.Text = ""
    txtEdition.Text = ""
    txtPrice.Text = ""
    txtPublisher.Text = ""
    'DTDate.Value = 0
    txtPage.Text = ""
    txtDirectory.Text = ""
    cboBookType.Text = ""
    Frame1.Visible = False

    txtSearch.Enabled = True
    cboSearch.Enabled = True
    btnSearch.Enabled = True

    txtSearch.Text = ""
    txtSearch.SetFocus
    btnSave.Enabled = False
    
    btnNew.Enabled = True
End If



End Sub

Private Sub btnSearch_Click()

'modCon.Connected
Set dbRec = New ADODB.Recordset

With dbRec
'On Error GoTo myerrsearch
If cboSearch.ListIndex = 0 Then
dbRec.Open "Select * from LBook", dbCon, 3, 3
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtSearch.Text = ""
            txtSearch.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboSearch.ListIndex = 1 Then
dbRec.Open "Select * from LBook where LBookCode like'%" & txtSearch.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtSearch.Text = ""
            txtSearch.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboSearch.ListIndex = 2 Then
dbRec.Open "Select * from LBook where LTitle like'%" & txtSearch.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtSearch.Text = ""
            txtSearch.SetFocus
            Exit Sub
        End If
        .Close
ElseIf cboSearch.ListIndex = 3 Then
dbRec.Open "Select * from LBook where LAuthor like'%" & txtSearch.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtSearch.Text = ""
            txtSearch.SetFocus
            Exit Sub
        End If
        .Close
Else
dbRec.Open "Select * from LBook where LPublisher like'%" & txtSearch.Text & "%'", dbCon, 3, 3
'Call SearchData
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            txtSearch.Text = ""
            txtSearch.SetFocus
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



Private Sub btnUpdate_Click()
    Frame1.Visible = True
    MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
    'modCon.Connected
    Set dbRec = New ADODB.Recordset
    dbRec.Open "Select * from LBook where ID=" & Text1.Text & "", dbCon, adOpenDynamic, adLockOptimistic
'    If dbRec.EOF = False Then
    With dbRec
        !LBookCode = txtBookCode.Text
        !LTitle = txtTitle.Text
        !LAuthor = txtAuthor.Text
        !LISBN = txtISBN.Text
        !LEdition = txtEdition.Text
        !LPrice = txtPrice.Text
        !LPublisher = txtPublisher.Text
        !LPublishedDate = DTDate.Value
        !LPages = txtPage.Text
        !LBooktype = cboBookType.Text
        dbRec.Update
        Call RefreshBookMaster
        '.Close
    End With
'    End If
    Set dbRec = Nothing
    
    'Set dbCon = Nothing
    txtBookCode.Text = ""
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtISBN.Text = ""
    txtEdition.Text = ""
    txtPrice.Text = ""
    txtPublisher.Text = ""
    DTDate.Value = Date
    txtPage.Text = ""
    txtDirectory.Text = ""
    cboBookType.Text = ""
    Frame1.Visible = False

    txtSearch.Enabled = True
    cboSearch.Enabled = True
    btnSearch.Enabled = True

    txtSearch.Text = ""
    txtSearch.SetFocus
    btnSave.Enabled = False
    btnUpdate.Enabled = False
    btnDelete.Enabled = False
    btnNew.Enabled = True

End Sub


Private Sub cboSearch_Click()
txtSearch.Text = vbNullString
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

btnSave.Enabled = False
btnUpdate.Enabled = False
btnDelete.Enabled = False
Frame1.Visible = False

modCon.Connected
Call RefreshBookMaster

'AltLVBackground lvBookRec, vbWhite, &H80000002
cboSearch.Clear
cboSearch.AddItem "All"
cboSearch.AddItem "Book Code"
cboSearch.AddItem "Title"
cboSearch.AddItem "Author"
cboSearch.AddItem "Publisher"
cboSearch.ListIndex = 0


Set dbRec3 = New ADODB.Recordset
dbRec3.Open "Select * from BookType", dbCon, 3, 3
Do While Not dbRec3.EOF
    cboBookType.AddItem dbRec3!TypesOfBook
dbRec3.MoveNext
Loop

DisableClose Me.hwnd

End Sub



Private Sub lvBookRec_DblClick()
btnUpdate.Enabled = True
btnDelete.Enabled = True
btnNew.Enabled = False
txtSearch.Enabled = False
btnSearch.Enabled = False
Frame1.Visible = True
On Error Resume Next
Text1.Text = lvBookRec.SelectedItem
txtBookCode.Text = lvBookRec.SelectedItem.SubItems(1)
txtTitle.Text = lvBookRec.SelectedItem.SubItems(2)
txtAuthor.Text = lvBookRec.SelectedItem.SubItems(3)
txtISBN.Text = lvBookRec.SelectedItem.SubItems(4)
txtEdition.Text = lvBookRec.SelectedItem.SubItems(5)
txtPrice.Text = lvBookRec.SelectedItem.SubItems(6)
txtPublisher.Text = lvBookRec.SelectedItem.SubItems(7)
DTDate.Value = lvBookRec.SelectedItem.SubItems(8)
txtPage.Text = lvBookRec.SelectedItem.SubItems(9)
cboBookType.Text = lvBookRec.SelectedItem.SubItems(10)

If Frame1.Visible = True Then
lvBookRec.SelectedItem.Selected = False
lvBookRec.SelectedItem = Nothing


End If

End Sub

Private Sub lvBookRec_LostFocus()
If Frame1.Visible = False Then
btnUpdate.Enabled = False
btnDelete.Enabled = False
btnNew.Enabled = True

'lvBookRec.ListItems.Item = False

End If
End Sub

'Private Sub AltLVBackground(lv As ListView, _
'    ByVal BackColorOne As OLE_COLOR, _
'    ByVal BackColorTwo As OLE_COLOR)
'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
'Dim lH      As Long
'Dim lSM     As Byte
'Dim picAlt  As PictureBox
'    With lvBookRec
'        If .View = lvwReport And .ListItems.Count Then
'            Set picAlt = Me.Controls.Add("VB.PictureBox", "picAlt")
'            lSM = .Parent.ScaleMode
 '           .Parent.ScaleMode = vbTwips
 '           .PictureAlignment = lvwTile
 '           lH = .ListItems(1).Height
 '           With picAlt
 '               .BackColor = BackColorOne
 '               .AutoRedraw = True
 '               .Height = lH * 2
 '               .BorderStyle = 0
 '               .Width = 10 * Screen.TwipsPerPixelX
 '               picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
 '               Set lvBookRec.Picture = .Image
 '           End With
 '           Set picAlt = Nothing
 '           Me.Controls.Remove "picAlt"
 '           lvBookRec.Parent.ScaleMode = lSM
 '       End If
 '   End With
'End Sub




Public Sub RefreshBookMaster()
lvBookRec.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from LBook Order by LTitle", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
        Do While Not .EOF
            lvBookRec.ListItems.Add , , !ID, 1, 1
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
            lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
        .MoveNext
        Loop
        .Close
    End With
    'dbCon.Close
Set dbRec = Nothing
'Set dbCon = Nothing
Label2.Caption = lvBookRec.ListItems.Count
End Sub

Public Sub SearchData()
        If .RecordCount >= 1 Then
            lvBookRec.ListItems.Clear
                Do While Not .EOF
                    lvBookRec.ListItems.Add , , !ID, 1, 1
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(1) = "" & !LBookCode
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(2) = "" & !LTitle
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(3) = "" & !LAuthor
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(4) = "" & !LISBN
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(5) = "" & !LEdition
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(6) = "" & !LPrice
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(7) = "" & !LPublisher
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(8) = "" & !LPublishedDate
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(9) = "" & !LPages
                    lvBookRec.ListItems(lvBookRec.ListItems.Count).SubItems(10) = "" & !LBooktype
                .MoveNext
                Loop
        Else
            MsgBox "No Record Found!", vbExclamation, "Warning"
            Exit Sub
        End If
        .Close
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnSearch_Click
End If
End Sub
