VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Management System"
   ClientHeight    =   7095
   ClientLeft      =   1245
   ClientTop       =   1860
   ClientWidth     =   18255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   18255
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   1588
      ButtonWidth     =   2143
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New Book"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Explorer"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add Member"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User Manager"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Export"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Filter Member"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Book Reports"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Maintenance"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Back Up"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Logout"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Shutdown"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   720
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   20625
      TabIndex        =   1
      Top             =   840
      Width           =   20655
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "asdsadasdsad"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   17640
         TabIndex        =   3
         Top             =   60
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Library Management System By Joel Kiptoo Deus                                                            "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   14055
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   240
      Top             =   4920
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1949
            MinWidth        =   1949
            Picture         =   "MainForm.frx":2CC73
            Text            =   "Logged As:"
            TextSave        =   "Logged As:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2240
            MinWidth        =   2240
            Picture         =   "MainForm.frx":2D20D
            Text            =   "Current User:"
            TextSave        =   "Current User:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1269
            MinWidth        =   1269
            Picture         =   "MainForm.frx":2D7A7
            Text            =   "Time:"
            TextSave        =   "Time:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1835
            MinWidth        =   1835
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1480
            MinWidth        =   1480
            Picture         =   "MainForm.frx":2DD41
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "2/13/2020"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   17180
            MinWidth        =   17180
            Text            =   "Library Management System (Beta Version)"
            TextSave        =   "Library Management System (Beta Version)"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      Top             =   0
      Width           =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   31
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":2E385
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":2EFD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":2F1B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":2FE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":30A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":316A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":322F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":32F4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":33B9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":347EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":35441
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":36093
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":36CE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":37937
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":38589
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":391DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":39E2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3AA7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3B6D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3C323
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3CF75
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3DBC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3E819
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3F46B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":400BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":40D0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":41961
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":425B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":43205
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":43E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":44AA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAccount 
      Caption         =   "&Account"
      Begin VB.Menu mnuUserMan 
         Caption         =   "User Administration"
         Begin VB.Menu mnuUsrMan 
            Caption         =   "User Manager"
         End
         Begin VB.Menu mnuUserLogHis 
            Caption         =   "User Log History"
         End
      End
      Begin VB.Menu mnuMemMan 
         Caption         =   "Member Management"
         Begin VB.Menu mnuMemReg 
            Caption         =   "Member Registration"
         End
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock System"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown Computer"
      End
   End
   Begin VB.Menu mnuBooks 
      Caption         =   "&Books"
      Begin VB.Menu mnuBookRec 
         Caption         =   "Book Records"
      End
      Begin VB.Menu mnuBorrow 
         Caption         =   "Book Borrow"
      End
      Begin VB.Menu mnuRet 
         Caption         =   "Book Return"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "R&eports"
      Begin VB.Menu mnuBookMas 
         Caption         =   "Book Masterlist Report"
      End
      Begin VB.Menu mnuMemMas 
         Caption         =   "Member Masterlist Report"
      End
      Begin VB.Menu mnuUsrLogHistory 
         Caption         =   "User Log In History Report"
      End
      Begin VB.Menu mnuFiltRep 
         Caption         =   "Filter Reports"
         Begin VB.Menu mnuFULHist 
            Caption         =   "Filter User Log History Report"
         End
         Begin VB.Menu mnuFilterMem 
            Caption         =   "Filter Member Report"
         End
         Begin VB.Menu mnuFilterBorBoRep 
            Caption         =   "Filter Borrowed Book Report"
         End
         Begin VB.Menu mnuFilterRetBokRep 
            Caption         =   "Filter Returned Book Report"
         End
      End
   End
   Begin VB.Menu mnuCat 
      Caption         =   "Cate&gory"
      Begin VB.Menu mnuCourse 
         Caption         =   "New Course"
      End
      Begin VB.Menu mnuUsertyp 
         Caption         =   "New Usertype"
      End
      Begin VB.Menu mnuBooktyp 
         Caption         =   "New Booktype"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "To&ols"
      Begin VB.Menu mnuExp 
         Caption         =   "Explorer"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilities"
      Begin VB.Menu mnuBack 
         Caption         =   "Back Up Database"
      End
      Begin VB.Menu mnuRest 
         Caption         =   "Restore Database"
      End
      Begin VB.Menu mnuMainte 
         Caption         =   "Maintenance"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Ab&out"
      Begin VB.Menu mnuSys 
         Caption         =   "The System"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "H&elp"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
modCon.Connected

Label1.Caption = ""

DisableClose Me.hwnd
End Sub

Private Sub mnuBack_Click()
frmBackUp.Show 1
End Sub

Private Sub mnuBookMas_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from LBook", dbCon, 3, 3
Set rptBookMaster.DataSource = dbRec
Set dbRec = Nothing
rptBookMaster.Show 1
End Sub

Private Sub mnuBookRec_Click()
frmBookMasterlist.Show 1
End Sub

Private Sub mnuExit_Click()

End Sub

Private Sub mnuBooktyp_Click()
frmBooktype.Show 1
End Sub

Private Sub mnuBorrow_Click()
frmBorrowerCenter.Show 1
End Sub

Private Sub mnuCalc_Click()
Shell ("calc.exe")
End Sub

Private Sub mnuCourse_Click()
frmCourses.Show 1
End Sub

Private Sub mnuExp_Click()
Shell ("explorer.exe")
End Sub

Private Sub mnuFilterBorBoRep_Click()
frmFiltBorrowedBook.Show 1
End Sub

Private Sub mnuFilterMem_Click()
frmFiltMem.Show 1
End Sub

Private Sub mnuFilterRetBokRep_Click()
frmFiltReturnedBooks.Show 1
End Sub

Private Sub mnuFULHist_Click()
frmUserLog.Show 1
End Sub

Private Sub mnuLock_Click()
Load frmLock
frmLock.Show 1

End Sub

Private Sub mnuLogOut_Click()
Dim LogOut As String
Dim UserLog1 As Integer
Dim UserLog2 As String
LogOut = MsgBox("Do you really want to Log Out?", vbInformation + vbYesNo, "Logout")
If LogOut = vbYes Then
    UserLog1 = Text1.Text
    UserLog2 = Text2.Text
    
    Set dbRec2 = New ADODB.Recordset
        dbRec2.Open "Select * from UserLog where UserLogID=" & UserLog1 & "", dbCon, adOpenDynamic, adLockOptimistic
            'If dbRec2.EOF = False Then
                With dbRec2
                    !UserType = UserLog2
                    !TimeLogout = Time
                    .Update
                    
                End With
            'End If
            Unload MainForm
            frmLogin.Show
            Load frmLogin
Else
    Exit Sub
    dbRec2.Close
End If
Set dbRec2 = Nothing
End Sub

Private Sub mnuMainte_Click()
frmMaintenance.Show 1
End Sub

Private Sub mnuMemMas_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member", dbCon, 3, 3
Set rptMemberMaster.DataSource = dbRec
Set dbRec = Nothing
rptMemberMaster.Show 1
End Sub

Private Sub mnuMemReg_Click()
frmMembers.Show 1
End Sub

Private Sub mnuNotepad_Click()
Shell ("notepad.exe")
End Sub

Private Sub mnuRest_Click()
frmRestore.Show 1
End Sub

Private Sub mnuRet_Click()
frmReturnedCenter.Show 1
End Sub

Private Sub mnuShutdown_Click()
Call Shell("shutdown -s")
End Sub

Private Sub mnuSys_Click()
frmSystem.Show 1
End Sub

Private Sub mnuUserLogHis_Click()
frmUserLog.Show 1
End Sub

Private Sub mnuUsertyp_Click()
frmUserPrev.Show 1
End Sub

Private Sub mnuUsrLogHistory_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserLog", dbCon, 3, 3
Set rptUserLog.DataSource = dbRec
Set dbRec = Nothing
rptUserLog.Show 1
End Sub

Private Sub mnuUsrMan_Click()
frmUserManager.Show 1


End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Time Check:  " & Time

End Sub



Private Sub Timer2_Timer()
Dim welcome As String

welcome = "Welcome to Library Management System By Joel Kiptoo Deus"
Static i%

    i = i + 1
    Label1.Caption = Label1.Caption & Mid(welcome, i, 1)
    If Label1.Caption = welcome Then
    Timer2.Enabled = False
 End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1: frmBookMasterlist.Show 1
Case 2: Shell ("notepad.exe")
Case 3: frmMembers.Show 1
Case 4: frmUserManager.Show 1
Case 5: frmBackUp.Show 1
Case 6: frmFiltMem.Show 1
Case 7: Call mnuBookMas_Click
Case 8: frmMaintenance.Show 1
Case 9: frmBackUp.Show 1
Case 10: frmSystem.Show 1
Case 11:
Case 12: Call mnuLogOut_Click
Case 13: Call mnuShutdown_Click
End Select

End Sub


