VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup and Export Database"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   240
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export to Excel Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   5265
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   70
         Left            =   5280
         Top             =   1800
      End
      Begin prjLMS.jcbutton btnExp 
         Height          =   495
         Left            =   90
         TabIndex        =   3
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Export [BOOK] Data"
         UseMaskCOlor    =   -1  'True
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   2760
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prjLMS.jcbutton btnExpMem 
         Height          =   495
         Left            =   90
         TabIndex        =   5
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Export [MEMBER] Data"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnExpReturn 
         Height          =   495
         Left            =   90
         TabIndex        =   6
         Top             =   2160
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Export [RETURN] Data"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnExpBorrow 
         Height          =   495
         Left            =   90
         TabIndex        =   7
         Top             =   1560
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Export [BORROW] Data"
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   3015
         Width           =   45
      End
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   90
      TabIndex        =   9
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
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
   Begin prjLMS.jcbutton btnBackUp 
      Height          =   495
      Left            =   90
      TabIndex        =   10
      Top             =   1920
      Width           =   5055
      _ExtentX        =   8916
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
      Caption         =   "Backup"
      UseMaskCOlor    =   -1  'True
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   195
      Left            =   160
      TabIndex        =   11
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmBackUp.frx":0000
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Database"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2970
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can back up your database record on this section."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3885
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim DataArray(1 To 1000, 1 To 10) As Variant
Dim R As Integer
Dim NumberOfRows As Integer

Private Sub btnBackup_Click()
modBackup.MDbackupdatabases
'Timer2.Enabled = True
'FileCopy App.Path & "\LMS\Database\data.mdb", App.Path & "\Database\data-" & Format(Date, "dd-mm-yyyy") & ".mdb"
End Sub

Private Sub btnCancel_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub


Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnExp_Click()

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from LBook", dbCon, adOpenDynamic, adLockOptimistic
NumberOfRows = dbRec.RecordCount
dbRec.MoveFirst
For R = 1 To NumberOfRows
DataArray(R, 1) = dbRec.Fields("LBookCode")
DataArray(R, 2) = dbRec.Fields("LTitle")
DataArray(R, 3) = dbRec.Fields("LAuthor")
DataArray(R, 4) = dbRec.Fields("LISBN")
DataArray(R, 5) = dbRec.Fields("LEdition")
DataArray(R, 6) = dbRec.Fields("LPrice")
DataArray(R, 7) = dbRec.Fields("LPublisher")
DataArray(R, 8) = dbRec.Fields("LPublishedDate")
DataArray(R, 9) = dbRec.Fields("LPages")
DataArray(R, 10) = dbRec.Fields("LBooktype")
dbRec.MoveNext
Next

Set oSheet = oBook.Worksheets(1)
oSheet.Range("A1:J1").Font.Bold = True

oSheet.Range("A1 :J1").Value = Array("Book Code", "Title", "Author", "ISBN", "Edition", "Price", "Publisher", "Published Date", "Pages", "Booktype")
oSheet.Range("A2").Resize(NumberOfRows, 10).Value = DataArray

oBook.SaveAs App.Path & "\Export\Books-" & Format(Date, "dd-mm-yyyy") & ".xls"
oExcel.Quit
dbRec.MoveFirst
Set dbRec = Nothing
Timer1.Enabled = True
End Sub

Private Sub btnExpBorrow_Click()
Dim DataArrayBor(1 To 1000, 1 To 5) As Variant
Dim oExcelll As Object
Dim oBookkk As Object
Dim oSheettt As Object
Dim rrr As Integer
Dim NumberOfRowsss As Integer

Set oExcelll = CreateObject("Excel.Application")
Set oBookkk = oExcelll.Workbooks.Add

Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow", dbCon, adOpenDynamic, adLockOptimistic
NumberOfRowsss = dbRec.RecordCount
dbRec.MoveFirst
For rrr = 1 To NumberOfRowsss
DataArrayBor(rrr, 1) = dbRec.Fields("BorrowDate")
DataArrayBor(rrr, 2) = dbRec.Fields("ExpirationDate")
DataArrayBor(rrr, 3) = dbRec.Fields("Title")
DataArrayBor(rrr, 4) = dbRec.Fields("MembersName")
DataArrayBor(rrr, 5) = dbRec.Fields("BookQty")
dbRec.MoveNext
Next

Set oSheettt = oBookkk.Worksheets(1)
oSheettt.Range("A1:E1").Font.Bold = True

oSheettt.Range("A1 :E1").Value = Array("Date Borrowed     ", "Due Date     ", "Book Title     ", "Borrower's Name     ", "Quantity")
oSheettt.Range("A2").Resize(NumberOfRowsss, 5).Value = DataArrayBor

oBookkk.SaveAs App.Path & "\Export\Borrow-" & Format(Date, "dd-mm-yyyy") & ".xls"
oExcelll.Quit
dbRec.MoveFirst
Set dbRec = Nothing
Timer1.Enabled = True

End Sub

Private Sub btnExpMem_Click()
Dim DataArrayMem(1 To 1000, 1 To 10) As Variant
Dim oExcell As Object
Dim oBookk As Object
Dim oSheett As Object
Dim rr As Integer
Dim NumberOfRowss As Integer

Set oExcell = CreateObject("Excel.Application")
Set oBookk = oExcell.Workbooks.Add

Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member", dbCon, adOpenDynamic, adLockOptimistic
NumberOfRowss = dbRec.RecordCount
dbRec.MoveFirst
For rr = 1 To NumberOfRowss
DataArrayMem(rr, 1) = dbRec.Fields("StudentID")
DataArrayMem(rr, 2) = dbRec.Fields("Name")
DataArrayMem(rr, 3) = dbRec.Fields("Address")
DataArrayMem(rr, 4) = dbRec.Fields("DOB")
DataArrayMem(rr, 5) = dbRec.Fields("PhoneNo")
DataArrayMem(rr, 6) = dbRec.Fields("Email")
DataArrayMem(rr, 7) = dbRec.Fields("DateCreated")
DataArrayMem(rr, 8) = dbRec.Fields("Gender")
DataArrayMem(rr, 9) = dbRec.Fields("Course")
DataArrayMem(rr, 10) = dbRec.Fields("Section")
dbRec.MoveNext
Next

Set oSheett = oBookk.Worksheets(1)
oSheett.Range("A1:J1").Font.Bold = True

oSheett.Range("A1 :J1").Value = Array("Student ID", "Name", "Address", "DOB", "PhoneNo", "Email", "Date Created", "Gender", "Course", "Section")
oSheett.Range("A2").Resize(NumberOfRowss, 10).Value = DataArrayMem

oBookk.SaveAs App.Path & "\Export\Members-" & Format(Date, "dd-mm-yyyy") & ".xls"
oExcell.Quit
dbRec.MoveFirst
Set dbRec = Nothing
Timer1.Enabled = True

End Sub

Private Sub btnExpReturn_Click()
Dim DataArrayRet(1 To 1000, 1 To 5) As Variant
Dim oExcellll As Object
Dim oBookkkk As Object
Dim oSheetttt As Object
Dim rrrr As Integer
Dim NumberOfRowssss As Integer

Set oExcellll = CreateObject("Excel.Application")
Set oBookkkk = oExcellll.Workbooks.Add

Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Return", dbCon, adOpenDynamic, adLockOptimistic
NumberOfRowssss = dbRec.RecordCount
dbRec.MoveFirst
For rrrr = 1 To NumberOfRowssss
DataArrayRet(rrrr, 1) = dbRec.Fields("ReturnDate")
DataArrayRet(rrrr, 2) = dbRec.Fields("Title")
DataArrayRet(rrrr, 3) = dbRec.Fields("Returnee")
DataArrayRet(rrrr, 4) = dbRec.Fields("PenaltyAmount")
DataArrayRet(rrrr, 5) = dbRec.Fields("ReturnBook")
dbRec.MoveNext
Next

Set oSheetttt = oBookkkk.Worksheets(1)
oSheetttt.Range("A1:E1").Font.Bold = True

oSheetttt.Range("A1 :E1").Value = Array("Date Returned", "Book Title", "Borrower's Name", "Penalty", "Returned Quantity")
oSheetttt.Range("A2").Resize(NumberOfRowssss, 5).Value = DataArrayRet

oBookkkk.SaveAs App.Path & "\Export\Return-" & Format(Date, "dd-mm-yyyy") & ".xls"
oExcellll.Quit
dbRec.MoveFirst
Set dbRec = Nothing
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd

Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
Dim a As Long
Timer1.Enabled = True
ProgressBar1.Max = 101
ProgressBar1.Value = ProgressBar1.Value + 1
btnExpMem.Enabled = False
btnExp.Enabled = False
btnExpBorrow.Enabled = False
btnExpReturn.Enabled = False
If ProgressBar1.Value = 20 Then
Label7.Caption = "Preparing to Export Files..."
ElseIf ProgressBar1.Value = 40 Then
Label7.Caption = "Loading Database..."
ElseIf ProgressBar1.Value = 60 Then
Label7.Caption = "Loading Excel..."
ElseIf ProgressBar1.Value = 80 Then
Label7.Caption = "Loading Components..."
ElseIf ProgressBar1.Value = 90 Then
Label7.Caption = "Export Complete..."
ElseIf ProgressBar1.Value = 101 Then
MsgBox "Export Completed!!!, Locate your files in EXPORT Folder.", vbInformation, "Export Success!"
ProgressBar1.Value = 0
Timer1.Enabled = False
Label7.Caption = ""
btnExpMem.Enabled = True
btnExp.Enabled = True
btnExpBorrow.Enabled = True
btnExpReturn.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Dim a As Long
Timer2.Enabled = True
ProgressBar2.Max = 101
ProgressBar2.Value = ProgressBar2.Value + 1
btnBackUp.Enabled = False
If ProgressBar2.Value = 20 Then
Label1.Caption = "Preparing to Backup Files..."
ElseIf ProgressBar2.Value = 40 Then
Label1.Caption = "Loading Database..."
ElseIf ProgressBar2.Value = 60 Then
Label1.Caption = "Loading Access..."
ElseIf ProgressBar2.Value = 80 Then
Label1.Caption = "Loading Components..."
ElseIf ProgressBar2.Value = 90 Then
Label1.Caption = "Backup Complete..."
ElseIf ProgressBar2.Value = 101 Then
MsgBox "Backup Completed!!!, Locate your files in BACKUP Folder.", vbInformation, "Export Success!"
ProgressBar2.Max = 0
Timer2.Enabled = False
Label1.Caption = ""
btnBackUp.Enabled = True
End If
End Sub
