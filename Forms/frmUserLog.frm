VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log In History"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "frmUserLog.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvUserLog 
      Height          =   3255
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Database Record"
      Top             =   1740
      Width           =   9370
      _ExtentX        =   16536
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserLog ID"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usertype"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Log Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time Login"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Time Log out"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTSearch 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      ToolTipText     =   "Select date covered"
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107937793
      CurrentDate     =   41847
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   9855
      _extentx        =   17383
      _extenty        =   53
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   9855
      _extentx        =   17383
      _extenty        =   53
   End
   Begin prjLMS.jcbutton btnClose 
      Height          =   495
      Left            =   8040
      TabIndex        =   3
      ToolTipText     =   "Close Dialog"
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmUserLog.frx":059A
      backcolor       =   15199212
      caption         =   "Close"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnSearch 
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      ToolTipText     =   "Search Record"
      Top             =   1320
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      buttonstyle     =   2
      font            =   "frmUserLog.frx":05C2
      backcolor       =   15199212
      caption         =   "Search"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintInd 
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "Reload Database"
      Top             =   5160
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmUserLog.frx":05EA
      backcolor       =   15199212
      caption         =   "Print Individual"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintDate 
      Height          =   495
      Left            =   6600
      TabIndex        =   12
      ToolTipText     =   "Reload Database"
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmUserLog.frx":0612
      backcolor       =   15199212
      caption         =   "Print by Date"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintUsertype 
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      ToolTipText     =   "Reload Database"
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmUserLog.frx":063A
      backcolor       =   15199212
      caption         =   "Print by Type"
      usemaskcolor    =   -1  'True
   End
   Begin prjLMS.jcbutton btnPrintUser 
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "Reload Database"
      Top             =   5160
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmUserLog.frx":0662
      backcolor       =   15199212
      caption         =   "Print by Username"
      usemaskcolor    =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User Log in As of     :"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   1350
      Width           =   1485
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
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1440
      TabIndex        =   7
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
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   360
      Picture         =   "frmUserLog.frx":068A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view all the login history of the user who access the system."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   4845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login History"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2925
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
Attribute VB_Name = "frmUserLog"
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

Private Sub btnRefresh_Click()
Call UserLog
End Sub

Private Sub btnPrintDate_Click()
frmUserLogPrintDate.Show 1
End Sub

Private Sub btnPrintInd_Click()
If Text1.Text = "" Then
MsgBox "Please select a record in the database to print.", vbCritical, "Warning"
Else
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserLog where UserLogID=" & CInt(Text1.Text) & " ", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptUserLogPrintInd
            Set rptUserLogPrintInd.DataSource = dbRec
                .Sections("Section1").Controls("Text3").DataField = "UserLogID"
                .Sections("Section1").Controls("Text2").DataField = "Username"
                .Sections("Section1").Controls("Text1").DataField = "UserType"
                .Sections("Section1").Controls("Text7").DataField = "LogDate"
                .Sections("Section1").Controls("Text4").DataField = "TimeLogin"
                .Sections("Section1").Controls("Text5").DataField = "TimeLogout"
                .Show 1
            Set dbRec = Nothing
        End With
    End If
End If

End Sub

Private Sub btnPrintUser_Click()
frmUserLogPrintUser.Show 1
End Sub

Private Sub btnPrintUsertype_Click()
frmUserLogPrintType.Show 1
End Sub

Private Sub btnSearch_Click()
lvUserLog.ListItems.Clear

Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Userlog where LogDate like'%" & DTSearch.Value & "%'", dbCon, 3, 3
    With dbRec
            Do While Not .EOF
                lvUserLog.ListItems.Add , , !UserLogID, 1, 1
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(1) = "" & !UserName
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(2) = "" & !UserType
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(3) = "" & !LogDate
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(4) = "" & !TimeLogin
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(5) = "" & !TimeLogout
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
        
End Sub

Private Sub Form_Load()
DTSearch.Value = Date
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

modCon.Connected
Call UserLog
  
DisableClose Me.hwnd
End Sub


Private Sub UserLog()
Dim dbRec As ADODB.Recordset

lvUserLog.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserLog Order by LogDate", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvUserLog.ListItems.Add , , !UserLogID, 1, 1
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(1) = "" & !UserName
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(2) = "" & !UserType
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(3) = "" & !LogDate
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(4) = "" & !TimeLogin
                lvUserLog.ListItems(lvUserLog.ListItems.Count).SubItems(5) = "" & !TimeLogout
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
'Set dbCon = Nothing
                
Label2.Caption = lvUserLog.ListItems.Count
End Sub

Private Sub lvUserLog_Click()
Text1.Text = lvUserLog.SelectedItem.Text

End Sub
