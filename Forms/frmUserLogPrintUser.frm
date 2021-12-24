VERSION 5.00
Begin VB.Form frmUserLogPrintUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Username"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Print by Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cboUsername 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin prjLMS.jcbutton btnPrint 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         buttonstyle     =   2
         font            =   "frmUserLogPrintUser.frx":0000
         backcolor       =   15199212
         caption         =   "Print"
         usemaskcolor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         buttonstyle     =   2
         font            =   "frmUserLogPrintUser.frx":0028
         backcolor       =   15199212
         caption         =   "Cancel"
         usemaskcolor    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmUserLogPrintUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnPrint_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserLog where Username='" & cboUsername.Text & "' Order by LogDate", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptUserLogPrintUser
            Set rptUserLogPrintUser.DataSource = dbRec
                .Show 1
            Set dbRec = Nothing
        End With
    Else
        MsgBox "No Record Found!", vbCritical, "Warning"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
modCon.Connected
DisableClose Me.hwnd


Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserAccount", dbCon, 3, 3
Do While Not dbRec.EOF
    cboUsername.AddItem dbRec!UserName
dbRec.MoveNext
Loop
End Sub
