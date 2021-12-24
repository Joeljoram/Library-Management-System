VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUserLogPrintDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Print by Date Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   30
      TabIndex        =   0
      Top             =   50
      Width           =   3255
      Begin MSComCtl2.DTPicker dtpDL 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   107216897
         CurrentDate     =   41864
      End
      Begin prjLMS.jcbutton btnPrint 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         buttonstyle     =   2
         font            =   "frmUserLogPrintDate.frx":0000
         backcolor       =   15199212
         caption         =   "Print"
         usemaskcolor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         buttonstyle     =   2
         font            =   "frmUserLogPrintDate.frx":0028
         backcolor       =   15199212
         caption         =   "Cancel"
         usemaskcolor    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmUserLogPrintDate"
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
dbRec.Open "Select * from UserLog where LogDate like'%" & dtpDL.Value & "%' Order by Username", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptUserLogPrintDate
            Set rptUserLogPrintDate.DataSource = dbRec
                .Show 1
            Set dbRec = Nothing
        End With
    Else
        MsgBox "No record found on this date!", vbCritical, "Warning"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
modCon.Connected
DisableClose Me.hwnd
End Sub
