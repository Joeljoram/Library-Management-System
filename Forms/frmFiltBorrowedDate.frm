VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFiltBorrowedDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Print by Due Date"
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
      Left            =   80
      TabIndex        =   4
      Top             =   1510
      Width           =   3255
      Begin prjLMS.jcbutton btnPrintDue 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Caption         =   "Print"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnCancelDue 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Caption         =   "Cancel"
         UseMaskCOlor    =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpDD 
         Height          =   285
         Left            =   120
         TabIndex        =   7
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
         Format          =   58916865
         CurrentDate     =   41864
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print by Date Borrowed"
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
      Left            =   60
      TabIndex        =   0
      Top             =   80
      Width           =   3255
      Begin MSComCtl2.DTPicker dtpDB 
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
         Format          =   58916865
         CurrentDate     =   41864
      End
      Begin prjLMS.jcbutton btnPrint 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Caption         =   "Print"
         UseMaskCOlor    =   -1  'True
      End
      Begin prjLMS.jcbutton btnCancel 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         Caption         =   "Cancel"
         UseMaskCOlor    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmFiltBorrowedDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me

End Sub

Private Sub btnCancelDue_Click()
Unload Me
End Sub

Private Sub btnPrint_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow where BorrowDate like'%" & dtpDB.Value & "%' Order by BorrowDate", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltBorrowDate
            Set rptFiltBorrowDate.DataSource = dbRec
                .Show 1
            Set dbRec = Nothing
        End With
    Else
        MsgBox "No record found on this date!", vbCritical, "Warning"
        Exit Sub
    End If
    
End Sub

Private Sub btnPrintDue_Click()
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Borrow where ExpirationDate like'%" & dtpDD.Value & "%' Order by ExpirationDate", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltBorrowDate
            Set rptFiltBorrowDate.DataSource = dbRec
                .Sections("Section4").Controls("Label12").Caption = "PRINT BY DUE DATE"
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
