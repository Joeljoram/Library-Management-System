VERSION 5.00
Begin VB.Form frmRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   5775
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnRestore 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Restore Database"
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "Restore"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnCancel 
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Cancel restoration"
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "Cancel"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search the database source file to restore your system."
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   3870
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmRestore.frx":0000
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restore Database"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2865
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can restore your database record on this section."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   3780
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
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd
End Sub
