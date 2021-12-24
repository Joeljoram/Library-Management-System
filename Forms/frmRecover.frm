VERSION 5.00
Begin VB.Form frmRecover 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forgot Password"
   ClientHeight    =   3615
   ClientLeft      =   6750
   ClientTop       =   4485
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6585
   Begin VB.TextBox txtUsername 
      Height          =   360
      Left            =   2040
      MaxLength       =   18
      TabIndex        =   4
      ToolTipText     =   "Type your Mobile Number"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   18
      PasswordChar    =   "l"
      TabIndex        =   3
      ToolTipText     =   "Type your Name"
      Top             =   2205
      Width           =   4455
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnRecover 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Recover"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnCancel 
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Make sure your mobile number is registered in our database."
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Width           =   4065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2325
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type your Mobile No:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1365
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recover your Username and Password through sms gateway."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   4365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Recovery Gateway"
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
      Top             =   120
      Width           =   4650
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   360
      Picture         =   "frmRecover.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmRecover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
Load frmLogin
frmLogin.Show
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd
End Sub
