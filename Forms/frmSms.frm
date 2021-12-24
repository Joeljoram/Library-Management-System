VERSION 5.00
Begin VB.Form frmSms 
   Caption         =   "Message"
   ClientHeight    =   8175
   ClientLeft      =   8130
   ClientTop       =   3630
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   5775
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   1920
      MaxLength       =   18
      PasswordChar    =   "l"
      TabIndex        =   13
      ToolTipText     =   "Type your Name"
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "CAPTCHA"
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2640
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      ToolTipText     =   "Type your Name"
      Top             =   2400
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   18
      TabIndex        =   1
      ToolTipText     =   "Type your Name"
      Top             =   1965
      Width           =   3735
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3480
      MaxLength       =   18
      TabIndex        =   0
      ToolTipText     =   "Type your Mobile Number"
      Top             =   1320
      Width           =   2175
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnRecover 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   7560
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Send"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   -720
      TabIndex        =   4
      Top             =   7440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type the words:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   7035
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   690
   End
   Begin VB.Image Image4 
      Height          =   720
      Left            =   360
      Picture         =   "frmSms.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Message Gateway"
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
      TabIndex        =   8
      Top             =   120
      Width           =   4125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text a message using your PC"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type your Mobile No:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1365
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   555
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
Attribute VB_Name = "frmSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "0906"
Combo1.AddItem "0907"
Combo1.AddItem "0908"
Combo1.AddItem "0909"
Combo1.ListIndex = 0

End Sub
