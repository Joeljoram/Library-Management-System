VERSION 5.00
Begin VB.Form frmSystem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Library Management System Information"
   ClientHeight    =   4815
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   240
      Picture         =   "frmSystem.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   9735
      _extentx        =   17171
      _extenty        =   53
   End
   Begin prjLMS.jcbutton btnOk 
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      ToolTipText     =   "Close Dialog"
      Top             =   3840
      Width           =   1575
      _extentx        =   2778
      _extenty        =   873
      buttonstyle     =   2
      font            =   "frmSystem.frx":6B12
      backcolor       =   15199212
      caption         =   "Ok"
      usemaskcolor    =   -1  'True
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
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Designed for The Eldoret National Polytechnic             "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   690
      Index           =   1
      Left            =   5760
      TabIndex        =   8
      Top             =   2400
      Width           =   3285
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed and Developed by: Joel Kiptoo Deus joekiki05@gmail.com     "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   0
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   2955
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5640
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Beta Version 1.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Top             =   3240
      Width           =   1860
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmSystem.frx":6B3A
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Its all about the system and license."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Information"
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
      TabIndex        =   0
      Top             =   240
      Width           =   3060
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
Unload Me
Load MainForm
MainForm.Show
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd
End Sub
