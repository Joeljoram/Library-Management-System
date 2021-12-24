VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3585
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4680
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2118.137
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Password"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
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
      Left            =   1440
      MaxLength       =   18
      PasswordChar    =   "l"
      TabIndex        =   2
      ToolTipText     =   "Type your Password"
      Top             =   1920
      Width           =   3135
   End
   Begin prjLMS.jcbutton btnCancel 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
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
      Caption         =   "Exit"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtUsername 
      Height          =   360
      Left            =   1440
      MaxLength       =   18
      TabIndex        =   1
      ToolTipText     =   "Type your Username"
      Top             =   1515
      Width           =   3135
   End
   Begin prjLMS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   53
   End
   Begin prjLMS.jcbutton btnLog 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
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
      Caption         =   "Login"
      UseMaskCOlor    =   -1  'True
   End
   Begin prjLMS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   53
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Forgot Password?"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3120
      MouseIcon       =   "frmLogin.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2606
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   2400
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Login Attempts: "
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Account"
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
      TabIndex        =   5
      Top             =   120
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "frmLogin.frx":030A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Username and Password to Login."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2835
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Dim ans As Integer
If MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Confirm Exit") = vbYes Then
        End
    Else
        Exit Sub
End If
End Sub

Private Sub Form_Load()
modCon.Connected
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd
End Sub
Private Sub btnLog_Click()
Static ctr As Byte
Dim dbRec As New ADODB.Recordset
Dim sRole As String

dbRec.Open "Select * From UserAccount Where Username='" & LCase(txtUsername.Text) & "'", dbCon, adOpenStatic, adLockReadOnly
If dbRec.RecordCount <> 0 Then
    If LCase(txtPassword.Text) = dbRec!Password Then
        MsgBox "Username and Password Successfully Log In", vbInformation, "Log In"
            UserLog = txtUsername.Text
            sRole = dbRec!UserType
            Set dbRec2 = New ADODB.Recordset
            dbRec2.Open "Select * from UserLog", dbCon, 3, 3
            With dbRec2
                .AddNew
                .Fields("Username") = UserLog
                .Fields("LogDate") = Date
                .Fields("TimeLogin") = Time
                .Update
                 MainForm.Text1.Text = dbRec2.Fields(0)
            End With
            If sRole = "Administrator" Then
                MainForm.mnuUtil.Enabled = True
                MainForm.mnuFULHist = True
                MainForm.mnuUserMan = True
                MainForm.mnuUsertyp = True
                MainForm.mnuCat = True
                MainForm.mnuUsrLogHistory = True
                MainForm.Toolbar1.Buttons(4).Enabled = True
                MainForm.Toolbar1.Buttons(5).Enabled = True
                MainForm.Toolbar1.Buttons(8).Enabled = True
                MainForm.Toolbar1.Buttons(9).Enabled = True
            Else
                MainForm.mnuUtil.Enabled = False
                MainForm.mnuFULHist = False
                MainForm.mnuUserMan = False
                MainForm.mnuUsertyp = False
                MainForm.mnuCat = False
                MainForm.mnuUsrLogHistory = False
                MainForm.Toolbar1.Buttons(4).Enabled = False
                MainForm.Toolbar1.Buttons(5).Enabled = False
                MainForm.Toolbar1.Buttons(8).Enabled = False
                MainForm.Toolbar1.Buttons(9).Enabled = False
            End If
            Unload Me
            MainForm.Show
            MainForm.StatusBar1.Panels(2) = dbRec.Fields(5)
            MainForm.Text2.Text = dbRec.Fields(5)
            MainForm.StatusBar1.Panels(4) = dbRec.Fields(4)
            Set dbRec2 = Nothing
    Else
            ctr = ctr + 1
            Label4.Caption = ctr
            MsgBox "Invalid Password, Please Try again!", vbCritical, "Warning"
            txtUsername.Text = ""
            txtPassword.Text = ""
            txtUsername.SetFocus
    End If
Else
        ctr = ctr + 1
        Label4.Caption = ctr
        MsgBox "Invalid Login, Please Try again!", vbCritical, "Warning"
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.SetFocus
        dbRec.Close
    End If
    
    Set dbRec = Nothing
  
If ctr = 3 Then
       MsgBox "You have attempted to Login 3 Times in the system. Please call the Developer for Assistance.", vbCritical, "Warning Information"
            End
    End If

End Sub
Private Sub chkShow_Click()
If chkShow.Value = 1 Then
    txtPassword.PasswordChar = ""
    txtPassword.Font = "MS Sans Serif"
Else
    txtPassword.Font = "Wingdings"
    txtPassword.PasswordChar = "l"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = False
End Sub

Private Sub Label6_Click()
Me.Hide
frmRecover.Show
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = True
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
    txtPassword.FontItalic = False
    txtPassword.ForeColor = vbBlack
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call btnLog_Click
End If
End Sub

Private Sub txtUsername_Click()
    txtUsername.Text = ""
    txtUsername.FontItalic = False
    txtUsername.ForeColor = vbBlack
End Sub


