VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9945
   ClientLeft      =   3015
   ClientTop       =   1950
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   Picture         =   "frmLock.frx":0000
   ScaleHeight     =   99450
   ScaleMode       =   0  'User
   ScaleWidth      =   17475
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "EXIT SYSTEM"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   4920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   6720
      TabIndex        =   1
      Top             =   2640
      Width           =   5295
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System LOCKED!"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   540
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   4005
      End
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   54.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   4560
      Width           =   8535
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
modCon.Connected

End Sub

Private Sub Timer1_Timer()
'Call Randomize
Label1.ForeColor = RGB(Rnd() * 256, Rnd() * 256, Rnd() * 256)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from UserAccount where Password='" & txtPassword.Text & "'", dbCon, 3, 3
If dbRec.RecordCount <> 0 Then

    If txtPassword.Text = dbRec!Password Then
        Unload Me
        dbRec.Close
    Else
        'Call hlfocus(Text1)
        MsgBox "INVALID PASSWORD. Please type the correct password.", vbCritical, "Warning"
        txtPassword.Text = ""
        txtPassword.SetFocus
    
    End If
End If
End If

End Sub
