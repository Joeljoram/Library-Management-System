VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltMemSec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Section"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvSection 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6800
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Section"
         Object.Width           =   6068
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4200
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
            Picture         =   "frmFiltMemSec.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjLMS.jcbutton btnCancel 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3960
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
   Begin prjLMS.jcbutton btnPrint 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   3960
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
End
Attribute VB_Name = "frmFiltMemSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
frmFiltMem.Show 1

End Sub

Private Sub btnPrint_Click()
If Text1.Text = "" Then
MsgBox "Please select a record in the database to print.", vbCritical, "Warning"
Else
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member where [Section]='" & Text1.Text & "' Order by Name", dbCon, 3, 3
    If dbRec.RecordCount > 0 Then
        With rptFiltMemPrintSec
            Set rptFiltMemPrintSec.DataSource = dbRec
                .Show 1
            Set dbRec = Nothing
        End With
    End If
End If
End Sub

Private Sub Form_Load()
modCon.Connected
Call RefreshSection
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
DisableClose Me.hwnd
End Sub
Public Sub RefreshSection()
lvSection.ListItems.Clear
Set dbRec = New ADODB.Recordset
dbRec.Open "Select * from Member Order by [Section]", dbCon, adOpenForwardOnly, adLockPessimistic
    With dbRec
            Do While Not .EOF
                lvSection.ListItems.Add , , !Section, 1, 1
            .MoveNext
            Loop
            .Close
    End With
Set dbRec = Nothing
End Sub


Private Sub lvSection_Click()
Text1.Text = lvSection.SelectedItem.Text
End Sub

