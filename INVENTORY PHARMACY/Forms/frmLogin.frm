VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3375
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1994.062
   ScaleMode       =   0  'User
   ScaleWidth      =   6013.948
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   855
      Left            =   4200
      Picture         =   "frmLogin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   855
      Left            =   2640
      Picture         =   "frmLogin.frx":08A3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2925
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   2925
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Picture         =   "frmLogin.frx":0C43
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to exit?", vbOKCancel) = vbOK Then
Unload Me
End If

End Sub

Private Sub cmdOK_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Dim flag As Boolean

rs.Open "Select Username, password from login", con, adOpenKeyset, adLockOptimistic

While rs.EOF <> True
If Me.txtUserName = rs!UserName And Me.txtPassword = rs!Password Then
flag = True
End If
rs.MoveNext
Wend
If flag = True Then
Me.Hide
sndPlaySound App.Path & "\Welcome.wav", SND_ASYNC
frmMenu.Show

Else
sndPlaySound App.Path & "\wrong.wav", SND_ASYNC
End If



End Sub

Private Sub Form_Load()
Main
txtUserName.Text = ""
txtPassword.Text = ""
End Sub

