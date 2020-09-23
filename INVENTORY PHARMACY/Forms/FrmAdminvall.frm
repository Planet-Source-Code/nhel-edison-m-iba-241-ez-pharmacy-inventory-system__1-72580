VERSION 5.00
Begin VB.Form FrmAdminvall 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Admin"
   ClientHeight    =   3300
   ClientLeft      =   2580
   ClientTop       =   4170
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Picture         =   "FrmAdminvall.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Picture         =   "FrmAdminvall.frx":03A0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "$"
         TabIndex        =   0
         Top             =   720
         Width           =   2925
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2925
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please supply Admin username and password to view sales."
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Width           =   3735
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
         TabIndex        =   6
         Top             =   840
         Width           =   1080
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   960
      End
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Picture         =   "FrmAdminvall.frx":0801
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "FrmAdminvall"
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

rs.Open "Select Username, password from Admin", con, adOpenKeyset, adLockOptimistic

While rs.EOF = False
If Me.txtUserName = rs!UserName And Me.txtPassword = rs!Password Then
flag = True
End If
rs.MoveNext
Wend
If flag = True Then
Unload Me
RptAllsales.Show vbModal
Else
MsgBox "Invalid username or password", vbInformation, "Error"
End If
End Sub
