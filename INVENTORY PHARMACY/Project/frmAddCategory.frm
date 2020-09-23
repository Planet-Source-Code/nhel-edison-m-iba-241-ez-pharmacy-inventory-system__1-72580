VERSION 5.00
Begin VB.Form frmAddCategory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2160
      TabIndex        =   2
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
         TabIndex        =   4
         Top             =   840
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
         TabIndex        =   3
         Top             =   1440
         Width           =   2925
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
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   6
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please supply Admin username and password in order  to add a new Category."
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
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   4215
      End
   End
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
      Picture         =   "frmAddCategory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Left            =   4440
      Picture         =   "frmAddCategory.frx":03A0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Picture         =   "frmAddCategory.frx":0801
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "frmAddCategory"
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
frmMaster.Show vbModal
Else
MsgBox "Invalid username or password", vbInformation, "Error"
End If
End Sub
