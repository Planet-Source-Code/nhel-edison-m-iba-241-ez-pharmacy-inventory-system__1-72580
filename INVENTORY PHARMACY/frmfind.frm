VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   5265
   ClientLeft      =   3435
   ClientTop       =   2625
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmfind.frx":0000
      Left            =   2280
      List            =   "frmfind.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmfind.frx":0004
      Left            =   2280
      List            =   "frmfind.frx":0006
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtexdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtbal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtshelf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtpdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shelf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   120
         Picture         =   "frmfind.frx":0008
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   120
         Picture         =   "frmfind.frx":17F8
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   120
         Picture         =   "frmfind.frx":2FE8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   120
         Picture         =   "frmfind.frx":47D8
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   2400
         Picture         =   "frmfind.frx":5FC8
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   2400
         Picture         =   "frmfind.frx":6575
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   2400
         Picture         =   "frmfind.frx":6B22
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   2400
         Picture         =   "frmfind.frx":70CF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image10 
      Height          =   1215
      Left            =   600
      Picture         =   "frmfind.frx":767C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image close 
      Height          =   345
      Left            =   5640
      Picture         =   "frmfind.frx":7B43
      Stretch         =   -1  'True
      Top             =   360
      Width           =   360
   End
   Begin VB.Image Image9 
      Height          =   1215
      Left            =   600
      Picture         =   "frmfind.frx":D355
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
frmMenu.Show
End Sub

Private Sub Combo1_Click()
Frame1.Visible = 1
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select Drug Name from Master where Category = '" & Combo1 & "'", con, adOpenKeyset, adLockOptimistic
rs.Open "Select MfdDate, Qty, ExpDate, Shelf from Master where Category = '" & Combo2 & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF = False And rs.BOF <> True Then
Me.txtpdate = rs!MfdDate
Me.txtbal = rs!Qty
Me.txtexdate = rs!ExpDate
Me.txtshelf = rs!Shelf
End If
If Val(txtbal.Text) = 0 Then
MsgBox Me.Combo1 & " is not available in stock", vbInformation, "Stock Query"
Else: MsgBox "There are " & Val(txtbal.Text) & " " & Me.Combo1 & "(s) in stock", vbInformation, "Stock Query"
End If
Set con = Nothing

Set rs = Nothing
End Sub

Private Sub Combo1_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select DrugName from Master order by DrugName", con, adOpenKeyset, adLockOptimistic

While rs.EOF <> True And rs.BOF <> True
Combo1.AddItem rs!DrugName
rs.MoveNext
Wend
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Me.Combo1.Clear
End Sub

