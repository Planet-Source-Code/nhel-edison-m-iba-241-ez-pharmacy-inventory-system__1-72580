VERSION 5.00
Begin VB.Form print 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   6360
      TabIndex        =   37
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5880
      Picture         =   "Print.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7440
      Picture         =   "Print.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame receipt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BALIK PO KAYO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   33
         Top             =   8520
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MARAMING SALAMAT PO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   8160
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   240
         Picture         =   "Print.frx":1194
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Farmacia De Borja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   31
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Sta.Cruz Laguna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN# 221-934-853-000VAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES TRANSACTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   28
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   2280
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   3360
         TabIndex        =   24
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         TabIndex        =   23
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "CASH"
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
         TabIndex        =   22
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   2040
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAID:"
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
         TabIndex        =   19
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE:"
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
         TabIndex        =   18
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Of Items:"
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
         TabIndex        =   17
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier:"
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
         TabIndex        =   16
         Top             =   6840
         Width           =   615
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   14
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   7200
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "THIS SERVES AS YOUR OFFICIAL RECEIPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   7800
         Width           =   4695
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   6480
         Width           =   615
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   2040
         TabIndex        =   8
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   ".  00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Top             =   4800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   ".  00"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   5160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   ".  00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   5640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   ".  00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   6000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   ".000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   6480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   ".  00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00004080&
      Height          =   1215
      Left            =   5760
      TabIndex        =   36
      Top             =   3360
      Width           =   3495
   End
End
Attribute VB_Name = "print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub


    
Private Sub cmdPrint_Click()
 
Label3.Visible = False
 cmdPrint.Visible = False
 cmdExit.Visible = False
    
    PrintForm
    
    cmdPrint.Visible = True
    cmdExit.Visible = True
    
End Sub

Private Sub Command1_Click()
frmSplash.Show
Me.Hide
End Sub

