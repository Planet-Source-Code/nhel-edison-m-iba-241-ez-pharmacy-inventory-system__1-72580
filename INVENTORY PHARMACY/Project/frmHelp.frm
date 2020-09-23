VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   12180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12180
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6855
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12091
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      FileName        =   "E:\final\blink\INVENTORY PHARMACY\help.rtf"
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   17160
      TabIndex        =   0
      Top             =   12000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      TabIndex        =   3
      Top             =   10920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   9840
      Picture         =   "frmHelp.frx":08DF
      Top             =   10800
      Width           =   6000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User'sManual"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.Image Image9 
      Height          =   12240
      Left            =   0
      Picture         =   "frmHelp.frx":88C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub Command3_Click()


End If


End Sub

