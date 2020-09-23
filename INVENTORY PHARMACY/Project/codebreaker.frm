VERSION 5.00
Begin VB.Form frmprog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   645
   ClientLeft      =   7230
   ClientTop       =   5100
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1800
      Top             =   2880
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CBREAKER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   3240
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   1800
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   2040
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   2280
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   2520
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   2760
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   3000
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   1560
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   1320
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   1080
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   840
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   600
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   360
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   120
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   -120
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   -120
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chocobo MY innovation.. Thankz from all the forums that i enter... u make my program might and processful....
'AMA Computer University Sta.Cruz Campus...

Dim i As Integer

Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
i = 0
End Sub

Private Sub Timer1_Timer()

frmprog.Width = 3400
If i <> 15 Then
Shape2(i).FillColor = vbBlue
i = i + 1
Else

frmprog.Width = 6700

If i = 15 Then
Shape2(0).FillColor = vbRed
Shape2(1).FillColor = vbRed
Shape2(2).FillColor = vbRed
Shape2(3).FillColor = vbRed
Shape2(4).FillColor = vbRed
Shape2(5).FillColor = vbRed
Shape2(6).FillColor = vbRed
Shape2(7).FillColor = vbRed
Shape2(8).FillColor = vbRed
Shape2(9).FillColor = vbRed
Shape2(10).FillColor = vbRed
Shape2(11).FillColor = vbRed
Shape2(12).FillColor = vbRed
Shape2(13).FillColor = vbRed
Shape2(14).FillColor = vbRed

i = 0
End If
Programmers.Show
Unload Me
End If

End Sub
