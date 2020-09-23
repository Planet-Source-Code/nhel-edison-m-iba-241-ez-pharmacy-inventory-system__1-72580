VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Programmers 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11010
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer20 
      Interval        =   1000
      Left            =   5520
      Top             =   10440
   End
   Begin VB.Timer Timer19 
      Interval        =   1000
      Left            =   5880
      Top             =   10440
   End
   Begin VB.Timer Timer18 
      Interval        =   1000
      Left            =   5640
      Top             =   6480
   End
   Begin VB.Timer Timer17 
      Interval        =   1000
      Left            =   5880
      Top             =   6720
   End
   Begin VB.Timer Timer16 
      Interval        =   1000
      Left            =   5280
      Top             =   8520
   End
   Begin VB.Timer Timer15 
      Interval        =   1000
      Left            =   5640
      Top             =   8520
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6480
      TabIndex        =   83
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Timer Timer14 
         Interval        =   1000
         Left            =   3720
         Top             =   480
      End
      Begin VB.Timer Timer13 
         Interval        =   1000
         Left            =   2880
         Top             =   240
      End
      Begin VB.Image Image66 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":0000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label80 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ferdinne Borja"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   95
         Top             =   2760
         Width           =   4455
      End
      Begin VB.Label Label79 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   94
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Label78 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sta. Cruz, Laguna"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   93
         Top             =   4200
         Width           =   4455
      End
      Begin VB.Label Label77 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LD Information technology (LD-L1)"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   92
         Top             =   4920
         Width           =   4455
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gimik, hnging out with friends"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   91
         Top             =   5640
         Width           =   4455
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGE:"
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
         Left            =   600
         TabIndex        =   90
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   600
         TabIndex        =   89
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE:"
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
         Left            =   600
         TabIndex        =   88
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOBBIES:"
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
         Left            =   600
         TabIndex        =   87
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
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
         Left            =   480
         TabIndex        =   86
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image Image65 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":7EE9
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Image Image64 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":ED39
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   4800
      End
      Begin VB.Label Label82 
         BackColor       =   &H00000000&
         Caption         =   "Label32"
         Height          =   1815
         Left            =   480
         TabIndex        =   85
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label81 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   84
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image63 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":16D1D
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   4800
      End
      Begin VB.Image Image62 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":1ED01
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   4800
      End
      Begin VB.Image Image61 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":26CE5
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   4800
      End
      Begin VB.Image Image60 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":2ECC9
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   4800
      End
      Begin VB.Image Image59 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":36CAD
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   1800
      End
      Begin VB.Image Image58 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":3DAFD
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   1800
      End
      Begin VB.Image Image57 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":4494D
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   1800
      End
      Begin VB.Image Image56 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":4B79D
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Image Image54 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":525ED
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6480
      TabIndex        =   70
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Timer Timer12 
         Interval        =   1000
         Left            =   2880
         Top             =   240
      End
      Begin VB.Timer Timer11 
         Interval        =   1000
         Left            =   3720
         Top             =   480
      End
      Begin VB.Image Image55 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":5B2E1
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
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
         Left            =   480
         TabIndex        =   82
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOBBIES:"
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
         Left            =   600
         TabIndex        =   81
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE:"
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
         Left            =   600
         TabIndex        =   80
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   600
         TabIndex        =   79
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGE:"
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
         Left            =   600
         TabIndex        =   78
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Texting"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   77
         Top             =   5640
         Width           =   4455
      End
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LD Information technology (LD-L1)"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   76
         Top             =   4920
         Width           =   4455
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sta. Cruz, Laguna"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   75
         Top             =   4200
         Width           =   4455
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   74
         Top             =   3480
         Width           =   4455
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jona Torres"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   73
         Top             =   2760
         Width           =   4455
      End
      Begin VB.Image Image53 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":63FD5
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Image Image52 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":6AE25
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   1800
      End
      Begin VB.Image Image51 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":71C75
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   1800
      End
      Begin VB.Image Image50 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":78AC5
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   1800
      End
      Begin VB.Image Image49 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":7F915
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   4800
      End
      Begin VB.Image Image48 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":878F9
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   4800
      End
      Begin VB.Image Image47 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":8F8DD
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   4800
      End
      Begin VB.Image Image46 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":978C1
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   4800
      End
      Begin VB.Label Label70 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   72
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label69 
         BackColor       =   &H00000000&
         Caption         =   "Label32"
         Height          =   1815
         Left            =   480
         TabIndex        =   71
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image45 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":9F8A5
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   4800
      End
      Begin VB.Image Image44 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":A7889
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Image Image43 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":AE6D9
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6480
      TabIndex        =   57
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Timer Timer10 
         Interval        =   1000
         Left            =   3720
         Top             =   480
      End
      Begin VB.Timer Timer9 
         Interval        =   1000
         Left            =   2880
         Top             =   240
      End
      Begin VB.Image Image40 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":B6BE4
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leonard Brena"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   69
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   68
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sta. Cruz, Laguna"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   67
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LD Information technology (LD-L1)"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   66
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Texting"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   65
         Top             =   5760
         Width           =   4455
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGE:"
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
         Left            =   600
         TabIndex        =   64
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   600
         TabIndex        =   63
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE:"
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
         Left            =   600
         TabIndex        =   62
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOBBIES:"
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
         Left            =   600
         TabIndex        =   61
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
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
         Left            =   480
         TabIndex        =   60
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Image Image42 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":BF0EF
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Image Image41 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":C5F3F
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   4800
      End
      Begin VB.Label Label58 
         BackColor       =   &H00000000&
         Caption         =   "Label32"
         Height          =   1815
         Left            =   480
         TabIndex        =   59
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label57 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   58
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image39 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":CDF23
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   4800
      End
      Begin VB.Image Image38 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":D5F07
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   4800
      End
      Begin VB.Image Image37 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":DDEEB
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   4800
      End
      Begin VB.Image Image36 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":E5ECF
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   4800
      End
      Begin VB.Image Image35 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":EDEB3
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   1800
      End
      Begin VB.Image Image34 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":F4D03
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   1800
      End
      Begin VB.Image Image33 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":FBB53
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1800
      End
      Begin VB.Image Image32 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1029A3
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Image Image31 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":1097F3
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Timer Timer8 
      Interval        =   1000
      Left            =   5640
      Top             =   4440
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   5280
      Top             =   4440
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6480
      TabIndex        =   42
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Timer Timer6 
         Interval        =   1000
         Left            =   2880
         Top             =   240
      End
      Begin VB.Timer Timer5 
         Interval        =   1000
         Left            =   3720
         Top             =   480
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cooking"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   54
         Top             =   5760
         Width           =   4455
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LD Information technology (LD-L1)"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   53
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pangil, Laguna"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   52
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   51
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Randymar Talavera"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   50
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOBBIES:"
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
         Left            =   480
         TabIndex        =   49
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE:"
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
         Left            =   480
         TabIndex        =   48
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   480
         TabIndex        =   47
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGE:"
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
         Left            =   480
         TabIndex        =   46
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
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
         Left            =   480
         TabIndex        =   45
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Image Image30 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":111B8D
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Image Image29 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":119F27
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Image Image28 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":120D77
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1800
      End
      Begin VB.Image Image27 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":127BC7
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   1800
      End
      Begin VB.Image Image26 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":12EA17
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   1800
      End
      Begin VB.Image Image25 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":135867
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   4800
      End
      Begin VB.Image Image24 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":13D84B
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   4800
      End
      Begin VB.Image Image23 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":14582F
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   4800
      End
      Begin VB.Image Image22 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":14D813
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   4800
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label35 
         BackColor       =   &H00000000&
         Caption         =   "Label32"
         Height          =   1815
         Left            =   480
         TabIndex        =   43
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image21 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":1557F7
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Image Image20 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":15E2FA
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   4800
      End
      Begin VB.Image Image19 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1662DE
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1800
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   5880
      Top             =   2640
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5640
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6480
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3720
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2880
         Top             =   240
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nhel Edison M. Ibañez"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   41
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
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
         Left            =   600
         TabIndex        =   40
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Image Image8 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":16D12E
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Image Image13 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":173F7E
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   4800
      End
      Begin VB.Image Image7 
         Height          =   1575
         Left            =   600
         Picture         =   "Programmers.frx":17BF62
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Caption         =   "Label32"
         Height          =   1815
         Left            =   480
         TabIndex        =   38
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   37
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Playing my PS2 and PSP"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         Top             =   5760
         Width           =   4455
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LD Information technology (LD-L1)"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Paete Laguna"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Image Image17 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":184A65
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   4800
      End
      Begin VB.Image Image16 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":18CA49
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   4800
      End
      Begin VB.Image Image15 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":194A2D
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   4800
      End
      Begin VB.Image Image14 
         Height          =   600
         Left            =   2280
         Picture         =   "Programmers.frx":19CA11
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   4800
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOBBIES:"
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
         Left            =   480
         TabIndex        =   31
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Image Image12 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1A49F5
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   1800
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COURSE:"
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
         Left            =   480
         TabIndex        =   30
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Image Image11 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1AB845
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   1800
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   480
         TabIndex        =   29
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Image Image10 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1B2695
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1800
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGE:"
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
         Left            =   480
         TabIndex        =   28
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Image Image9 
         Height          =   600
         Left            =   480
         Picture         =   "Programmers.frx":1B94E5
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1800
      End
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3960
      TabIndex        =   26
      Text            =   "F:\blinknew\blink\INVENTORY PHARMACY\Project\wmpaud6.wav"
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   20
      Top             =   8640
      Width           =   2295
      Begin VB.Image Image6 
         Height          =   1575
         Left            =   120
         Picture         =   "Programmers.frx":1C0335
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   15
      Top             =   6600
      Width           =   2295
      Begin VB.Image Image5 
         Height          =   1575
         Left            =   120
         Picture         =   "Programmers.frx":1C821E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   4560
      Width           =   2295
      Begin VB.Image Image4 
         Height          =   1575
         Left            =   120
         Picture         =   "Programmers.frx":1D0F12
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
      Begin VB.Image Image3 
         Height          =   1575
         Left            =   120
         Picture         =   "Programmers.frx":1D941D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   600
      ScaleHeight     =   1995
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   480
      Width           =   2295
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   120
         Picture         =   "Programmers.frx":1E17B7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label87 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4200
      TabIndex        =   100
      Top             =   10440
      Width           =   495
   End
   Begin VB.Label Label85 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4200
      TabIndex        =   98
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label83 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4200
      TabIndex        =   96
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4200
      TabIndex        =   56
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   55
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label view11 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image18 
      Height          =   2415
      Left            =   6360
      Picture         =   "Programmers.frx":1EA2BA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label view1 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   2160
      Width           =   495
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   30
      Left            =   3960
      TabIndex        =   25
      Top             =   105
      Visible         =   0   'False
      Width           =   135
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   238
      _cy             =   53
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   14040
      Picture         =   "Programmers.frx":1FD1B9
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "prettyfairy_27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   24
      Top             =   9840
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   9960
      Width           =   2415
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Faye Borja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Jona Torres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "lei_11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Leonard Brena"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "stuffy_28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Randy Talavera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "chocobo_27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Nhel Edison Ibañez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   101
      Top             =   10440
      Width           =   495
   End
   Begin VB.Label Label86 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   99
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   97
      Top             =   6360
      Width           =   495
   End
End
Attribute VB_Name = "Programmers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
WindowsMediaPlayer1.URL = Text1.Text

End Sub

Private Sub Image1_Click()
Dim a

a = MsgBox("Are you Sure you want to Exit", vbYesNo + vbQuestion)
If a = vbYes Then
frmMenu.Show
Unload Me
End If
End Sub



Private Sub Label45_Click()
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame4.Visible = False
End Sub

Private Sub Label46_Click()
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame4.Visible = False
End Sub

Private Sub Label83_Click()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label84_Click()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
Frame5.Visible = False

End Sub

Private Sub Label85_Click()
Frame4.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label86_Click()
Frame4.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label87_Click()
Frame5.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Label88_Click()
Frame5.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
End Sub

Private Sub Timer1_Timer()
Label31.Visible = True
Label32.Visible = False


End Sub

Private Sub Timer15_Timer()
Label85.Visible = True
Label86.Visible = False
End Sub

Private Sub Timer16_Timer()
Label85.Visible = False
Label86.Visible = True
End Sub

Private Sub Timer17_Timer()
Label83.Visible = True
Label84.Visible = False
End Sub

Private Sub Timer18_Timer()
Label83.Visible = False
Label84.Visible = True
End Sub

Private Sub Timer19_Timer()
Label87.Visible = False
Label88.Visible = True
End Sub

Private Sub Timer2_Timer()
Label31.Visible = False
Label32.Visible = True


End Sub

Private Sub Timer20_Timer()
Label87.Visible = True
Label88.Visible = False
End Sub

Private Sub Timer3_Timer()
view1.Visible = True
view11.Visible = False

End Sub

Private Sub Timer4_Timer()
view11.Visible = True
view1.Visible = False
End Sub

Private Sub Timer5_Timer()
Label35.Visible = False
Label36.Visible = True
End Sub

Private Sub Timer6_Timer()
Label35.Visible = True
Label36.Visible = False
End Sub

Private Sub Timer7_Timer()
Label45.Visible = False
Label46.Visible = True
End Sub

Private Sub Timer8_Timer()
Label45.Visible = True
Label46.Visible = False
End Sub

Private Sub view1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub view11_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub


