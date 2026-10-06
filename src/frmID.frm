VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmID 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ID"
   ClientHeight    =   8790
   ClientLeft      =   1410
   ClientTop       =   1020
   ClientWidth     =   12600
   Icon            =   "frmID.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdCerca 
      Caption         =   "Cerca"
      Height          =   285
      Left            =   11040
      TabIndex        =   130
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   128
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   127
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   126
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   125
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   124
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   0
      Left            =   7080
      TabIndex        =   123
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   0
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   122
      Top             =   3360
      Width           =   3855
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   1
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   121
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   1
      Left            =   7080
      TabIndex        =   120
      Text            =   "Text1"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   119
      Text            =   "Text1"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   1
      Left            =   5160
      TabIndex        =   118
      Text            =   "Text1"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   117
      Text            =   "Text1"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   116
      Text            =   "Text1"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   115
      Text            =   "Text1"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   2
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   114
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   2
      Left            =   7080
      TabIndex        =   113
      Text            =   "Text1"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   2
      Left            =   5880
      TabIndex        =   112
      Text            =   "Text1"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   2
      Left            =   5160
      TabIndex        =   111
      Text            =   "Text1"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   110
      Text            =   "Text1"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   109
      Text            =   "Text1"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   108
      Text            =   "Text1"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   3
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   107
      Top             =   4440
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   106
      Text            =   "Text1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   105
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   104
      Text            =   "Text1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   103
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   102
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   101
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   4
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   100
      Top             =   4800
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   4
      Left            =   7080
      TabIndex        =   99
      Text            =   "Text1"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   4
      Left            =   5880
      TabIndex        =   98
      Text            =   "Text1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   4
      Left            =   5160
      TabIndex        =   97
      Text            =   "Text1"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   96
      Text            =   "Text1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   94
      Text            =   "Text1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   5
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   93
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   5
      Left            =   7080
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   5
      Left            =   5880
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   90
      Text            =   "Text1"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   89
      Text            =   "Text1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   6
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   86
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   6
      Left            =   7080
      TabIndex        =   85
      Text            =   "Text1"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   6
      Left            =   5880
      TabIndex        =   84
      Text            =   "Text1"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   82
      Text            =   "Text1"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   7
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   5880
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   7
      Left            =   7080
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   7
      Left            =   5880
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   7
      Left            =   5160
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   74
      Text            =   "Text1"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   8
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   8
      Left            =   7080
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   8
      Left            =   5880
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   8
      Left            =   5160
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.ComboBox cmbOperatore 
      DataSource      =   "datDifetto"
      Height          =   315
      Index           =   9
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   6600
      Width           =   3855
   End
   Begin VB.TextBox txtO 
      Height          =   285
      Index           =   9
      Left            =   7080
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox txtQ 
      Height          =   285
      Index           =   9
      Left            =   5880
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtI 
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtTot 
      Height          =   285
      Left            =   5880
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   23
      Left            =   8040
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   22
      Left            =   2760
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   21
      Left            =   360
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   1440
      Width           =   780
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   20
      Left            =   11520
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   19
      Left            =   10680
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   18
      Left            =   9840
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   17
      Left            =   8760
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   16
      Left            =   7440
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   15
      Left            =   6840
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   14
      Left            =   6120
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   13
      Left            =   5400
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   12
      Left            =   4560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   11
      Left            =   3840
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   780
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   10
      Text            =   "01/01/00"
      Top             =   2280
      Width           =   780
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   6
      Left            =   11520
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   5
      Left            =   8760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtCal 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   10080
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":2DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":343A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":3AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":4132
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":4586
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":49DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmID.frx":4E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   8160
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   1111
      ButtonWidth     =   1323
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Reclamo Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su reclamo corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica il Reclamo Corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ripeti"
            Key             =   "G"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clienti"
            Key             =   "K"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prodotti"
            Key             =   "Y"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Marcia"
            Key             =   "W"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Traferisci"
            Key             =   "T"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   400
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape6 
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Inserire l'Id cercato"
      Height          =   255
      Left            =   8640
      TabIndex        =   129
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Totale"
      Height          =   255
      Left            =   5280
      TabIndex        =   58
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Operatore"
      Height          =   255
      Left            =   7560
      TabIndex        =   56
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Qta"
      Height          =   255
      Left            =   6000
      TabIndex        =   55
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Turno"
      Height          =   255
      Left            =   5160
      TabIndex        =   54
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Fine"
      Height          =   255
      Left            =   3840
      TabIndex        =   53
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Inizio"
      Height          =   255
      Left            =   2640
      TabIndex        =   52
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   255
      Left            =   1200
      TabIndex        =   51
      Top             =   3120
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label40 
      Caption         =   "Imballo"
      Height          =   255
      Left            =   8040
      TabIndex        =   50
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label39 
      Caption         =   "DataApertura"
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label38 
      Caption         =   "CDL"
      Height          =   255
      Left            =   3000
      TabIndex        =   48
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label37 
      Caption         =   "Alt"
      Height          =   255
      Left            =   11640
      TabIndex        =   44
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label36 
      Caption         =   "Sp"
      Height          =   255
      Left            =   10680
      TabIndex        =   43
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label35 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   9840
      TabIndex        =   42
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label34 
      Caption         =   "Mescola"
      Height          =   255
      Left            =   8760
      TabIndex        =   41
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label33 
      Caption         =   "Subbio"
      Height          =   255
      Left            =   7440
      TabIndex        =   40
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label32 
      Caption         =   "Bobine"
      Height          =   255
      Left            =   6840
      TabIndex        =   39
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "Qta M"
      Height          =   255
      Left            =   6120
      TabIndex        =   38
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label30 
      Caption         =   "QtaKG"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "Lunghezza"
      Height          =   255
      Left            =   4560
      TabIndex        =   36
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "Altezza"
      Height          =   255
      Left            =   3840
      TabIndex        =   35
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label27 
      Caption         =   "Spessore"
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "Finitura"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "Scadenza"
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "Promessa"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "Capitolato"
      Height          =   255
      Left            =   11640
      TabIndex        =   30
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Ragione Sociale"
      Height          =   255
      Left            =   9480
      TabIndex        =   29
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Capitolato"
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Descrizione Prodotto"
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Codice"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "ODL"
      Height          =   255
      Left            =   1320
      TabIndex        =   25
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "ID"
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Commessa in produzione"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   5550
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmID.frx":5282
      Top             =   120
      Width           =   450
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   12375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   4455
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   10695
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private mStatus As Integer
Private Enum CMDS
    cNuovo = 1
    cSalva = 3
    cAnnulla = 5
    cModifica = 7
    cChiudi = 9
    cCerca = 11
    cCliente = 13
    cProdotto = 15
    cMarcia = 17
    cStampa = 19
    cTrasferisci = 21
    cEsci = 23
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Dim rsComponente As Recordset, rsN As Recordset
Dim fS As Boolean, dTot As Double, dC As Double, dF As Double, dP As Double
Dim fSalva As Boolean
Dim sSe As String, sO As String 'sO=Codice Operaio
Public Property Let Status(vValue As Integer)
 Dim iPos As Integer, s As String
   s = Me.Key
   iPos = InStr(s, ",")
   If iPos <> 0 Then
      s = Mid$(s, iPos + 2, 5)
   End If
  
   
   
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "ATC" Then ' controllo che non possano modificare i dati dal MKG
        Select Case s
            Case "FATTA"
'        If s = "FATTA" Then
'            MsgBox "Fatta"
                .Buttons(cNuovo).Enabled = False
                .Buttons(cSalva).Enabled = False
                .Buttons(cAnnulla).Enabled = False
                .Buttons(cModifica).Enabled = False
                .Buttons(cChiudi).Enabled = False
                .Buttons(cStampa).Enabled = True
                .Buttons(cCerca).Enabled = True
                .Buttons(cCliente).Enabled = True
                .Buttons(cProdotto).Enabled = True
                .Buttons(cMarcia).Enabled = True
                .Buttons(cTrasferisci).Enabled = False
                .Buttons(cEsci).Enabled = True
            Case "TRASF"
                .Buttons(cNuovo).Enabled = False
                .Buttons(cSalva).Enabled = False
                .Buttons(cAnnulla).Enabled = False
                .Buttons(cModifica).Enabled = False
                .Buttons(cChiudi).Enabled = False
                .Buttons(cCliente).Enabled = True
                .Buttons(cProdotto).Enabled = True
                .Buttons(cMarcia).Enabled = True
                .Buttons(cStampa).Enabled = True
                .Buttons(cCerca).Enabled = True
                .Buttons(cTrasferisci).Enabled = False
                .Buttons(cEsci).Enabled = True
            
            Case Else
              
                Select Case mStatus
                   Case cNew, cMod
                        .Buttons(cNuovo).Enabled = False
                        .Buttons(cSalva).Enabled = True
                        .Buttons(cAnnulla).Enabled = True
                        .Buttons(cModifica).Enabled = False
                        .Buttons(cChiudi).Enabled = False
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cCliente).Enabled = True
                        .Buttons(cProdotto).Enabled = True
                        .Buttons(cMarcia).Enabled = True
                        .Buttons(cTrasferisci).Enabled = False
                        .Buttons(cEsci).Enabled = False
                   Case cExists
                        .Buttons(cNuovo).Enabled = False
                        .Buttons(cSalva).Enabled = False
                        .Buttons(cAnnulla).Enabled = False
                        .Buttons(cModifica).Enabled = True
                        .Buttons(cChiudi).Enabled = Not (sStato = "C")
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cCliente).Enabled = True
                        .Buttons(cProdotto).Enabled = True
                        .Buttons(cMarcia).Enabled = True
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cTrasferisci).Enabled = False
                        .Buttons(cEsci).Enabled = True
                        fS = True
                   Case cNotExists
                        .Buttons(cNuovo).Enabled = True
                        .Buttons(cSalva).Enabled = False
                        .Buttons(cAnnulla).Enabled = False
                        .Buttons(cModifica).Enabled = False
                        .Buttons(cChiudi).Enabled = False
                        .Buttons(cTrasferisci).Enabled = False
                        .Buttons(cCliente).Enabled = True
                        .Buttons(cProdotto).Enabled = True
                        .Buttons(cMarcia).Enabled = True
                        .Buttons(cCerca).Enabled = True
                        .Buttons(cStampa).Enabled = True
                        .Buttons(cEsci).Enabled = True
                        fS = False
                End Select
        End Select
      
    Else
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCliente).Enabled = True
            .Buttons(cProdotto).Enabled = True
            .Buttons(cMarcia).Enabled = True
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCliente).Enabled = True
            .Buttons(cProdotto).Enabled = True
            .Buttons(cMarcia).Enabled = True
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cCliente).Enabled = True
            .Buttons(cProdotto).Enabled = True
            .Buttons(cMarcia).Enabled = True
            .Buttons(cCerca).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    End If
    
    
    
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
    CercaID
    CercaComponenti
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub Rete()
   Dim lDiff As Long
    Load frmAggiunte
    With frmAggiunte
        .Top = Me.Top + (Me.Height - frmAggiunte.Height)
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
        .Key = Me.Key
        .Show vbModal
    End With
End Sub

Private Sub cmdCerca_Click()
CercaIdC
txtId.Enabled = False

End Sub

Private Sub Form_Load()
Dim btn As Button
    fS = False
    sStato = ""
   ListCodiceOperatore
    CampiT False
    'Nascondo tutti i pulsanti
'    cmdNew.Visible = False
'    cmdSalva.Visible = False
'    cmdAnnulla.Visible = False
'    cmdModifica.Visible = False
'   For Each btn In tbCommand.Buttons
'     If btn.Key <> "X" Then
'        btn.Enabled = False
'     End If
'   Next
   CampiC False
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
      l = shpRisposta.Left * 2 + shpRisposta.Width + Me.Width - Me.ScaleWidth
      If Me.Width < l Then
         Me.Width = l
      End If
      l = 0
      Me.Refresh
      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Load frmVenditori
'    frmVenditori.Show
'   rsATC.Close
'   Set rsATC = Nothin
   
'   rsComponente.Close
'   Set rsComponente = Nothing

End Sub
Private Sub CercaIDN()
Dim sqlc As String, i As Integer
'dTot = 0
'   sSe = ValID(Me.Key)
'
'    sqlc = "SELECT Formule.ID, Formule.Mescola, Form.MP, Form.Qta, " & _
'                "DescrizMesc.Descrizione " & _
            "FROM (Formule LEFT JOIN Form ON Formule.Mescola = " & _
                "Form.Mescola) LEFT JOIN DescrizMesc ON Form.MP " & _
                "= DescrizMesc.MP " & _
            "GROUP BY Formule.ID, Formule.Mescola, Form.MP, " & _
                "Form.Qta, DescrizMesc.Descrizione " & _
            "HAVING Formule.ID='" & sSe & "'"


'Set rsN = db.OpenRecordset(sqlc, dbOpenSnapshot)

            
'            For i = 0 To rsN.RecordCount - 1
'            If Not IsNull(rsN("MP")) Then
'                txtMP(i).Visible = True
'                txtMP(i).Text = rsN("MP")
'            Else
 '               txtMP(i) = ""
 '           End If
 '           If Not IsNull(rsN("Descrizione")) Then
'                txtMPD(i).Visible = True
'                txtMPD(i).Text = rsN("Descrizione")
'            Else
'                txtMPD(i).Text = ""
'            End If
'            If Not IsNull(rsN("Qta")) Then
'                txtMPP(i).Visible = True
'                txtMPP(i).Text = Trim$(rsN("Qta"))
'                dTot = dTot + txtMPP(i)
'            Else
 '               txtMPP(i).Text = ""
 '           End If
'''            If Not IsNull(rsN("Lotto")) Then
 '               txtLotto(i).Visible = True
 ''               txtLotto(i).Text = rsN("Lotto")
 ' '          Else
  '              txtLotto(i).Text = ""
  ''          End If
 '
'        rsN.MoveNext
 '       Next i
 '
 '  txtTot.Text = dTot

End Sub
Private Sub CercaComponenti()
Dim sqlc As String, i As Integer, dTot As Double
Dim rsT As Recordset, dSe As Double
sSe = ValID(Me.Key)
' If fNuovoRecord Then
'    sqlc = "SELECT DescrizMesc.Descrizione, Form.MP, Form.Qta, " & _
                "Formule.ID, '' AS Lotto " & _
            "FROM (Form INNER JOIN Formule ON Form.Mescola " & _
                "= Formule.Mescola) LEFT JOIN DescrizMesc ON " & _
                "Form.MP = DescrizMesc.MP " & _
            "WHERE Formule.ID='" & sSe & "'"
    
'Else
'    sqlc = "SELECT Formule.ID, Formule.SeID " & _
            "From Formule " & _
            "Where Formule.ID = '" & sSe & "'"
'    Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
'        If Not rsT.EOF Then
'            dSe = rsT("SeId")
'        End If
'    rsT.Close
'    Set rsT = Nothing
'
'    sqlc = "SELECT MP.MP, DescrizMesc.Descrizione, MP.Qta, " & _
                "MP.QtaC, MP.QtaT, MP.QtaP, MP.Lotto, SE.Ciclo, SE.Fattore, SE.Parte " & _
            "FROM ((Formule INNER JOIN MP ON Formule.SeID " & _
                "= MP.SeId) LEFT JOIN DescrizMesc ON MP.MP = " & _
                "DescrizMesc.MP) INNER JOIN SE ON Formule.SeID = SE.SeId " & _
            "GROUP BY MP.MP, DescrizMesc.Descrizione, MP.Qta, " & _
                "MP.QtaC, MP.QtaT, MP.QtaP, MP.Lotto, MP.SeId, " & _
                "SE.Ciclo, SE.Fattore, SE.Parte " & _
            "HAVING MP.SeId=" & dSe

'End If

'    Dim a As Integer
    
'    a = 0
'    Set rsComponente = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    If Not rsComponente.EOF Then
'        rsComponente.MoveLast
'
'        i = rsComponente.RecordCount
'        rsComponente.MoveFirst
'
'        For a = 0 To i - 1
'         If Not IsNull(rsComponente(a)) Then
'            Debug.Print (rsComponente(a))
'            End If
            
'        rsComponente.MoveNext
'        Next a
        

'
'        FoundID
'        Me.Status = cExists
'    Else
'        Me.Status = cNotExists
'    End If
End Sub
Private Sub FoundID()
'Dim i As Integer, dTot As Double
'dTot = 0
'dC = 0
'dF = 0
'dP = 0
'        For i = 0 To rsComponente.RecordCount - 1
'            If Not IsNull(rsComponente("MP")) Then
'                txtMP(i).Visible = True
'                txtMP(i).Text = rsComponente("MP")
'            Else
'                txtMP(i) = ""
'            End If
'            If Not IsNull(rsComponente("Descrizione")) Then
'                txtMPD(i).Visible = True
'                txtMPD(i).Text = rsComponente("Descrizione")
'            Else
'                txtMPD(i).Text = ""
 ''           End If
            'Controllo se già memorizzato
 '           If Not fNuovoRecord Then
''            If Not IsNull(rsComponente("Qta")) Then
 ''               txtMPP(i).Text = rsComponente("Qta")
'            Else
' '               If Not IsNull(rsComponente("Qta")) Then
'                    txtMPP(i).Text = rsComponente("Qta")
 '               Else
'                    txtMPP(i).Text = ""
' '               End If
''
'            End If
'            dTot = dTot + txtMPP(i)
'            txtMPP(i).Visible = True
'            If Not fNuovoRecord Then
''            If Not IsNull(rsComponente("Qta")) Then
'                txtFatt(i).Text = rsComponente("Qtac")
'            Else
'                If Not IsNull(rsComponente("QtaC")) Then
'                    txtFatt(i).Text = rsComponente("QtaC")
'                Else
'                    txtFatt(i).Text = ""
'                End If
'
'            End If
'            dF = dF + txtFatt(i)
'            txtFatt(i).Visible = True
'            If Not fNuovoRecord Then
''            If Not IsNull(rsComponente("Qta")) Then
'                txtFT(i).Text = rsComponente("QtaT")
'            Else
'                If Not IsNull(rsComponente("QtaT")) Then
'                    txtFT(i).Text = rsComponente("QtaT")
'                Else
'                    txtFT(i).Text = ""
'                End If
                
'            End If'
'            dC = dC + txtFT(i)
'            txtFT(i).Visible = True
'            If Not fNuovoRecord Then
''            If Not IsNull(rsComponente("Qta")) Then
'                txtPC(i).Text = rsComponente("QtaP")
'            Else
'                If Not IsNull(rsComponente("QtaP")) Then
'                    txtPC(i).Text = rsComponente("QtaP")
'                Else
'                    txtPC(i).Text = ""
'                End If
'
'            End If
'            If Not IsNull(txtPC(i).Text) And txtPC(i).Text <> "Text1" Then
'                dP = dP + txtPC(i)
'                txtPC(i).Visible = True
'            End If
'
'            If Not IsNull(rsComponente("Lotto")) Then
'                txtLotto(i).Visible = True
'                txtLotto(i).Text = rsComponente("Lotto")
'            Else
'                txtLotto(i).Text = ""
'            End If
'            txtF.Text = rsComponente("Fattore")
'            txtCicli.Text = rsComponente("Ciclo")
 '           txtParte.Text = rsComponente("Parte")'

'        rsComponente.MoveNext
'        Next i
'
'    txtTot.Text = dTot
'    txtSF.Text = dF
 '   txtSFT.Text = dC
 '   txtSPC.Text = dP
End Sub
Private Sub CercaID()
Dim sqlc As String, rsId As Recordset, i As Integer, rsT As Recordset
Dim dQ As Double
'Dim iPos As Integer
'   s = Me.Key
'   iPos = InStr(s, ",")
'   If iPos <> 0 Then
'      s = Left$(s, iPos - 1)
'   End If
'sSe = ValID(Me.Key)
'sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, Qta, Caricata, SeId " & _
'        "From Formule " & _
        "WHERE ID='" & sSe & "'"
'Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    If Not rsT.EOF Then
'        sSe = rsT("SeId")
 '   End If
''sqlc = "SELECT Formule.Mescola, Formule.Data, Formule.OdL, " & _
            "Formule.ID, Formule.CodArt, Formule.Qta, Formule.Caricata, Formule.SeID " & _
        "From Formule " & _
        "WHERE Formule.SeID=" & sSe
'sqlc = "SELECT Formule.Mescola, DescrizMesc.Descrizione, " & _
            "Formule.Data, Formule.OdL, Formule.ID, " & _
            "Formule.CodArt, Formule.Qta, Formule.Caricata, Formule.SeID " & _
        "FROM Formule LEFT JOIN DescrizMesc ON Formule.Mescola " & _
            "= DescrizMesc.MP " & _
        "WHERE Formule.SeID=" & sSe
       '
'Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
        
'Inizializziamo
'txtId.Text = ""
'txtODL.Text = ""
'txtQta.Text = ""
'dQ = 0
        '
'        Do While Not rsId.EOF
'
'            txtMescola.Text = rsId("Mescola")
'            If Not IsNull(rsId("Descrizione")) Then
'                txtDM.Text = rsId("Descrizione")    'Descrizione mescola
'            End If
 '           If Not IsNull(txtId.Text) And txtId.Text <> "" Then
'                txtId.Text = txtId.Text & ", " & rsId("ID")
 '           Else
'                txtId.Text = rsId("ID")
'            End If
'            If Not IsNull(txtODL.Text) And txtODL.Text <> "" Then
 '               txtODL.Text = txtODL.Text & ", " & rsId("ODL")
'            Else
'                txtODL.Text = rsId("ODL")
'            End If
 '           If Not IsNull(txtQta.Text) And txtQta.Text <> "" Then
 '               dQ = dQ + rsId("Qta")
 '           Else
 '               dQ = rsId("Qta")
 ''           End If
 '           txtQta.Text = dQ
  '
  '          txtSeId.Text = rsId("SeId")
  '          txtData.Text = Date
  '
  '      rsId.MoveNext
  '      Loop'

End Sub

Private Sub Campi(Flag As Boolean)
'    Dim i As Integer, dr As Double
'
'    If fS Then
 ''       dr = rsComponente.RecordCount - 1
 '   Else
 '       dr = txtMP.Count - 1
  ''  End If
  '  For i = 0 To dr
   ' txtMP(i).Visible = Flag
   ''     txtMPD(i).Visible = Flag
   '     txtMPP(i).Visible = Flag
   '     txtLotto(i).Visible = Flag
'  ''      txtCMP(i).Visible = Flag
'  '      txtCMPQ(i).Visible = Flag
   '     txtFatt(i).Visible = Flag
    '    txtFT(i).Visible = Flag
    ''    txtPC(i).Visible = Flag
        
   ' Next i
    
    
    
End Sub

Private Sub CancellaCampi()
End Sub

Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
    Select Case Button.Key
        Case "X"
            Unload Me
        Case "N"
            fNuovoRecord = True
            '         Campi True
            '         CancellaCampi
            CampiT True
            
'            CercaIDN
            Me.Status = cNew
'            CampiD True
        Case "S"
            Controllo
            If fSalva Then
'                SalvaMP
                Me.Status = cExists
                '        Campi False
'                CampiD False
            End If
            
        Case "A"
            If fNuovoRecord Then
'                CancellaCampi
                Me.Status = cNotExists
            Else
'                CercaComponenti
                Me.Status = cExists
            End If
'            '         Campi False'
'            CampiD False
        Case "M"
            fNuovoRecord = False
            Campi True
            Me.Status = cMod
'            Controtipo
'            CampiD True
        Case "C"
            iRes = MsgBox("Vuoi chiudere questa produzione?", _
            vbYesNo + vbDefaultButton2 + vbQuestion)
            If iRes = vbYes Then
'                sqlc = "Update Formule set Caricata = 'P' " & _
                    "Where SeId = " & txtSeId.Text & " and " & _
                    "Caricata = 'S'"
                db.Execute (sqlc)
'                sqlc = "Update SE set Qta = '" & txtSFT.Text & "' " & _
                        "where SeId = " & txtSeId.Text
                db.Execute (sqlc)
                
                sStato = "C"
                Me.Status = cExists
                frmTV.ChangedTV = True
            End If
        Case "P"
            fNuovoRecord = False
'            Rete
            Me.PrintForm
            Me.Status = cExists
        Case "G"
'            CercaIdC
            CampiC False
            CampiT False
            txtId.Text = ""
            txtId.Enabled = True
            txtId.SetFocus
'            Load frmAggiunte
'            frmAggiunte.Key = txtSeId.Text
'            frmAggiunte.Show
'            Unload Me
        
        Case "K"    'Cliente
'            CercaIdC
            
            Load frmCliente
'            frmAggiunte.Key = txtSeId.Text
            frmCliente.Show
'            Unload Me
        
        Case "Y"        'Prodotto
'            CercaIdC
            
            Load frmProdotto
'            frmAggiunte.Key = txtSeId.Text
            frmProdotto.Show
'            Unload Me
        
        Case "W"        'Marcia
'            CercaIdC
            
            Load frmMarcia
'            frmAggiunte.Key = txtSeId.Text
            frmMarcia.Show
'            Unload Me
        
        
        
        Case "T"
            iRes = MsgBox("Vuoi trasferire questa produzione?", _
            vbYesNo + vbDefaultButton2 + vbQuestion)
            If iRes = vbYes Then
'                sqlc = "Update Formule set Caricata = 'T' " & _
                    "Where SeId = " & txtSeId.Text & " and " & _
                    "Caricata = 'F'"
                db.Execute (sqlc)
'                sStato = "C"
                Me.Status = cExists
                frmTV.ChangedTV = True
            End If
'            Load frmExportMFG
'            frmExportMFG.Key = txtSeId.Text
'            frmExportMFG.Show
            
    End Select
End Sub

Private Sub Controtipo()
'Dim sqlc As String, rsCon As Recordset
'Dim sWhere As String, i As Integer
'
'    For i = 0 To rsComponente.RecordCount - 1
'        sWhere = txtMP(i).Text
'        sqlc = "SELECT MP, Controtipo " & _
'                "From Controtipo " & _
'                "WHERE MP='" & sWhere & "'"
'        Set rsCon = db.OpenRecordset(sqlc, dbOpenDynaset)
'            If Not rsCon.EOF Then
'                txtCMP(i).Text = rsCon("Controtipo")
'                txtCMPQ(i).Visible = True
'     '           txtCMPQ(i).Text = ""
'            Else
'                txtCMP(i).Text = ""
'                txtCMP(i).Visible = False
'                txtCMPQ(i).Visible = False
'            End If
'                txtCMPQ(i).Text = ""
'
'    Next i
'
End Sub
Private Sub CampiC(Flag As Boolean)
Dim i As Integer
For i = 0 To 24 - 1
txtCal(i).Visible = Flag
Next i




'Dim dr As Double, i As Integer
'    If fS Then
'        dr = rsComponente.RecordCount - 1
'    Else
'        dr = txtMP.Count - 1
'    End If
 '   For i = 0 To dr
 '   txtMP(i).Enabled = Flag
 ''       txtMPD(i).Enabled = Flag
 '       txtMPP(i).Enabled = Flag
 '       txtLotto(i).Enabled = Flag
''        txtCMP(i).Enabled = Flag
' '       txtCMPQ(i).Enabled = Flag
  ''      txtFatt(i).Enabled = Flag
  '      txtFT(i).Enabled = Flag
   '     txtPC(i).Enabled = Flag
   ' Next i
'
End Sub

Private Sub Controllo()

'If Not IsNull(txtF.Text) And txtF.Text <> "" Then
'    If Not IsNull(txtCicli.Text) And txtCicli.Text <> "" Then
'        ControlloQta
 '   Else
 '       fSalva = False
 ''       MsgBox "Prima di salvare bisogna inserire i dati relativi ai cicli", vbExclamation
'    End If
'Else
'    fSalva = False
'    MsgBox "Prima di salvare bisogna inserire i dati relativi ai " & _
            "fattori di moltiplicazione", vbExclamation
'End If


End Sub
Private Sub ControlloQta()
'Dim iRes As Integer, sqlc As String
'Dim dQta As Double, dSomma As Double, dEcc As Double'
''
'dQta = txtQta.Text * 1.1
'If Not IsNull(txtSFT.Text) And txtSFT.Text <> "" Then
'    dSomma = txtSFT.Text
'Else
'    fSalva = False
'    MsgBox "Inserire i cicli ed il fattore di moltiplicazione", vbCritical
'    Exit Sub
'End If'

'dEcc = txtQta.Text * 1.2
'    If dSomma < dQta Then
'        iRes = MsgBox("La quantità in preparazione è inferiore all'ordine " & _
'                    " del cliente più un 10%. Volete continuare?", _
'               vbYesNo + vbDefaultButton2 + vbQuestion)
'        If iRes = vbYes Then
'            fSalva = True
'        Else
'            fSalva = False
'        End If''

'    Else
'        If dSomma > dEcc Then
'            iRes = MsgBox("La quantità in preparazione è superiore all'ordine " & _
'                        " del cliente più un 20%. Volete continuare?", _
 '                  vbYesNo + vbDefaultButton2 + vbQuestion)
'            If iRes = vbYes Then
'                fSalva = True
 '           Else
'                fSalva = False
'            End If
'        Else
'            fSalva = True
'        End If
'
'    End If
End Sub

Private Sub CercaIdC()
Dim sqlc As String, rsT As Recordset
Dim i As Integer, nFields As Integer

i = 0
sqlc = "SELECT  ID,  ODL,  CodiceArticolo,  Articolo,  CapitolatoArticolo,  " & _
            "Cliente,  CapitolatoCliente,  DataP,  DataS,  Finitura,  " & _
            "Spessore,  Altezza,  Lunghezza,  QtaKg,  QtaOrdinata,  " & _
            "LunghezzaBobine,  Subbio,  CodiceM,  PTipo,  PSpessore,  " & _
            "PAltezza,  DataApertura,  CDL,  Imballo " & _
        "From ODL " & _
        "WHERE ID=" & txtId.Text
        
        
Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
   nFields = rsT.Fields.Count
    
      For i = 0 To nFields - 1
         Select Case rsT.Fields(i).Type
            
            Case dbText '
                If Not IsNull(rsT(i)) Then
                     txtCal(i) = rsT(i)
                 Else
                     txtCal(i) = ""
                 End If
            Case dbDate
                If Not IsNull(rsT(i)) Then
                     txtCal(i) = rsT(i)
                 Else
                     txtCal(i) = "01/01/00"
                 End If
            Case Else
                If Not IsNull(rsT(i)) Then
                     txtCal(i) = rsT(i)
                 Else
                     txtCal(i) = 0
                 End If
         End Select
        txtCal(i).Visible = True
        txtCal(i).Locked = True

      Next i
    

 
        
End Sub
Private Sub CampiT(Flag As Boolean)
Dim i As Integer
For i = 0 To 10 - 1
'txtCal(i).Visible = Flag
'Next i


txtD(i).Visible = Flag
txtD(i).Text = ""
txtI(i).Visible = Flag
txtI(i).Text = ""
txtF(i).Visible = Flag
txtF(i).Text = ""
txtT(i).Visible = Flag
txtT(i).Text = ""
txtQ(i).Visible = Flag
txtQ(i).Text = ""
txtO(i).Visible = False
cmbOperatore(i).Visible = Flag
Next i
txtTot.Visible = Flag



'Dim dr As Double, i As Integer
'    If fS Then
'        dr = rsComponente.RecordCount - 1
'    Else
'        dr = txtMP.Count - 1
'    End If
 '   For i = 0 To dr
 '   txtMP(i).Enabled = Flag
 ''       txtMPD(i).Enabled = Flag
 '       txtMPP(i).Enabled = Flag
 '       txtLotto(i).Enabled = Flag
''        txtCMP(i).Enabled = Flag
' '       txtCMPQ(i).Enabled = Flag
  ''      txtFatt(i).Enabled = Flag
  '      txtFT(i).Enabled = Flag
   '     txtPC(i).Enabled = Flag
   ' Next i
'
End Sub

Private Sub ListCodiceOperatore()
   Dim rsOperatore As Recordset, sqlc As String, sMsg As String
   Dim i As Integer
   sqlc = "SELECT IdOperaio, Cognome, Nome " & _
            "FROM Operaio"
   Set rsOperatore = db.OpenRecordset(sqlc, dbOpenSnapshot)
   For i = 0 To 10 - 1
   
    cmbOperatore(i).Clear
   Next i
   Do While Not rsOperatore.EOF
      For i = 0 To 10 - 1
        cmbOperatore(i).AddItem rsOperatore(1) & "  " & rsOperatore(2)
      Next i
      rsOperatore.MoveNext
      
   Loop
   rsOperatore.Close
   Set rsOperatore = Nothing
   For i = 0 To 10 - 1
    cmbOperatore(i).ListIndex = -1
   Next i
End Sub
Private Function FoundCodiceOperatore(cod As String)
   Dim rsOperatore As Recordset, sqlc As String
   sqlc = "SELECT IdOperaio, Cognome, Nome " & _
            "FROM Operaio " & _
            "where IdOperaio=" & sO

Set rsOperatore = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsOperatore.EOF Then
       FoundCodiceOperatore = rsOperatore("DescrizioneOperatore")
   Else
       FoundCodiceOperatore = ""
   End If
   rsOperatore.Close
   Set rsOperatore = Nothing
End Function

Private Function SalvaCodiceOperatore(s As String)
   Dim rsOperatore As Recordset, sqlc As String, sDescrizione As String
   Dim iPos As Integer
   iPos = InStr(s, "'")
   If iPos <> 0 Then
      s = Left$(s, iPos) & "'" & Right$(s, Len(s) - iPos)
      
   End If
   sqlc = "SELECT IdOperaio, Cognome, Nome " & _
            "FROM Operaio " & _
            "where Cognome=" & Trim$(s) & "'"
   Set rsOperatore = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsOperatore.EOF Then
      SalvaCodiceOperatore = rsOperatore("CodOperatore")
   Else
      SalvaCodiceOperatore = ""
   End If
   rsOperatore.Close
   Set rsOperatore = Nothing
End Function


