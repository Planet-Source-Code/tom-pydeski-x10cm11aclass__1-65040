VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCM11A 
   Caption         =   "Tom Py's X10 CM11A"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "X10CM11AClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Interface Data"
      Height          =   1870
      Left            =   4080
      TabIndex        =   143
      Top             =   600
      Width           =   3615
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         Height          =   1450
         Left            =   90
         TabIndex        =   144
         Top             =   250
         Width           =   3450
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   6120
      Top             =   0
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "Initialize X10 Class"
      Height          =   615
      Left            =   4080
      TabIndex        =   142
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   45
      TabIndex        =   126
      Top             =   -75
      Width           =   3975
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   127
         Top             =   240
         Width           =   3495
         Begin VB.VScrollBar VScroll1 
            Height          =   1215
            LargeChange     =   10
            Left            =   3120
            Max             =   0
            Min             =   100
            TabIndex        =   141
            Top             =   360
            Value           =   100
            Width           =   255
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "All Lights Off"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   140
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "All Lights On"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   139
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cmbHouseCode 
            Height          =   360
            Left            =   0
            TabIndex        =   133
            Text            =   "HC"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cmbDeviceCode 
            Height          =   360
            ItemData        =   "X10CM11AClass.frx":074A
            Left            =   1200
            List            =   "X10CM11AClass.frx":074C
            TabIndex        =   132
            Text            =   "DC"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtData2 
            Height          =   285
            Left            =   2400
            TabIndex        =   131
            Text            =   "50"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TxtDimVal 
            Height          =   285
            Left            =   2400
            TabIndex        =   130
            Text            =   "50"
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cmbCommand 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "X10CM11AClass.frx":074E
            Left            =   0
            List            =   "X10CM11AClass.frx":0750
            TabIndex        =   129
            Text            =   "cmbCommand"
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send Exec"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   128
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Data2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   138
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Dim Value (%) (data1):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   137
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Command:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   136
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "DeviceCode:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   135
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Housecode:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   134
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.Frame PortFrame 
      Caption         =   "Com Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5280
      TabIndex        =   124
      Top             =   0
      Width           =   855
      Begin VB.TextBox txtComPort 
         Height          =   285
         Left            =   240
         TabIndex        =   125
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
   End
   Begin RichTextLib.RichTextBox txtEvent 
      Height          =   10215
      Left            =   7800
      TabIndex        =   2
      Top             =   240
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   18018
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"X10CM11AClass.frx":0752
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   12
      Left            =   5880
      TabIndex        =   112
      Top             =   7800
      Width           =   1815
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1360
         TabIndex        =   119
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   840
         TabIndex        =   118
         Text            =   "3"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   12
         Left            =   1500
         TabIndex        =   122
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   116
         Text            =   "B"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   12
         Left            =   90
         Picture         =   "X10CM11AClass.frx":07CD
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   12
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   114
         Text            =   "X10CM11AClass.frx":159F
         Top             =   160
         Width           =   1215
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         ItemData        =   "X10CM11AClass.frx":15AC
         Left            =   120
         List            =   "X10CM11AClass.frx":15E3
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   840
         TabIndex        =   121
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   120
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   11
      Left            =   3960
      TabIndex        =   102
      Top             =   7800
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         ItemData        =   "X10CM11AClass.frx":16A3
         Left            =   120
         List            =   "X10CM11AClass.frx":16DA
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   11
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   108
         Text            =   "X10CM11AClass.frx":179A
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   11
         Left            =   90
         Picture         =   "X10CM11AClass.frx":17A7
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   720
         Width           =   1300
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   106
         Text            =   "B"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   11
         Left            =   1500
         TabIndex        =   27
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   840
         TabIndex        =   104
         Text            =   "2"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1360
         TabIndex        =   103
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   111
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   110
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   10
      Left            =   2040
      TabIndex        =   92
      Top             =   7800
      Width           =   1815
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   10
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   99
         Text            =   "X10CM11AClass.frx":2579
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   10
         Left            =   90
         Picture         =   "X10CM11AClass.frx":2589
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   720
         Width           =   1300
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   97
         Text            =   "B"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   10
         Left            =   1500
         TabIndex        =   37
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   840
         TabIndex        =   95
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1360
         TabIndex        =   94
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         ItemData        =   "X10CM11AClass.frx":335B
         Left            =   120
         List            =   "X10CM11AClass.frx":3392
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   101
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   100
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   9
      Left            =   120
      TabIndex        =   82
      Top             =   7800
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         ItemData        =   "X10CM11AClass.frx":3452
         Left            =   120
         List            =   "X10CM11AClass.frx":3489
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1360
         TabIndex        =   88
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   840
         TabIndex        =   87
         Text            =   "9"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   9
         Left            =   1500
         TabIndex        =   46
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   85
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   9
         Left            =   90
         Picture         =   "X10CM11AClass.frx":3549
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   83
         Text            =   "X10CM11AClass.frx":431B
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   91
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   90
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   8
      Left            =   5880
      TabIndex        =   72
      Top             =   5160
      Width           =   1815
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   8
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   79
         Text            =   "X10CM11AClass.frx":4333
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   8
         Left            =   90
         Picture         =   "X10CM11AClass.frx":433F
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   720
         Width           =   1300
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   77
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   8
         Left            =   1500
         TabIndex        =   56
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   840
         TabIndex        =   75
         Text            =   "8"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1360
         TabIndex        =   74
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         ItemData        =   "X10CM11AClass.frx":5111
         Left            =   120
         List            =   "X10CM11AClass.frx":5148
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   81
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   80
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   7
      Left            =   3960
      TabIndex        =   62
      Top             =   5160
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         ItemData        =   "X10CM11AClass.frx":5208
         Left            =   120
         List            =   "X10CM11AClass.frx":523F
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1360
         TabIndex        =   68
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   840
         TabIndex        =   67
         Text            =   "7"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   7
         Left            =   1500
         TabIndex        =   66
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   65
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   7
         Left            =   90
         Picture         =   "X10CM11AClass.frx":52FF
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   63
         Text            =   "X10CM11AClass.frx":60D1
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   71
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   6
      Left            =   2040
      TabIndex        =   52
      Top             =   5160
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         ItemData        =   "X10CM11AClass.frx":60DA
         Left            =   120
         List            =   "X10CM11AClass.frx":6111
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1360
         TabIndex        =   58
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   840
         TabIndex        =   57
         Text            =   "6"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   6
         Left            =   1500
         TabIndex        =   76
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   6
         Left            =   90
         Picture         =   "X10CM11AClass.frx":61D1
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   6
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   53
         Text            =   "X10CM11AClass.frx":6FA3
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   61
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   5
      Left            =   120
      TabIndex        =   42
      Top             =   5160
      Width           =   1815
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   49
         Text            =   "X10CM11AClass.frx":6FB0
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   5
         Left            =   90
         Picture         =   "X10CM11AClass.frx":6FBB
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   720
         Width           =   1300
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   47
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   5
         Left            =   1500
         TabIndex        =   86
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   840
         TabIndex        =   45
         Text            =   "5"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1360
         TabIndex        =   44
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         ItemData        =   "X10CM11AClass.frx":7D8D
         Left            =   120
         List            =   "X10CM11AClass.frx":7DC4
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   50
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   4
      Left            =   5880
      TabIndex        =   32
      Top             =   2520
      Width           =   1815
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1360
         TabIndex        =   39
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   840
         TabIndex        =   38
         Text            =   "4"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   4
         Left            =   1500
         TabIndex        =   96
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   4
         Left            =   90
         Picture         =   "X10CM11AClass.frx":7E84
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "X10CM11AClass.frx":8C56
         Top             =   160
         Width           =   1215
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         ItemData        =   "X10CM11AClass.frx":8C60
         Left            =   120
         List            =   "X10CM11AClass.frx":8C97
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   3
      Left            =   3960
      TabIndex        =   22
      Top             =   2520
      Width           =   1815
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1360
         TabIndex        =   29
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   840
         TabIndex        =   28
         Text            =   "3"
         Top             =   1680
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   3
         Left            =   1500
         TabIndex        =   105
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   3
         Left            =   90
         Picture         =   "X10CM11AClass.frx":8D57
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "X10CM11AClass.frx":9B29
         Top             =   160
         Width           =   1215
      End
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "X10CM11AClass.frx":9B37
         Left            =   120
         List            =   "X10CM11AClass.frx":9B6E
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   31
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   2
      Left            =   2040
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         ItemData        =   "X10CM11AClass.frx":9C2E
         Left            =   120
         List            =   "X10CM11AClass.frx":9C65
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "X10CM11AClass.frx":9D25
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdEvent 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   2
         Left            =   90
         Picture         =   "X10CM11AClass.frx":9D31
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1300
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Text            =   "2"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1360
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   2
         Left            =   1500
         TabIndex        =   117
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame frmX10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
      Begin VB.ComboBox cmbX10Com 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "X10CM11AClass.frx":AB03
         Left            =   120
         List            =   "X10CM11AClass.frx":AB3A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtDim 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1360
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDevice 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "X10CM11AClass.frx":ABFA
         Left            =   840
         List            =   "X10CM11AClass.frx":ABFC
         TabIndex        =   8
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox cmbHouse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "X10CM11AClass.frx":ABFE
         Left            =   120
         List            =   "X10CM11AClass.frx":AC00
         TabIndex        =   7
         Text            =   "A"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdEvent 
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Index           =   1
         Left            =   90
         Picture         =   "X10CM11AClass.frx":AC02
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtDeviceName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "X10CM11AClass.frx":B9D4
         Top             =   160
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   1335
         Index           =   1
         Left            =   1500
         TabIndex        =   123
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.Label lblDevice 
         Caption         =   "Device:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblHouse 
         Caption         =   "House:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.CommandButton ButClrCM11 
      Caption         =   "Clear CM11A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearText 
      Caption         =   "Clear Events"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image imgOn 
      Height          =   510
      Index           =   1
      Left            =   8640
      Picture         =   "X10CM11AClass.frx":B9E2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgOff 
      Height          =   510
      Index           =   1
      Left            =   8520
      Picture         =   "X10CM11AClass.frx":C7B4
      Top             =   960
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgOff 
      Height          =   480
      Index           =   0
      Left            =   6360
      Picture         =   "X10CM11AClass.frx":D586
      Top             =   1080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image imgOn 
      Height          =   480
      Index           =   0
      Left            =   6360
      Picture         =   "X10CM11AClass.frx":E548
      Top             =   1560
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label4 
      Caption         =   "X10 Events:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   45
      Width           =   975
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mSetHouse 
         Caption         =   "Set &House Code"
         Begin VB.Menu mHouse 
            Caption         =   "A"
            Index           =   0
         End
      End
      Begin VB.Menu mBar 
         Caption         =   "-"
      End
      Begin VB.Menu mShow 
         Caption         =   "Show Extended Controls"
      End
      Begin VB.Menu mMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&TrayMenu"
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu mData 
         Caption         =   "&Interface Data"
      End
   End
End
Attribute VB_Name = "frmCM11A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'X10 CM11 pc interface
'submitted by Tom Pydeski
'X10 put out some nifty home control products that communicate
'over the power lines and respond to commands to turn on; turn off; dim; etc.
'The CM11A is device that interfaces a pc to the command line protocol and also
'contains timers and macros.
'This project is a class I wrote to emulate what the keware ocx does.
'Keware (http://www.homeseer.com/downloads/index.htm)
'It can control devices in any housecode from your pc.
'The HouseCode can be any 1 of 16 (A through P) and each house code can handle
'any of 16 devices (1 to 16).
'I started out with the firecracker class I wrote
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64719&lngWId=1
'and I experimented with it and took information from the x10 protocol file
'I made some switches and used a vertical progress bar to set the dim level.
'I implemented a device status for each device.  Finally, I was able to implement
'the all on and all off commands.
'Of course all of this is useless if you don't have the x10 hardware that it
'interfaces with and controls (http://www.x10.com/automation/ck18a_s_ps32.html)
'(They were practically giving away the firecracker starter kit a few years back.)
'www.x10.com
'
'
'comm port parameters
'4800,N,8,1
'
'The housecodes and device codes range from A to P and 1 to 16 respectively although they do not follow a binary sequence. The encoding format for these codes is as follows
'
'HouseDeviceBinary
'A  1   0110    6
'B  2   1110    E
'C  3   0010    2
'D  4   1010    A
'E  5   0001    1
'F  6   1001    9
'G  7   0101    5
'H  8   1101    D
'I  9   0111    7
'J  10  1111    E
'K  11  0011    3
'L  12  1011    B
'M  13  0000    0
'N  14  1000    8
'O  15  0100    4
'P  16  1100    C
'
'1.2 Function Codes.
'
'Function           Binary Value
'All Units Off      0000    0
'All Lights On      0001    1
'On                 0010    2
'Off                0011    3
'Dim                0100    4
'Bright             0101    5
'All Lights Off     0110    6
'Extended Code      0111    7
'Hail Request       1000    8
'Hail Acknowledge   1001    9
'Pre-set Dim (1)    1010    10
'Pre-set Dim (2)    1011    11
'Extended Transfer  1100    12
'Status On          1101    13
'Status Off         1110    14
'Status Request     1111    15
'
'
'
'2.  Serial Parameters.
'
'The serial parameters for communications between the interface and PC are as follows:
'
'   Baud Rate:  4,800bps
'   Parity:     None
'   Data Bits:  8
'   Stop Bits:  1
'
'Cable connections:
'
'   Signal  DB9 Connector   RJ11 Connector
'    SIN        Pin 2       Pin 1
'    SOUT       Pin 3       Pin 3
'    GND        Pin 5       Pin 4
'    RI         Pin 9       Pin 2
'
'where:     SIN Serial input to PC (output from the interface)
'           SOUT    Serial output from PC (input to the interface)
'           GND Signal ground
'           RI  Ring signal (input to PC)
'
'
'A single normal command takes eleven cycles of the AC line to finish.
'All legal commands must first start with the header 1110, a unique code as described below.
'The header bits take two cycles at one bit per half cycle.
'The next four cycles are the four-bit House Code, but it takes eight bits total
'because each bit is sent true then complemented.
'This is similar to biphase encoding, as the bit value changes
'state half-way through the transmission, and improves transmission reliability.
'The last five AC cycles are the Unit / Function Code,
'a five bit code that takes ten bits (again, true then complemented).
'For any codes except the DIM, BRIGHT and the data following the
' EXTENDED DATA function, there's a mandatory three cycle pause
'before sending another command DIM and BRIGHT don't necessarily
'need a pause, and the data after the EXTENDED DATA command
'absolutely MUST follow immediately until all bytes have been sent.
'The EXTENDED DATA code is handy, as any number of eight-bit bytes may follow.
'The data bytes must follow the true/complement rule, so will take eight cycles per byte,
'with no pause between bytes until complete.
'The only legal sequence that doesn't conform to the true/complement rule are the
'start bits 1110 that lead the whole thing off,
'likely because the modules need some way to tell when it's OK to start listening again.
'
'A full transmission containing everything looks like this (see the end of this section
'for the actual command codes):
'1 1 1 0  H8 /H8 H4 /H4 H2 /H2 H1 /H1  D8 /D8 D4 /D4 D2 /D2 D1 /D1 F /F
'(start)        (House code)                 (Unit/Function code)
'So, to turn on Unit 12 of House code A, send the following:
'
'1 1 1 0    01 10 10 01     10 01 10 10 01 (House A, Unit 12)
'
'then wait at least three full AC cycles and send it again,
'1 1 1 0    01 10 10 01     10 01 10 10 01 (House A, Unit 12)
'
'then wait three and send:
'
'1 1 1 0    01 10 10 01     01 01 10 01 10 (House A, Function ON)
'
'again wait three cycles and send it the last time. Total transmission would have been 264 discrete bits (don't forget the 3-phase) and 'would take 53 cycles of the AC line, or about .883 seconds.
'
'=================================
'below for house code A
'Off Code
'c
'Dec = 6 99
'Hex = 6 63
'Bin = 0000 0110  0110 0011
'            ^^    ^^   ^^
'          HouseA Dev1  Off
'On Code
'b
'Dec = 6 98
'Hex = 6 62
'Bin = 0000 0110  0110 0010
'            ^^    ^^   ^^
'          HouseA Dev1  On
'
'==============================================
'
'House Code A - Device Code 1
'f
'Dec = 4 102
'Hex = 4 66
'Bin = 0000 0100  0110 0110
'
'House Code A - Device Code 2
'n
'Dec = 4 110
'Hex = 4 6E
'Bin = 0000 0100  0110 1110
'
'House Code A - Device Code 3
'b
'Dec = 4 98
'Hex = 4 62
'Bin = 0000 0100  0110 0010
'
'House Code A - Device Code 4
'j
'Dec = 4 106
'Hex = 4 6A
'Bin = 0000 0100  0110 1010
'
'House Code A - Device Code 5
'a
'Dec = 4 97
'Hex = 4 61
'Bin = 0000 0100  0110 0001
'
'House Code A - Device Code 6
'i
'Dec = 4 105
'Hex = 4 69
'Bin = 0000 0100  0110 1001
'
'House Code A - Device Code 7
'e
'Dec = 4 101
'Hex = 4 65
'Bin = 0000 0100  0110 0101
'
'House Code A - Device Code 8
'm
'Dec = 4 109
'Hex = 4 6D
'Bin = 0000 0100  0110 1101
'
'House Code A - Device Code 9
'g
'Dec = 4 103
'Hex = 4 67
'Bin = 0000 0100  0110 0111
'
'House Code A - Device Code 10
'o
'Dec = 4 111
'Hex = 4 6F
'Bin = 0000 0100  0110 1111
'
'House Code A - Device Code 11
'c
'Dec = 4 99
'Hex = 4 63
'Bin = 0000 0100  0110 0011
'
'House Code A - Device Code 12
'k
'Dec = 4 107
'Hex = 4 6B
'Bin = 0000 0100  0110 1011
'
'House Code A - Device Code 13
'`
'Dec = 4 96
'Hex = 4 60
'Bin = 0000 0100  0110 0000
'
'House Code A - Device Code 14
'h
'Dec = 4 104
'Hex = 4 68
'Bin = 0000 0100  0110 1000
'
'House Code A - Device Code 15
'd
'Dec = 4 100
'Hex = 4 64
'Bin = 0000 0100  0110 0100
'
'=========================
'
'House Code A - Device Code 1
'f
'Dec = 4 102
'Hex = 4 66
'Bin = 0000 0100  0110 0110
'
'House Code B - Device Code 1
'æ
'Dec = 4 230
'Hex = 4 E6
'Bin = 0000 0100  1110 0110
'
'â
'Dec = 6 226
'Hex = 6 E2
'Bin = 0000 0110  1110 0010
'            ^^   ^^
'           HouseB  On
'
'ã
'Dec = 6 227
'Hex = 6 E3
'Bin = 0000 0110  1110 0011
'            ^^   ^^
'           HouseB Off
'
Dim LastDim(16) As Integer
Dim IgnScrl As Byte
Dim WithEvents X10CM11 As clsX10CM11
Attribute X10CM11.VB_VarHelpID = -1
Const AckStr$ = "Ã" 'Chr$(195)
Dim IgnoreHouse As Byte
Dim IgnoreKeys As Byte
Dim h As Integer
Dim i As Integer
Dim F As Integer
Dim lRet As Long
Dim UseClass As Byte
Dim Released As Byte
'hand cursor
Dim lHandle As Long
Const HandCursor = 32649&
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Dim Inits As Byte

Private Sub Form_Click()
X10CM11.GetStatus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    DimDown
End If
If KeyCode = vbKeyUp Then
    DimUp
End If
End Sub

Private Sub Form_Load()
On Error GoTo Oops
With Me
    .Left = GetSettingIni(App.Title, "Settings", "Left", 0)
    .Top = GetSettingIni(App.Title, "Settings", "Top", 0)
End With
Set X10CM11 = New clsX10CM11
X10CM11.HouseCode = Val(GetSettingIni(App.Title, "Settings", "HouseCode"))
X10CM11.DeviceCode = Val(GetSettingIni(App.Title, "Settings", "DeviceCode"))
'
WindowState = vbMaximized
IgnScrl = 1
X10Init = 0
txtEvent.SelText = "In order to poll the PC, the interface will continually send: 5Ah = 90 = ""Z""" & vbCrLf
txtEvent.SelText = "The PC must send an acknowledgment: C3h = 195 = ""Ã""" & vbCrLf
'setup combo box to list all commands
cmbCommand.Clear
cmbCommand.List(0) = "All Units Off"
cmbCommand.List(1) = "All Lights On"
cmbCommand.List(2) = "On"
cmbCommand.List(3) = "Off"
cmbCommand.List(4) = "Dim"
cmbCommand.List(5) = "Bright"
cmbCommand.List(6) = "All Lights Off"
cmbCommand.List(7) = "Extended"
cmbCommand.List(8) = "Hail Request"
cmbCommand.List(9) = "Hail Ack"
cmbCommand.List(10) = "Pre-set dim1"
cmbCommand.List(11) = "Pre-set dim2"
cmbCommand.List(12) = "Extended Data"
cmbCommand.List(13) = "Status On"
cmbCommand.List(14) = "Status Off"
cmbCommand.List(15) = "Status Request"
cmbCommand.ListIndex = 2
For i = 0 To cmbCommand.ListCount - 1
    'Debug.Print "cmbCommand.List("; i; ") = """; cmbCommand.List(i)
Next i
For i = 0 To 15
    LastDim(i) = 0
Next
IgnScrl = 0
'
X10CM11.HouseCode = Val(GetSettingIni(App.Title, "Settings", "HouseCode"))
X10CM11.DeviceCode = Val(GetSettingIni(App.Title, "Settings", "DeviceCode"))
For h = 0 To 15
    X10(h).Configured = False
    For i = 1 To 16
        X10(h).DeviceName(i) = GetSettingIni(App.Title, "HouseCode" & h, "DeviceName" & i)
        If Len(Trim$(X10(h).DeviceName(i))) > 0 Then
            X10(h).Configured = True
        Else
            X10(h).DeviceName(i) = Chr$(Asc("A") + h) & i
        End If
    Next i
Next h
DoEvents
'setup combo box to list all possible devices
cmbDeviceCode.Clear
For i = 0 To 16
    cmbDeviceCode.AddItem i
Next i
cmbDeviceCode.ListIndex = X10CM11.DeviceCode
'
'setup combo box to list all possible houses
cmbHouseCode.Clear
cmbHouseCode.AddItem mHouse(0).Caption
For i = 1 To 15
    Load mHouse(i)
    mHouse(i).Caption = Chr$(Asc("A") + i)
    cmbHouseCode.AddItem mHouse(i).Caption
Next i
mHouse(X10CM11.HouseCode).Checked = True
IgnoreHouse = 1
cmbHouseCode.ListIndex = X10CM11.HouseCode
'
For i = 1 To 9
    txtDeviceName(i).Text = X10(0).DeviceName(i)
Next i
For i = 10 To txtDeviceName.UBound
    txtDeviceName(i).Text = X10(1).DeviceName(i - 9)
Next i
'
SetupHouse X10CM11.HouseCode
'get the status of all of the devices
ChDir App.Path
F = FreeFile
Open "X10Status.dat" For Random As #F Len = Len(X10Out(0))
For h = 0 To 15
    Get #F, h + 1, X10Out(h)
Next h
Close #F
'
If Val(Command$) > 0 Then
    'if we pass a device number in the command string,
    'then just send the command and exit
    cmdEvent_Click Val(Command$)
    DoEvents
    mExit_Click
End If
Show
UpdateDevice
GoTo Exit_Form_Load
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Form_Load "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in Form_Load"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Form_Load:
cmdInit_Click
'load our cute little hand cursor
'this was from the sample by LaVolpe (thanks!)
'at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63065&lngWId=1
lHandle = LoadCursor(0, HandCursor)
Inits = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If IgnoreKeys = 1 Then Exit Sub
If KeyAscii = Asc("Q") Or KeyAscii = Asc("q") Then
    Unload Me
    End
End If
If KeyAscii >= 48 And KeyAscii <= 57 Then '0-9
    X10CM11.DeviceCode = KeyAscii - 48
    GoTo Toggle
End If
Select Case UCase(Chr$(KeyAscii))
    Case "A"
        X10CM11.DeviceCode = 10
    Case "B"
        X10CM11.DeviceCode = 11
    Case "C"
        X10CM11.DeviceCode = 12
    Case "D"
        X10CM11.DeviceCode = 13
    Case "E"
        X10CM11.DeviceCode = 14
    Case "F"
        X10CM11.DeviceCode = 15
    Case Else
        Exit Sub
End Select
Toggle:
'toggle the selected device
X10Out(X10CM11.HouseCode).Device(X10CM11.DeviceCode) = 1 - X10Out(X10CM11.HouseCode).Device(X10CM11.DeviceCode)
If X10Out(X10CM11.HouseCode).Device(X10CM11.DeviceCode) = 1 Then
    'turn the device on
    XCommand = C_ON
Else
    'turn the device off
    XCommand = C_OFF
End If
UpdateDevice
Timer1.Enabled = False
X10CM11.Exec cmbHouseCode.Text, Str$(X10CM11.DeviceCode), XCommand
'delay so we don't fire again
Sleep (1000)
Timer1.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Released = 1
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
With txtEvent
    If Me.Width > .Left + 200 Then
        .Width = (Me.Width - .Left) - 200
    End If
    If Me.Height > .Top + 850 Then
        .Height = (Me.Height - .Top) - 850
    End If
End With
txtEvent.Width = txtEvent.Width
txtEvent.Height = txtEvent.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = 11
SaveSettingIni App.Title, "Settings", "Left", Me.Left
SaveSettingIni App.Title, "Settings", "Top", Me.Top
SaveSettingIni App.Title, "Settings", "HouseCode", X10CM11.HouseCode
SaveSettingIni App.Title, "Settings", "DeviceCode", X10CM11.DeviceCode
'save the devicenames for all of our housecodes
ChDir App.Path
F = FreeFile
Open "X10Labels.ini" For Random As #F Len = Len(X10(0))
For h = 0 To 15
    Put #F, h + 1, X10(h)
Next h
Close #F
'
F = FreeFile
ChDir App.Path
Open "X10Status.dat" For Random As #F Len = Len(X10Out(0))
For h = 0 To 15
    Put #F, h + 1, X10Out(h)
Next h
Close #F
'
h = X10CM11.HouseCode
'For h = 0 To 15
For i = 1 To 16
    SaveSettingIni App.Title, "HouseCode" & h, "DeviceName" & i, X10(h).DeviceName(i)
Next i
'Next h
End
End Sub

Private Sub cmbHouseCode_Change()
If IgnoreHouse = 1 Then Exit Sub
If Inits = 0 Then Exit Sub
X10CM11.HouseCode = cmbHouseCode.ListIndex + 1
SetupHouse cmbHouseCode.ListIndex
End Sub

Private Sub cmbHouseCode_Click()
If IgnoreHouse = 1 Then Exit Sub
If Inits = 0 Then Exit Sub
X10CM11.HouseCode = cmbHouseCode.ListIndex + 1
SetupHouse cmbHouseCode.ListIndex
End Sub

Private Sub cmbDeviceCode_Change()
If Inits = 0 Then Exit Sub
X10CM11.DeviceCode = cmbDeviceCode.ListIndex
End Sub

Private Sub cmbDeviceCode_Click()
If Inits = 0 Then Exit Sub
X10CM11.DeviceCode = cmbDeviceCode.ListIndex
End Sub

Private Sub cmdAll_Click(Index As Integer)
cmbCommand.ListIndex = Index
i = cmbCommand.ListIndex
Timer1.Enabled = False
X10CM11.Exec cmbHouseCode.Text, cmbDeviceCode.Text, i, Val(TxtDimVal) ', Val(TxtDimVal), Val(TxtData2)
If Index = 0 Then
    UpdateDevice
End If
Timer1.Enabled = True
End Sub

Private Sub cmdInit_Click()
Dim err As Integer
Dim init_error As Integer
Screen.MousePointer = 11
cmdInit.Caption = "Wait..."
cmdInit.Enabled = False
UseClass = 1
'-------------------------------------------------------------
X10CM11.ComPort = Val(txtComPort)
init_error = X10CM11.Init
If init_error <> 0 Then
    MsgBox "Error initializing CM11", vbExclamation + vbOKOnly
    GoTo ExitInit
End If
cmdSend.Enabled = True
cmdAll(0).Enabled = True
cmdAll(1).Enabled = True
ExitInit:
Timer1.Enabled = True
Screen.MousePointer = 0
End Sub

Private Sub ButClrCM11_Click()
Confirm = MsgBox("Are you sure you want to clear the interface memory?", vbYesNo + vbQuestion, "X-10 Interface")
'If Confirm = vbYes Then ctlX10CM11.ClearMem
'
End Sub

Private Sub cmdSend_Click()
Dim i As Integer
i = cmbCommand.ListIndex
If i = 16 Then
    i = -1
End If
Timer1.Enabled = False
X10CM11.Exec cmbHouseCode.Text, cmbDeviceCode.Text, i, Val(TxtDimVal) ', Val(TxtDimVal), Val(TxtData2)
Timer1.Enabled = True
End Sub

Private Sub lblStatus_Click()
X10CM11.GetStatus
End Sub

Private Sub mData_Click()
MsgBox StatMess$
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub cmdClearText_Click()
txtEvent = ""
End Sub

Private Sub pbDim_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewVal As Integer
If Button = 1 Then
    With pbDim(Index)
        NewVal = ((.Height - Y) / .Height) * 100
        If NewVal > 0 And NewVal < 100 Then
            .Value = NewVal
            .ToolTipText = .Value
            ChangeDim (Index)
        End If
    End With
End If
End Sub

Private Sub pbDim_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
Dim NewVal As Integer
With pbDim(Index)
    NewVal = ((.Height - Y) / .Height) * 100
    If Button = 1 Then
        If NewVal > 0 And NewVal < 100 Then
            .Value = NewVal
        End If
    End If
    .ToolTipText = NewVal
End With
End Sub

Sub ChangeDim(Index As Integer)
If pbDim(Index).Value > LastDim(Index) Then
    cmbX10Com(Index).ListIndex = 5 'bright
Else
    cmbX10Com(Index).ListIndex = 4 'dim
End If
txtDim(Index).Text = pbDim(Index).Value
If X10Init = 0 Then Exit Sub
If IgnScrl = 1 Then Exit Sub
Timer1.Enabled = False
If X10Out(X10CM11.HouseCode).Device(X10CM11.DeviceCode) = 0 Then
    'I'm not sure bright works right...It just turns the light on fully
    X10CM11.Exec cmbHouse(Index), cmbDevice(Index), 4, pbDim(Index).Value
Else
    X10CM11.Exec cmbHouse(Index), cmbDevice(Index), cmbX10Com(Index).ListIndex, pbDim(Index).Value - LastDim(Index)
End If
LastDim(Index) = pbDim(Index).Value
'
Timer1.Enabled = True
X10Out(X10CM11.HouseCode).Device(X10CM11.DeviceCode) = 1
UpdateDevice
End Sub

Sub DimUp()
DoEvents
Debug.Print "DimUp Pressed"
Do
    DoEvents
    If Released = 1 Then Exit Do
    Timer1.Enabled = False
    X10CM11.Exec cmbHouseCode.Text, Str$(X10CM11.DeviceCode), C_BRIGHT
    Timer1.Enabled = True
    Sleep 100
    DoEvents
    Sleep 100
Loop
Released = 0
Debug.Print "DimUp Released"
End Sub

Sub DimDown()
DoEvents
Debug.Print "DimDown Pressed"
Do
    DoEvents
    If Released = 1 Then Exit Do
    Timer1.Enabled = False
    X10CM11.Exec cmbHouseCode.Text, Str$(X10CM11.DeviceCode), C_DIM
    Timer1.Enabled = True
    Sleep 100
    DoEvents
    Sleep 100
Loop
Released = 0
Debug.Print "DimDown Released"
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
X10CM11.GetStatus
End Sub

Private Sub txtDeviceName_Change(Index As Integer)
If IgnoreHouse = 1 Then Exit Sub
X10(X10CM11.HouseCode).DeviceName(Index) = txtDeviceName(Index).Text
SaveSettingIni App.Title, "HouseCode" & X10CM11.HouseCode, "DeviceName" & Index, X10(X10CM11.HouseCode).DeviceName(Index)
End Sub

Private Sub txtDeviceName_GotFocus(Index As Integer)
'ignore keystrokes when we are focused on a text box
IgnoreKeys = 1
End Sub

Private Sub txtDeviceName_LostFocus(Index As Integer)
'ok...now we can act on keystrokes again
IgnoreKeys = 0
End Sub

Private Sub TxtDimVal_GotFocus()
'ignore keystrokes when we are focused on a text box
IgnoreKeys = 1
End Sub

Private Sub TxtDimVal_LostFocus()
'ok...now we can act on keystrokes again
IgnoreKeys = 0
End Sub

Private Sub VScroll1_Change()
TxtDimVal.Text = VScroll1.Value
End Sub

Private Sub cmbDevice_DropDown(Index As Integer)
cmbDevice(Index).Clear
For i = 1 To 16
    cmbDevice(Index).AddItem (i)
Next i
End Sub

Private Sub cmbHouse_DropDown(Index As Integer)
cmbHouse(Index).Clear
For i = 1 To 16
    cmbHouse(Index).AddItem Chr$(64 + i)
Next i
End Sub

Private Sub cmbX10Com_DropDown(Index As Integer)
cmbX10Com(Index).Clear
For i = 0 To cmbCommand.ListCount - 1
    cmbX10Com(Index).AddItem cmbCommand.List(i)
Next i
End Sub

Private Sub cmdEvent_Click(Index As Integer)
Dim HC As Integer
Dim DC As Integer
' 0 =All Units Off
' 1 =All Lights On
' 2 =On
' 3 =Off
' 4 =Dim
' 5 =Bright
' 6 =All Lights Off
' 7 =Extended
' 8 =Hail Request
' 9 =Hail Ack
' 10 =Pre-set dim1
' 11 =Pre-set dim2
' 12 =Extended Data
' 13 =Status On
' 14 =Status Off
' 15 =Status Request
IgnScrl = 1
HC = Asc(cmbHouse(Index).Text) - 65
DC = Val(cmbDevice(Index).Text)
X10CM11.HouseCode = HC
X10CM11.DeviceCode = DC
X10Out(HC).Device(DC) = 1 - X10Out(HC).Device(DC)
pbDim(Index).Value = X10Out(HC).Device(DC) * 100
If X10Out(HC).Device(DC) = 1 Then
    cmdEvent(Index).Picture = imgOn(1).Picture
    cmbX10Com(Index).ListIndex = 2 'on
Else
    cmdEvent(Index).Picture = imgOff(1).Picture
    cmbX10Com(Index).ListIndex = 3 'off
End If
Debug.Print cmbHouse(Index).Text; cmbDevice(Index).Text; " "; cmbX10Com(Index).Text; " "; pbDim(Index).Value; " "; pbDim(Index).Value; " "; Val(TxtData2)
Timer1.Enabled = False
X10CM11.Exec cmbHouse(Index).Text, cmbDevice(Index).Text, cmbX10Com(Index).ListIndex, pbDim(Index).Value
Timer1.Enabled = True
UpdateDevice
IgnScrl = 0
End Sub

Private Sub cmdEvent_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
End Sub

Sub SetupHouse(NewHouseIn As Integer)
IgnoreHouse = 1
X10CM11.HouseCode = NewHouseIn ' cmbHouse.ListIndex
If X10CM11.HouseCode < 0 Then
    X10CM11.HouseCode = 0
End If
For i = 0 To mHouse.UBound
    mHouse(i).Checked = False
Next i
mHouse(X10CM11.HouseCode).Checked = True
'setup combo box
cmbHouseCode.ListIndex = X10CM11.HouseCode
cmbHouseCode.Text = mHouse(X10CM11.HouseCode).Caption
'display new labels
For i = 1 To 9
    txtDeviceName(i).Text = X10(0).DeviceName(i)
    txtDeviceName(i).ToolTipText = "A" & cmbDevice(i).Text & ":" & txtDeviceName(i).Text
Next i
For i = 10 To txtDeviceName.UBound
    txtDeviceName(i).Text = X10(1).DeviceName(i - 9)
    txtDeviceName(i).ToolTipText = "B" & cmbDevice(i - 9).Text & ":" & txtDeviceName(i).Text
Next i
IgnoreHouse = 0
End Sub

Sub UpdateDevice()
On Error GoTo Oops
Dim DevLabNum As Integer
'Debug.Print "UpdateDevice -> ";
For i = 1 To 9
    txtDeviceName(i).BackColor = &H80000005
    'frmX10(i).BackColor = &H80000005
    If X10Out(0).Device(i) = 1 Then
        If LastDim(i) = 0 Then pbDim(i).Value = 100
        cmdEvent(i).Picture = imgOn(1).Picture
        Refresh
    Else
        pbDim(i).Value = 0
        cmdEvent(i).Picture = imgOff(1).Picture
    End If
    'Debug.Print X10Out(0).Device(i); " ";
Next i
For i = 10 To txtDeviceName.UBound
    txtDeviceName(i).BackColor = &H80000005
    'frmX10(i).BackColor = &H80000005
    If X10Out(1).Device(i - 9) = 1 Then
        pbDim(i).Value = 100
        cmdEvent(i).Picture = imgOn(1).Picture
        Refresh
    Else
        pbDim(i).Value = 0
        cmdEvent(i).Picture = imgOff(1).Picture
    End If
    'Debug.Print X10Out(1).Device(i - 9); " ";
Next i
Debug.Print
DevLabNum = X10CM11.DeviceCode
txtDeviceName(DevLabNum).BackColor = vbYellow
'frmX10(DevLabNum).BackColor = vbYellow
cmbDeviceCode.ListIndex = X10CM11.DeviceCode
Refresh
GoTo Exit_UpdateDevice
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine UpdateDevice "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in UpdateDevice"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
EMess$ = EMess$ & vbCrLf & "HouseCode = " & X10CM11.HouseCode
EMess$ = EMess$ & vbCrLf & "DeviceCode = " & X10CM11.DeviceCode
Alarm
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_UpdateDevice:
End Sub

Private Sub X10CM11_Initialized()
cmdInit.Caption = "Initialized"
End Sub

Private Sub X10CM11_X10Event(Devices As String, HouseCode As String, Command As Integer, Extra As String, Data2 As String)
txtEvent.Text = txtEvent.Text & "Event: " & HouseCode & Devices & " " & cmbCommand.List(Command)
If Val(Extra) > 0 Then
    txtEvent.Text = txtEvent.Text & " Ex = " & Extra
End If
If Val(Data2) > 0 Then
    txtEvent.Text = txtEvent.Text & " Data2= " & Data2
End If
txtEvent.Text = txtEvent.Text & vbCrLf
DoEvents
Refresh
UpdateDevice
frmCM11A.Timer1.Enabled = True
End Sub

Private Sub X10CM11_X10SingleEvent(Devices As String, HouseCode As String, Command As Integer, Extra As String, Data2 As String)
txtEvent.Text = txtEvent.Text & "Single Event: " & HouseCode & " " & Devices & vbCrLf  '& " " & Str(Command) & " Ex = " & extra & " data2= " & data2 & vbCrLf
frmCM11A.Timer1.Enabled = True
End Sub
