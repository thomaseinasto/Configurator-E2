VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evikon Series Configurator"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21735
   Icon            =   "E005.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1072.359
   ScaleMode       =   0  'User
   ScaleWidth      =   1448.748
   Begin VB.Frame Homepage 
      BackColor       =   &H8000000E&
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   9735
      Left            =   240
      TabIndex        =   234
      Top             =   720
      Width           =   15855
      Begin VB.CommandButton PVT10 
         Caption         =   "PVT10 Configurator"
         Height          =   735
         Left            =   6240
         TabIndex        =   241
         Top             =   6000
         Width           =   3375
         Visible         =   0   'False
      End
      Begin VB.CommandButton PVT100 
         Caption         =   "PVT100 Configurator"
         Height          =   735
         Left            =   6240
         TabIndex        =   240
         Top             =   4920
         Width           =   3375
         Visible         =   0   'False
      End
      Begin VB.CommandButton Custom 
         Caption         =   "Custom Registers"
         Height          =   735
         Left            =   6240
         TabIndex        =   237
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CommandButton E22XX_Conf 
         Caption         =   "E22XX Configurator"
         Height          =   735
         Left            =   6240
         TabIndex        =   236
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton E26XX_Conf 
         Caption         =   "E26XX Configurator"
         Height          =   735
         Left            =   6240
         TabIndex        =   235
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label secret3 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   15120
         TabIndex        =   245
         Top             =   9240
         Width           =   615
      End
      Begin VB.Label secret2 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   244
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Secret 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   15120
         TabIndex        =   243
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   675
         Index           =   2
         Left            =   2880
         Picture         =   "E005.frx":058A
         Top             =   8760
         Width           =   2490
      End
      Begin VB.Image Image1 
         Height          =   780
         Index           =   2
         Left            =   120
         Picture         =   "E005.frx":5DB0
         Top             =   8760
         Width           =   2685
      End
   End
   Begin VB.CommandButton button_scan 
      Caption         =   "DEFAULT"
      Height          =   372
      Left            =   7920
      TabIndex        =   173
      ToolTipText     =   "Set Slave ID to 1 and Baudrate to 9600"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton button_custom 
      Caption         =   "CUSTOM"
      Height          =   372
      Left            =   10800
      TabIndex        =   156
      ToolTipText     =   "User defined modbus registers"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   16320
      Top             =   1200
   End
   Begin VB.Frame Frame_Custom 
      BackColor       =   &H8000000E&
      Caption         =   "Custom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   8655
      Left            =   240
      TabIndex        =   93
      Top             =   1200
      Width           =   15855
      Visible         =   0   'False
      Begin VB.CheckBox Check_c_neg6 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   154
         ToolTipText     =   "For signed integer values"
         Top             =   4800
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg7 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   153
         ToolTipText     =   "For signed integer values"
         Top             =   5520
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg8 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   152
         ToolTipText     =   "For signed integer values"
         Top             =   6240
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg1 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   151
         ToolTipText     =   "For signed integer values"
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg2 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   150
         ToolTipText     =   "For signed integer values"
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg3 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   149
         ToolTipText     =   "For signed integer values"
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg4 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   148
         ToolTipText     =   "For signed integer values"
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check_c_neg5 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   147
         ToolTipText     =   "For signed integer values"
         Top             =   4080
         Width           =   255
      End
      Begin VB.TextBox Text_c_write8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   146
         Top             =   6120
         Width           =   855
      End
      Begin VB.TextBox Text_c_write2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   145
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text_c_write3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   144
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text_c_write4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   143
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text_c_write5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   142
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text_c_write6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   141
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Text_c_write7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   140
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text_c_adr7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   132
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   131
         Top             =   6120
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   130
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   129
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   128
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   127
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text_c_adr6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   126
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox Text_name2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   122
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox Text_name6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   119
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox Text_name8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   118
         Top             =   6120
         Width           =   2775
      End
      Begin VB.TextBox Text_name5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   117
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox Text_name4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   116
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox Text_name3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   115
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text_name7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   114
         Top             =   5400
         Width           =   2775
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   113
         Top             =   6240
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   112
         Top             =   5520
         Width           =   255
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   111
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   110
         Top             =   4800
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   109
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   720
         TabIndex        =   108
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   107
         Top             =   1920
         Width           =   210
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   106
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox Text_c_write1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7680
         TabIndex        =   105
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text_c_adr1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   103
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text_name1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   102
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Image Image2 
         Height          =   675
         Index           =   1
         Left            =   3000
         Picture         =   "E005.frx":6FDF
         Top             =   7800
         Width           =   2490
      End
      Begin VB.Image Image1 
         Height          =   780
         Index           =   1
         Left            =   120
         Picture         =   "E005.frx":C805
         Top             =   7680
         Width           =   2685
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Signed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   155
         ToolTipText     =   "For signed integer values"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label_c_read7 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   139
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label_c_read8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   138
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label_c_read2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   137
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label_c_read3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   136
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label_c_read4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   135
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label_c_read5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   134
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label_c_read6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   133
         Top             =   4800
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Write"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   125
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Read"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   124
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   123
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000E&
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   121
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Register description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   120
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label_c_read1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   104
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.CommandButton button_save 
      Caption         =   "SAVE"
      Height          =   372
      Left            =   9360
      TabIndex        =   53
      ToolTipText     =   "Save all textbox values"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame_E222X 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   8655
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   15855
      Visible         =   0   'False
      Begin VB.TextBox Text_SW 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   242
         Text            =   "1"
         Top             =   720
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.CommandButton Home 
         Caption         =   "Home Page"
         Height          =   375
         Left            =   600
         TabIndex        =   238
         Top             =   6960
         Width           =   1815
      End
      Begin VB.ComboBox Combo_sensor_type 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DA34
         Left            =   5400
         List            =   "E005.frx":DA36
         Style           =   2  'Dropdown List
         TabIndex        =   228
         Top             =   5640
         Width           =   1935
         Visible         =   0   'False
      End
      Begin VB.TextBox CalibrationSETPOINT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   226
         Text            =   "0"
         Top             =   240
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Slope"
         Height          =   375
         Left            =   1320
         TabIndex        =   225
         Top             =   2640
         Width           =   1095
         Visible         =   0   'False
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Zero"
         Height          =   375
         Left            =   360
         TabIndex        =   222
         Top             =   2640
         Width           =   855
         Visible         =   0   'False
      End
      Begin VB.ComboBox Combo_gas_type 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DA38
         Left            =   5400
         List            =   "E005.frx":DA6F
         Style           =   2  'Dropdown List
         TabIndex        =   218
         Top             =   5160
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.ComboBox Combo_gas_units 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DAC4
         Left            =   5400
         List            =   "E005.frx":DAD1
         Style           =   2  'Dropdown List
         TabIndex        =   217
         Top             =   4680
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.ComboBox Combo_LED 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DAE0
         Left            =   5400
         List            =   "E005.frx":DAEA
         Style           =   2  'Dropdown List
         TabIndex        =   202
         ToolTipText     =   "Indicator light emitting diode"
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox Combo_buzzer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DAF7
         Left            =   5400
         List            =   "E005.frx":DB01
         Style           =   2  'Dropdown List
         TabIndex        =   199
         ToolTipText     =   "Acoustic alarm"
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text_sensor_pulse 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   194
         Text            =   "0"
         Top             =   6600
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_heater_pulse 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   193
         Text            =   "0"
         Top             =   6120
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_const_E 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   192
         Text            =   "100"
         Top             =   6300
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_const_D 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   191
         Text            =   "32000"
         Top             =   5880
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_const_C 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   190
         Text            =   "100"
         Top             =   5460
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_const_B 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   189
         Text            =   "100"
         Top             =   5040
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_factory 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   11
         TabIndex        =   174
         Text            =   "0"
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox Text_RC_filter 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   170
         Text            =   "0"
         ToolTipText     =   "Integrating filter time constant, 1...32000 (seconds), 0 - no filter"
         Top             =   2640
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_RH_rate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   167
         Text            =   "0"
         ToolTipText     =   $"E005.frx":DB0E
         Top             =   2160
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_RH_slope 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   164
         Text            =   "512"
         ToolTipText     =   "Slope adjustment for gas data, 1...65535"
         Top             =   1680
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_response 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   162
         Text            =   "10"
         ToolTipText     =   "Response delay, ms 10...255"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox Text_hardware 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   159
         Text            =   "2608"
         Top             =   360
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.ComboBox Combo_global_AN 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DB3D
         Left            =   5400
         List            =   "E005.frx":DB47
         Style           =   2  'Dropdown List
         TabIndex        =   90
         ToolTipText     =   "Enable/disable analog outputs"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text_zero_RH 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   6
         TabIndex        =   57
         Text            =   "0"
         ToolTipText     =   "Zero adjustment for gas data, ADC; -32000...+32000 ADC units"
         Top             =   1200
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_zero_T 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   54
         Text            =   "0"
         ToolTipText     =   "-320,00...+320,00 °C"
         Top             =   720
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         Caption         =   "Relay 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   3495
         Left            =   11640
         TabIndex        =   40
         Top             =   4800
         Width           =   3975
         Begin VB.CheckBox Relay2Save 
            BackColor       =   &H8000000E&
            Height          =   315
            Left            =   3600
            TabIndex        =   224
            Top             =   240
            Width           =   255
            Visible         =   0   'False
         End
         Begin VB.TextBox Text_RE2_time 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   4
            TabIndex        =   181
            Text            =   "0"
            ToolTipText     =   "Minimal on/off time 0...1000 s "
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox Text_RE2_delay 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   4
            TabIndex        =   180
            Text            =   "0"
            ToolTipText     =   "Switching delay 0...1000 s"
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox Text_RE2_H 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   7
            TabIndex        =   44
            Text            =   "85"
            ToolTipText     =   "Relay ""ON"" setpoint; -40,00...+85,00 °C; 0 - 32000 gas units"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text_RE2_L 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   7
            TabIndex        =   43
            Text            =   "0"
            ToolTipText     =   "Relay ""OFF"" setpoint; -40,00...+85,00 °C; 0 - 32000 gas units"
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox Combo_RE2_onoff 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "E005.frx":DB54
            Left            =   120
            List            =   "E005.frx":DB64
            Style           =   2  'Dropdown List
            TabIndex        =   42
            ToolTipText     =   "Parameter tied to relay"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox Combo_RE2_mode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "E005.frx":DB80
            Left            =   120
            List            =   "E005.frx":DB93
            Style           =   2  'Dropdown List
            TabIndex        =   41
            ToolTipText     =   "Relay control logic"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label_RE2_time 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   187
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label_RE2_delay 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   186
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   183
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Delay"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   182
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label_RE2_L 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   81
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label_RE2_H 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   80
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label_RE2_mode 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   79
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label_RE2_onoff 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   78
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "HIGH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "LOW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Caption         =   "Relay 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   3495
         Left            =   7560
         TabIndex        =   33
         Top             =   4800
         Width           =   3975
         Begin VB.CheckBox Relay1Save 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3600
            TabIndex        =   223
            Top             =   240
            Width           =   255
            Visible         =   0   'False
         End
         Begin VB.TextBox Text_RE1_time 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   4
            TabIndex        =   177
            Text            =   "0"
            ToolTipText     =   "Minimal on/off time 0...1000 s "
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox Text_RE1_delay 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   4
            TabIndex        =   176
            Text            =   "0"
            ToolTipText     =   "Switching delay 0...1000 s"
            Top             =   2280
            Width           =   975
         End
         Begin VB.ComboBox Combo_RE1_mode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DBD3
            Left            =   120
            List            =   "E005.frx":DBE6
            Style           =   2  'Dropdown List
            TabIndex        =   37
            ToolTipText     =   "Relay control logic"
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox Combo_RE1_onoff 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DC26
            Left            =   120
            List            =   "E005.frx":DC36
            Style           =   2  'Dropdown List
            TabIndex        =   36
            ToolTipText     =   "Parameter tied to relay"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text_RE1_L 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   7
            TabIndex        =   35
            Text            =   "0"
            ToolTipText     =   "Relay ""OFF"" setpoint; -40,00...+85,00 °C; 0 - 32000 gas units"
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text_RE1_H 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            MaxLength       =   7
            TabIndex        =   34
            Text            =   "85"
            ToolTipText     =   "Relay ""ON"" setpoint; -40,00...+85,00 °C; 0 - 32000 gas units "
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label_RE1_time 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   185
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label_RE1_delay 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   184
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   179
            Top             =   2820
            Width           =   612
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "Delay"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   178
            Top             =   2340
            Width           =   612
         End
         Begin VB.Label Label_RE1_L 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   77
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label_RE1_H 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   76
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label_RE1_mode 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   75
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label_RE1_onoff 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   74
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "LOW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   1860
            Width           =   612
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "HIGH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   38
            Top             =   1380
            Width           =   612
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         Caption         =   "Analog Output 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   4455
         Left            =   11640
         TabIndex        =   29
         Top             =   240
         Width           =   3975
         Begin VB.Frame Frame13 
            BackColor       =   &H8000000E&
            Caption         =   "°C / gas units"
            Height          =   735
            Left            =   120
            TabIndex        =   98
            ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
            Top             =   3000
            Width           =   2175
            Begin VB.TextBox Text_AN2_100deg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1200
               MaxLength       =   5
               TabIndex        =   100
               Text            =   "85"
               ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox Text_AN2_0deg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               MaxLength       =   5
               TabIndex        =   99
               Text            =   "0"
               ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   101
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H8000000E&
            Caption         =   "mA / V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   94
            ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
            Top             =   2160
            Width           =   2175
            Begin VB.TextBox Text_AN2_20ma 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1200
               MaxLength       =   2
               TabIndex        =   96
               Text            =   "10"
               ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox Text_AN2_4ma 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               MaxLength       =   2
               TabIndex        =   95
               Text            =   "4"
               ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   97
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H8000000E&
            Caption         =   "100% Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1320
            TabIndex        =   69
            ToolTipText     =   "100% value for analog output ( Modbus register data)"
            Top             =   1320
            Width           =   975
            Begin VB.Label Label_AN2_100_value 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   70
               ToolTipText     =   "100% value for analog output ( Modbus register data)"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H8000000E&
            Caption         =   "0% Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "0% value for analog output ( Modbus register data)"
            Top             =   1320
            Width           =   975
            Begin VB.Label Label_AN2_0_value 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   68
               ToolTipText     =   "0% value for analog output ( Modbus register data)"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.ComboBox Combo_AN2_onoff 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DC52
            Left            =   2520
            List            =   "E005.frx":DC62
            Style           =   2  'Dropdown List
            TabIndex        =   32
            ToolTipText     =   "Parameter tied to analog output"
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ComboBox Combo_AN2_I_U 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DC7E
            Left            =   2520
            List            =   "E005.frx":DC88
            Style           =   2  'Dropdown List
            TabIndex        =   31
            ToolTipText     =   "Analog output type"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.ComboBox Combo_AN2_diag 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "E005.frx":DC9E
            Left            =   240
            List            =   "E005.frx":DCAB
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Analog Output current, when sensor missing/damaged"
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Jumpers settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   233
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Current settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   232
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label_AN2_diag 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   73
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label_AN2_I_U 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   72
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label_AN2_onoff 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   71
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000005&
         Caption         =   "Analog Output 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   4455
         Left            =   7560
         TabIndex        =   25
         Top             =   240
         Width           =   3975
         Begin VB.Frame Frame11 
            BackColor       =   &H80000005&
            Caption         =   "°C / gas units"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   86
            ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
            Top             =   3000
            Width           =   2175
            Begin VB.TextBox Text_AN1_100deg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1200
               MaxLength       =   5
               TabIndex        =   88
               Text            =   "85"
               ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox Text_AN1_0deg 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               MaxLength       =   5
               TabIndex        =   87
               Text            =   "0"
               ToolTipText     =   "-40...+85 °C; 0 - 32000 gas units"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   89
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H80000005&
            Caption         =   "mA / V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
            Top             =   2160
            Width           =   2175
            Begin VB.TextBox Text_AN1_4ma 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   10.5
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               MaxLength       =   2
               TabIndex        =   84
               Text            =   "4"
               ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox Text_AN1_20ma 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   10.5
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1200
               MaxLength       =   2
               TabIndex        =   83
               Text            =   "10"
               ToolTipText     =   "Analog output scale: 4...20 mA ; 0...10 V"
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H8000000E&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   85
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H80000005&
            Caption         =   "100% Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1320
            TabIndex        =   64
            ToolTipText     =   "100% value for analog output ( Modbus register data)"
            Top             =   1320
            Width           =   975
            Begin VB.Label Label_AN1_100_value 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   65
               ToolTipText     =   "100% value for analog output ( Modbus register data)"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000005&
            Caption         =   "0% Out"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "0% value for analog output ( Modbus register data)"
            Top             =   1320
            Width           =   975
            Begin VB.Label Label_AN1_0_value 
               Alignment       =   2  'Center
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   186
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   66
               ToolTipText     =   "0% value for analog output ( Modbus register data)"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.ComboBox Combo_AN1_diag 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DCCC
            Left            =   240
            List            =   "E005.frx":DCD9
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Analog Output current, when sensor missing/damaged"
            Top             =   3960
            Width           =   1815
         End
         Begin VB.ComboBox Combo_AN1_I_U 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DCFA
            Left            =   2520
            List            =   "E005.frx":DD04
            Style           =   2  'Dropdown List
            TabIndex        =   27
            ToolTipText     =   "Analog output type"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.ComboBox Combo_AN1_onoff 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "E005.frx":DD1A
            Left            =   2520
            List            =   "E005.frx":DD2A
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Parameter tied to analog output"
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Current settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   231
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Jumpers settings"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   230
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label_AN1_diag 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   62
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label_AN1_I_U 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   61
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label_AN1_onoff 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   60
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox Text_SN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "1"
         Top             =   1080
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.TextBox Text_slave_id 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "1"
         ToolTipText     =   "net address, 1...247"
         Top             =   4380
         Width           =   975
      End
      Begin VB.ComboBox Combo_baud 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DD46
         Left            =   1440
         List            =   "E005.frx":DD5F
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox Combo_stop_bit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":DD90
         Left            =   1440
         List            =   "E005.frx":DD9A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3540
         Width           =   975
      End
      Begin VB.Label RH_units 
         BackColor       =   &H8000000E&
         Caption         =   "%RH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   239
         Top             =   2160
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   675
         Index           =   0
         Left            =   3000
         Picture         =   "E005.frx":DDA4
         Top             =   7800
         Width           =   2490
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Sensor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   229
         Top             =   5640
         Width           =   1095
         Visible         =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   780
         Index           =   0
         Left            =   120
         Picture         =   "E005.frx":135CA
         Top             =   7680
         Width           =   2685
      End
      Begin VB.Label Label51 
         BackColor       =   &H8000000E&
         Caption         =   "Calibration Gas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   227
         Top             =   360
         Width           =   1095
         Visible         =   0   'False
      End
      Begin VB.Label Label55 
         BackColor       =   &H8000000E&
         Caption         =   "Concentration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   221
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   " - "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   220
         Top             =   1800
         Width           =   735
         Visible         =   0   'False
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   158
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Hardware"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   157
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ADC Units"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   48
         Top             =   1800
         Width           =   1395
         Visible         =   0   'False
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   47
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Serial No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label_gas_units 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   216
         Top             =   4680
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label_gas_type 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   215
         Top             =   5160
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Concentration units"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3420
         TabIndex        =   214
         Top             =   4740
         Width           =   1935
         Visible         =   0   'False
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Gas type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   213
         Top             =   5220
         Width           =   1332
         Visible         =   0   'False
      End
      Begin VB.Label Label_sensor_pulse 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   212
         Top             =   6600
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label_heater_pulse 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   211
         Top             =   6120
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label_const_C 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   210
         Top             =   5460
         Width           =   972
         Visible         =   0   'False
      End
      Begin VB.Label Label_const_B 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2460
         TabIndex        =   209
         Top             =   5040
         Width           =   912
         Visible         =   0   'False
      End
      Begin VB.Label Label_const_E 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   208
         Top             =   6300
         Width           =   972
         Visible         =   0   'False
      End
      Begin VB.Label Label_const_D 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   207
         Top             =   5880
         Width           =   972
         Visible         =   0   'False
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Heater pulse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   206
         Top             =   6120
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Sensor pulse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   205
         Top             =   6600
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "LED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   204
         Top             =   4260
         Width           =   1212
      End
      Begin VB.Label Label_LED 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   203
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Buzzer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   201
         Top             =   3780
         Width           =   1212
      End
      Begin VB.Label Label_buzzer 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   200
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "const C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   198
         Top             =   5580
         Width           =   1092
         Visible         =   0   'False
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "const B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   197
         Top             =   5160
         Width           =   1212
         Visible         =   0   'False
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "const E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   360
         TabIndex        =   196
         Top             =   6420
         Width           =   972
         Visible         =   0   'False
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "const D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   195
         Top             =   6000
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Delay, ms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4200
         TabIndex        =   175
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   172
         Top             =   2640
         Width           =   495
         Visible         =   0   'False
      End
      Begin VB.Label Label_RC_filter 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   171
         Top             =   2640
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Gas rate ADJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   169
         Top             =   2160
         Width           =   1455
         Visible         =   0   'False
      End
      Begin VB.Label Label_RH_rate 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   168
         Top             =   2160
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Gas slope ADJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3600
         TabIndex        =   166
         Top             =   1680
         Width           =   1575
         Visible         =   0   'False
      End
      Begin VB.Label Label_RH_slope 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   165
         Top             =   1680
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label_response 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   163
         Top             =   3960
         Width           =   972
      End
      Begin VB.Label Label_hardware 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   161
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label_software 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   160
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label_global_AN 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   92
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Analog OUT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   91
         Top             =   3300
         Width           =   1212
      End
      Begin VB.Label Label_zero_RH 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   59
         Top             =   1200
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "GAS zero ADJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   58
         Top             =   1200
         Width           =   1215
         Visible         =   0   'False
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Temp zero ADJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3840
         TabIndex        =   56
         Top             =   720
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.Label Label_zero_T 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   55
         Top             =   720
         Width           =   975
         Visible         =   0   'False
      End
      Begin VB.Label Label_meas_gas_units 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   52
         Top             =   2160
         Width           =   615
         Visible         =   0   'False
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "°C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   51
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label_hum 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   50
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label_temp 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   49
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Slave ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   360
         TabIndex        =   23
         Top             =   4440
         Width           =   972
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Baudrate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   3180
         Width           =   1092
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Stop bits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   21
         Top             =   3600
         Width           =   972
      End
      Begin VB.Label Label_SN 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label_slave_id 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   19
         Top             =   4356
         Width           =   732
      End
      Begin VB.Label Label_baud 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2580
         TabIndex        =   18
         Top             =   3120
         Width           =   792
      End
      Begin VB.Label Label_stop_bit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   17
         Top             =   3540
         Width           =   972
      End
   End
   Begin VB.CommandButton button_write 
      Caption         =   "WRITE"
      Height          =   372
      Left            =   6480
      TabIndex        =   7
      ToolTipText     =   "Write all settings into device"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton button_read 
      Caption         =   "READ"
      Height          =   372
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Read all settings from device"
      Top             =   480
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   16320
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Frame Frame4_status 
      BackColor       =   &H8000000E&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   792
      Left            =   12120
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      Begin VB.Label Label_status 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "- - -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   186
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   915
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Communication settings"
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo_port_nr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":147F9
         Left            =   360
         List            =   "E005.frx":1482D
         Style           =   2  'Dropdown List
         TabIndex        =   219
         ToolTipText     =   "Communication settings"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Combo_stop_bit0 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":14868
         Left            =   2520
         List            =   "E005.frx":14872
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Communication settings"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox Combo_baud0 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "E005.frx":1487C
         Left            =   1200
         List            =   "E005.frx":14895
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Communication settings"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text_slave_id0 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1061
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "1"
         ToolTipText     =   "Communication settings"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Stop bits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         ToolTipText     =   "Communication settings"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Slave ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         ToolTipText     =   "Communication settings"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Baudrate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Communication settings"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Communication settings"
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public slave_id_global As Byte
Public frame_delay As Boolean
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long) ' Sleep "ms"

'GAS TYPE************************************************************************************
Private Sub Combo_gas_type_Click()
If Combo_gas_type.ListIndex = 0 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("TGS2610"), 0
    Combo_sensor_type.AddItem ("TGS2611"), 1
    Combo_sensor_type.AddItem ("TGS2612"), 2
ElseIf Combo_gas_type.ListIndex = 1 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("S+4 2ECO"), 0
    Combo_sensor_type.AddItem ("GS+4COSLI-M"), 1
    Combo_sensor_type.AddItem ("GS+4CO"), 2
    Combo_sensor_type.AddItem ("GS+4COHC"), 3
ElseIf Combo_gas_type.ListIndex = 2 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("S+4OX"), 0
    Combo_sensor_type.AddItem ("O2-A2"), 1
    Combo_sensor_type.AddItem ("O2-A3"), 2
ElseIf Combo_gas_type.ListIndex = 3 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4NH3-100"), 0
    Combo_sensor_type.AddItem ("GS+4NH3-300"), 1
    Combo_sensor_type.AddItem ("TGS2444"), 2
ElseIf Combo_gas_type.ListIndex = 4 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("TGS2610"), 0
    Combo_sensor_type.AddItem ("TGS2611"), 1
ElseIf Combo_gas_type.ListIndex = 5 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("TGS2602"), 0
    Combo_sensor_type.AddItem ("TGS2620"), 1
ElseIf Combo_gas_type.ListIndex = 6 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("TGS2610"), 0
    Combo_sensor_type.AddItem ("GTGS2612"), 1
ElseIf Combo_gas_type.ListIndex = 7 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("TGS832"), 0
ElseIf Combo_gas_type.ListIndex = 8 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("O3/M-5"), 0
    Combo_sensor_type.AddItem ("O3/M-100"), 1
ElseIf Combo_gas_type.ListIndex = 9 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4H2S"), 0
    Combo_sensor_type.AddItem ("H2S-A1"), 1
ElseIf Combo_gas_type.ListIndex = 10 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("HCl-A1"), 0
ElseIf Combo_gas_type.ListIndex = 11 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4CL2"), 0
    Combo_sensor_type.AddItem ("CL2-A1"), 1
ElseIf Combo_gas_type.ListIndex = 12 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4SO2"), 0
    Combo_sensor_type.AddItem ("SO2-AF"), 1
ElseIf Combo_gas_type.ListIndex = 13 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("C2H4/M-10"), 0
    Combo_sensor_type.AddItem ("C2H4/M-200"), 1
ElseIf Combo_gas_type.ListIndex = 14 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4ETO"), 0
    Combo_sensor_type.AddItem ("ETO-A1"), 1
ElseIf Combo_gas_type.ListIndex = 15 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4NO"), 0
    Combo_sensor_type.AddItem ("NO-A1"), 1
    Combo_sensor_type.AddItem ("NO-AE"), 2
ElseIf Combo_gas_type.ListIndex = 16 Then
Combo_sensor_type.Clear
    Combo_sensor_type.AddItem ("GS+4NO2"), 0
    Combo_sensor_type.AddItem ("S+42NO2"), 1
    Combo_sensor_type.AddItem ("NO2-A1"), 2
    Combo_sensor_type.AddItem ("NO2-AE"), 3
End If
End Sub

Private Sub Combo_sensor_type_Click()
'CH4 Settings
If Combo_gas_type.ListIndex = 0 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'CO Settings
ElseIf Combo_gas_type.ListIndex = 1 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "42"
   Text_const_C.Text = "79"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "46"
   Text_const_C.Text = "85"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "48"
   Text_const_C.Text = "76"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 3 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "45"
   Text_const_C.Text = "76"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'O2 Settings
ElseIf Combo_gas_type.ListIndex = 2 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "93"
   Text_const_C.Text = "120"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "150"
   Text_const_C.Text = "240"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "150"
   Text_const_C.Text = "260"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
End If
'NH3 Settings
ElseIf Combo_gas_type.ListIndex = 3 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "47"
   Text_const_C.Text = "50"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "47"
   Text_const_C.Text = "50"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "14"
   Text_sensor_pulse.Text = "5"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'H2 Settings
ElseIf Combo_gas_type.ListIndex = 4 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'VOC Settings
ElseIf Combo_gas_type.ListIndex = 5 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'LPG Settings
ElseIf Combo_gas_type.ListIndex = 6 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'HFC Settings
ElseIf Combo_gas_type.ListIndex = 7 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "1000"
   Text_sensor_pulse.Text = "1000"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'O3 Settings
ElseIf Combo_gas_type.ListIndex = 8 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
End If
'H2S Settings
ElseIf Combo_gas_type.ListIndex = 9 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "73"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "116"
   Text_const_C.Text = "200"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'HCL Settings
ElseIf Combo_gas_type.ListIndex = 10 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "92"
   Text_const_C.Text = "200"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'CL2 Settings
ElseIf Combo_gas_type.ListIndex = 11 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
End If
'SO2 Settings
ElseIf Combo_gas_type.ListIndex = 12 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "48"
   Text_const_C.Text = "54"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'C2H4 Settings
ElseIf Combo_gas_type.ListIndex = 13 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'ETO Settings
ElseIf Combo_gas_type.ListIndex = 14 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "30"
   Text_const_C.Text = "500"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "46"
   Text_const_C.Text = "130"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'NO Settings
ElseIf Combo_gas_type.ListIndex = 15 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "59"
   Text_const_C.Text = "70"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "46"
   Text_const_C.Text = "65"
   Text_const_D.Text = "32000"
   Text_const_E.Text = "0"
End If
'NO2 Settings
ElseIf Combo_gas_type.ListIndex = 16 Then
If Combo_sensor_type.ListIndex = 0 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "73"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 1 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "100"
   Text_const_C.Text = "100"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 2 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "118"
   Text_const_C.Text = "180"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
ElseIf Combo_sensor_type.ListIndex = 3 Then
   Text_heater_pulse.Text = "0"
   Text_sensor_pulse.Text = "1"
   Text_const_B.Text = "85"
   Text_const_C.Text = "120"
   Text_const_D.Text = "-32000"
   Text_const_E.Text = "0"
End If
End If
End Sub
Private Sub Command2_Click()
Dim smd_d As Long
Dim smd_c As Long
Dim smd_b As Long
If Form1.Frame_E222X.Caption = "E26XX" Then
 If Label_gas_type = "O2" Then
   smd_d = (Label30)
 Else
   smd_d = -(Label30)
 End If
   '(20 - Text_zero_RH.Text) / Text_RH_slope
   smd_c = 2
   '(1410 - Text_zero_RH.Text) / (Text1 * (20 - Text_zero_RH.Text))
   Text_zero_RH = smd_d
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
Text_zero_T = Text_const_B.Text / (32000 * (Label_const_C.Caption - Label_const_D.Caption))
Call button_write_Click
Call button_read_Click
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
Text_zero_T = Text_const_B.Text / (32000 * (Label_const_C.Caption - Label_const_D.Caption))
Call button_write_Click
Call button_read_Click
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
Text_zero_T = Text_const_B.Text / (32000 * (Label_const_C.Caption - Label_const_D.Caption))
Call button_write_Click
Call button_read_Click
End If
End Sub

Private Sub Command3_Click()
Dim smd_b As Long
If Label_hum = "-" Then
MsgBox "You have to read the device first!", 48, "Calibration error"
End If
If CalibrationSETPOINT = 0 Then
MsgBox "Calibration setpoint is 0!", 48, "Invalid setpoint"
End If
If Label_hum = "0" Then
MsgBox "You cannot calibrate slope in clean air!", 48, "Calibration Error"
End If
If Label_hum > "0" Then
If CalibrationSETPOINT > 0 Then
   smd_b = Label_RH_slope * CalibrationSETPOINT / Label_hum
   Text_RH_slope = smd_b
End If
End If
End Sub

Private Sub Command5_Click()
If button_custom.Caption = "E26XX" Then
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   button_custom.Caption = "CUSTOM"
ElseIf button_custom.Caption = "HOME" Then
   Frame_Custom.Visible = False
   button_custom.Caption = "E26XX"
ElseIf button_custom.Caption = "CUSTOM" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   button_custom.Caption = "E26XX"
End If
End Sub
Private Sub Custom_Click()
   button_custom.Caption = "HOME"
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   Homepage.Visible = False
End Sub

Private Sub E22XX_Conf_Click()
   Frame_E222X.Caption = "E22XX"
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   Homepage.Visible = False
   Label44.Caption = "Temp Rate"
   Label41.Caption = "Temp Slope"
   Label13.Caption = "RH Zero"
   Label40.Caption = "RH Slope"
   Label43.Caption = "RH Rate"
   Label24.Caption = "Dew Point"
   Label55.Caption = "         Humidity"
   Text_hardware.Text = "2208"
   Text_SN.Text = "1"
   Text_RH_slope.Text = "0"
   Text_hardware.Visible = True
   Frame11.Caption = "°C / %RH"
   Frame13.Caption = "°C / %RH"
   RH_units.Visible = True
   Form1.Text_SN.Text = "1"
   Form1.Text_SW.Text = "502"
   Text_SW.Visible = False
   'K
   Text_const_B.Visible = True
   Label_const_B.Visible = True
   Label34.Visible = True
   Label34 = "K"
   Text_const_B.Text = "7000"
   'Tm
   Label_const_C.Visible = True
   Label35.Visible = True
   Label35 = "Tm"
   'Ts
   Label_const_D.Visible = True
   Label32.Visible = True
   Label32 = "Ts"
   Command2.Visible = True
   Command2.Caption = "SHC"
   Text_factory.Text = "0"
   Form1.slave_id_global = "1"
   Text_slave_id.Text = Val("1")
   Text_slave_id0 = Val("1")
   Combo_AN1_onoff.Clear
   Combo_AN1_onoff.AddItem ("OFF"), 0
   Combo_AN1_onoff.AddItem ("Temp"), 1
   Combo_AN1_onoff.AddItem ("Hum"), 2
   Combo_AN1_onoff.AddItem ("ModBus"), 3
   Combo_AN2_onoff.Clear
   Combo_AN2_onoff.AddItem ("OFF"), 0
   Combo_AN2_onoff.AddItem ("Temp"), 1
   Combo_AN2_onoff.AddItem ("Hum"), 2
   Combo_AN2_onoff.AddItem ("ModBus"), 3
   Combo_AN1_onoff.ListIndex = 2
   Combo_AN2_onoff.ListIndex = 1
   Combo_AN1_diag.ListIndex = 1
   Combo_AN2_diag.ListIndex = 1
   Combo_global_AN.ListIndex = 0
Call port_scan
End Sub

Private Sub E26XX_Conf_Click()
   Frame_E222X.Caption = "E26XX"
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   Homepage.Visible = False
   Text_SN.Text = "1"
   Label44.Caption = "Heater Pulse"
   Label41.Caption = "Sensor Pulse"
   Label13.Caption = "Gas Zero"
   Label40.Caption = "Gas Slope ADJ"
   Label43.Caption = "Gas Rate ADJ"
   Label24.Caption = "ADC Units"
   Label55.Caption = "Consentration"
   Text_hardware.Text = "2608"
   Text_RH_slope.Text = "502"
   Text_hardware.Visible = True
   Frame11.Caption = "°C / gas units"
   Frame13.Caption = "°C / gas units"
   RH_units.Visible = False
   Text_factory.Text = "0"
   Text_slave_id = Val("1")
   Text_slave_id0 = Val("1")
   Form1.slave_id_global = "1"
   Command2.Caption = "Zero"
   Combo_AN1_onoff.Clear
   Combo_AN1_onoff.AddItem ("OFF"), 0
   Combo_AN1_onoff.AddItem ("Temp"), 1
   Combo_AN1_onoff.AddItem ("GAS"), 2
   Combo_AN1_onoff.AddItem ("ModBus"), 3
   Combo_AN2_onoff.Clear
   Combo_AN2_onoff.AddItem ("OFF"), 0
   Combo_AN2_onoff.AddItem ("Temp"), 1
   Combo_AN2_onoff.AddItem ("GAS"), 2
   Combo_AN2_onoff.AddItem ("ModBus"), 3
   Combo_AN1_onoff.ListIndex = 2
   Combo_AN2_onoff.ListIndex = 1
   Combo_AN1_diag.ListIndex = 1
   Combo_AN2_diag.ListIndex = 1
   Combo_global_AN.ListIndex = 0
   Form1.Text_SN.Text = "1"
   Text_SW.Visible = False
   Label34 = "Const B"
   Label32 = "Const D"
   Label35 = "Const C"
Call port_scan
End Sub

Private Sub Form_Load()
Call buttons_disabled
Timer1.Enabled = False
Timer1.Interval = 2500
MSComm1.RThreshold = 0 ' events disabled

Frame_E222X.Visible = False
Frame_Custom.Visible = False

'***************************************************************************************************
'if DAT file missing, fill combo boxes
Combo_port_nr.ListIndex = 0
Combo_baud0.ListIndex = 3
Combo_stop_bit0.ListIndex = 0
Combo_baud.ListIndex = 3
Combo_stop_bit.ListIndex = 0
Combo_global_AN.ListIndex = 1
Combo_AN1_onoff.ListIndex = 0
Combo_AN2_onoff.ListIndex = 0
Combo_AN1_I_U.ListIndex = 0
Combo_AN2_I_U.ListIndex = 0
Combo_AN1_diag.ListIndex = 2
Combo_AN2_diag.ListIndex = 2
Combo_RE1_onoff.ListIndex = 0
Combo_RE2_onoff.ListIndex = 0
Combo_RE1_mode.ListIndex = 0
Combo_RE2_mode.ListIndex = 0
Combo_buzzer.ListIndex = 1
Combo_LED.ListIndex = 1
Combo_gas_units.ListIndex = 0
Combo_gas_type.ListIndex = 0
'***************************************************************************************************
Call read_from_txt ' READ SETTINGS from txt


slave_id_global = Val(Text_slave_id0.Text) ' SET GLOBAL SLAVE ID

Call port_scan
Timer1.Enabled = True

End Sub
Private Sub button_save_Click() ' SAVE SETTINGS to txt
Call write_to_txt
End Sub


Private Sub Home_Click()
   Frame_E222X.Visible = False
   Frame_Custom.Visible = False
   Homepage.Visible = True
   Command3.Visible = False
Label51.Visible = False
Command2.Visible = False
CalibrationSETPOINT.Visible = False
Command2.Visible = False
Text_hardware.Visible = False
Text_SN.Visible = False
Text_const_B.Visible = False
Text_const_C.Visible = False
Text_const_D.Visible = False
Text_const_E.Visible = False
Text_heater_pulse.Visible = False
Text_sensor_pulse.Visible = False
Label_const_B.Visible = False
Label_const_C.Visible = False
Label_const_D.Visible = False
Label_const_E.Visible = False
Label_heater_pulse.Visible = False
Label30.Visible = False
Label_sensor_pulse.Visible = False
Label34.Visible = False
Label35.Visible = False
Label32.Visible = False
Label33.Visible = False
Label44.Visible = False
Label41.Visible = False
Label36.Visible = False
Label38.Visible = False
Combo_gas_units.Visible = False
Label_gas_units.Visible = False
Combo_gas_type.Visible = False
Combo_sensor_type.Visible = False
Label52.Visible = False
Label_gas_type.Visible = False
Label12.Visible = False
Label30.Visible = False
Label24.Visible = False
Label13.Visible = False
Label40.Visible = False
Label43.Visible = False
Label45.Visible = False
Text_zero_T.Visible = False
Text_zero_RH.Visible = False
Text_RH_rate.Visible = False
Text_RH_slope.Visible = False
Text_RC_filter.Visible = False
Label_zero_T.Visible = False
Label_zero_RH.Visible = False
Label_RH_rate.Visible = False
Label_RH_slope.Visible = False
Label_RC_filter.Visible = False
End Sub

Private Sub PVT10_Click()
   Frame_E222X.Caption = "PVT10"
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   Homepage.Visible = False
   Label44.Caption = "Temp Rate"
   Label41.Caption = "Temp Slope"
   Label13.Caption = "RH Zero"
   Label40.Caption = "RH Slope"
   Label43.Caption = "RH Rate"
   Label24.Caption = "Dew Point"
   Label55.Caption = "         Humidity"
   Text_hardware.Text = "20566"
   Text_SW.Text = "21553"
   Text_SN.Text = "12288"
   Text_SN.Visible = True
   Text_SW.Visible = True
   Text_RH_slope.Text = "0"
   Text_hardware.Visible = True
   Frame11.Caption = "°C / %RH"
   Frame13.Caption = "°C / %RH"
   Text_slave_id0.Text = "16"
   RH_units.Visible = True
   'K
   Text_const_B.Visible = True
   Label_const_B.Visible = True
   Label34.Visible = True
   Label34 = "K"
   Text_const_B.Text = "7000"
   'Tm
   Label_const_C.Visible = True
   Label35.Visible = True
   Label35 = "Tm"
   'Ts
   Label_const_D.Visible = True
   Label32.Visible = True
   Label32 = "Ts"
   Command2.Visible = True
   Command2.Caption = "SHC"
   Text_slave_id.Text = Val("16")
   Form1.slave_id_global = "16"
   Text_factory.Text = "0"
   Combo_AN1_onoff.Clear
   Combo_AN1_onoff.AddItem ("OFF"), 0
   Combo_AN1_onoff.AddItem ("Temp"), 1
   Combo_AN1_onoff.AddItem ("Hum"), 2
   Combo_AN1_onoff.AddItem ("ModBus"), 3
   Combo_AN2_onoff.Clear
   Combo_AN2_onoff.AddItem ("OFF"), 0
   Combo_AN2_onoff.AddItem ("Temp"), 1
   Combo_AN2_onoff.AddItem ("Hum"), 2
   Combo_AN2_onoff.AddItem ("ModBus"), 3
   Combo_AN1_onoff.ListIndex = 2
   Combo_AN2_onoff.ListIndex = 1
   Combo_AN1_diag.ListIndex = 1
   Combo_AN2_diag.ListIndex = 1
   Combo_global_AN.ListIndex = 0
Call port_scan
End Sub

Private Sub PVT100_Click()
   Frame_E222X.Caption = "PVT100"
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   Homepage.Visible = False
   Label44.Caption = "Temp Rate"
   Label41.Caption = "Temp Slope"
   Label13.Caption = "RH Zero"
   Label40.Caption = "RH Slope"
   Label43.Caption = "RH Rate"
   Label24.Caption = "Dew Point"
   Label55.Caption = "         Humidity"
   Text_hardware.Text = "20566"
   Text_SW.Text = "21553"
   Text_SN.Text = "12336"
   Text_SN.Visible = True
   Text_SW.Visible = True
   Text_RH_slope.Text = "0"
   Text_hardware.Visible = True
   Frame11.Caption = "°C / %RH"
   Frame13.Caption = "°C / %RH"
   Text_slave_id0.Text = "16"
   RH_units.Visible = True
   'K
   Text_const_B.Visible = True
   Label_const_B.Visible = True
   Label34.Visible = True
   Label34 = "K"
   Text_const_B.Text = "7000"
   'Tm
   Label_const_C.Visible = True
   Label35.Visible = True
   Label35 = "Tm"
   'Ts
   Label_const_D.Visible = True
   Label32.Visible = True
   Label32 = "Ts"
   Command2.Visible = True
   Command2.Caption = "SHC"
   Form1.slave_id_global = "16"
   Text_slave_id = Val("16")
   Text_factory.Text = "0"
   Combo_AN1_onoff.Clear
   Combo_AN1_onoff.AddItem ("OFF"), 0
   Combo_AN1_onoff.AddItem ("Temp"), 1
   Combo_AN1_onoff.AddItem ("Hum"), 2
   Combo_AN1_onoff.AddItem ("ModBus"), 3
   Combo_AN2_onoff.Clear
   Combo_AN2_onoff.AddItem ("OFF"), 0
   Combo_AN2_onoff.AddItem ("Temp"), 1
   Combo_AN2_onoff.AddItem ("Hum"), 2
   Combo_AN2_onoff.AddItem ("ModBus"), 3
   Combo_AN1_onoff.ListIndex = 2
   Combo_AN2_onoff.ListIndex = 1
   Combo_AN1_diag.ListIndex = 1
   Combo_AN2_diag.ListIndex = 1
   Combo_global_AN.ListIndex = 0
Call port_scan
End Sub

Private Sub Relay1Save_Click()
If Form1.Relay1Save.Value = 1 Then
   Text_RE1_L.Text = Label_RE1_L
   Text_RE1_H.Text = Label_RE1_H
   End If
End Sub

Private Sub Secret_Click()
Secret.Caption = "-"
End Sub

Private Sub secret2_Click()
secret2.Caption = "-"
End Sub

Private Sub secret3_Click()
secret3.Caption = "-"
If Secret.Caption = "-" And secret2.Caption = "-" And secret3.Caption = "-" Then
PVT100.Visible = True
PVT10.Visible = True
End If
End Sub

Private Sub Text_factory_Change() ' PASSWORD CHECK
If Form1.Frame_E222X.Caption = "E26XX" Then
If Text_factory.Text = "0xA55A" Then
Text_hardware.Visible = True
Text_SN.Visible = True
Label30.Visible = True
Label38.Visible = True
Label51.Visible = True
Command2.Visible = True
Command3.Visible = True
CalibrationSETPOINT.Visible = True
Combo_gas_units.Visible = True
Label_gas_units.Visible = True
Label12.Visible = True
Label30.Visible = True
Label24.Visible = True
Label13.Visible = True
Label40.Visible = True
Label43.Visible = True
Label45.Visible = True
Text_zero_T.Visible = True
Text_zero_RH.Visible = True
Text_RH_rate.Visible = True
Text_RH_slope.Visible = True
Text_RC_filter.Visible = True
Label_zero_T.Visible = True
Label_zero_RH.Visible = True
Label_RH_rate.Visible = True
Label_RH_slope.Visible = True
Label_RC_filter.Visible = True
Command2.Visible = True
End If

If Text_factory.Text = "0xA3v1k0nA" Then
Command3.Visible = True
Label51.Visible = True
Command2.Visible = True
CalibrationSETPOINT.Visible = True
Command2.Visible = True
Text_hardware.Visible = True
Text_SN.Visible = True
Text_const_B.Visible = True
Text_const_C.Visible = True
Text_const_D.Visible = True
Text_const_E.Visible = True
Text_heater_pulse.Visible = True
Text_sensor_pulse.Visible = True
Label_const_B.Visible = True
Label_const_C.Visible = True
Label_const_D.Visible = True
Label_const_E.Visible = True
Label_heater_pulse.Visible = True
Label30.Visible = True
Label_sensor_pulse.Visible = True
Label34.Visible = True
Label35.Visible = True
Label32.Visible = True
Label33.Visible = True
Label44.Visible = True
Label41.Visible = True
Label36.Visible = True
Label38.Visible = True
Combo_gas_units.Visible = True
Label_gas_units.Visible = True
Combo_gas_type.Visible = True
Combo_sensor_type.Visible = True
Label52.Visible = True
Label_gas_type.Visible = True
Label12.Visible = True
Label30.Visible = True
Label24.Visible = True
Label13.Visible = True
Label40.Visible = True
Label43.Visible = True
Label45.Visible = True
Text_zero_T.Visible = True
Text_zero_RH.Visible = True
Text_RH_rate.Visible = True
Text_RH_slope.Visible = True
Text_RC_filter.Visible = True
Label_zero_T.Visible = True
Label_zero_RH.Visible = True
Label_RH_rate.Visible = True
Label_RH_slope.Visible = True
Label_RC_filter.Visible = True
End If
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
If Text_factory.Text = "0xA55A" Then
Text_hardware.Visible = True
Text_SN.Visible = True
Label12.Visible = True
Label37.Visible = True
Label13.Visible = True
Label40.Visible = True
Label43.Visible = True
Label44.Visible = True
Label41.Visible = True
Label45.Visible = True
Text_zero_T.Visible = True
Text_heater_pulse.Visible = True
Text_sensor_pulse.Visible = True
Text_zero_RH.Visible = True
Text_RH_rate.Visible = True
Text_RH_slope.Visible = True
Text_RC_filter.Visible = True
Label_zero_T.Visible = True
Label_sensor_pulse.Visible = True
Label_heater_pulse.Visible = True
Label_zero_RH.Visible = True
Label_RH_rate.Visible = True
Label_RH_slope.Visible = True
Label_RC_filter.Visible = True
Label24.Visible = True
Label30.Visible = True
End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
If Text_factory.Text = "0xA55A" Then
Text_hardware.Visible = True
Text_SN.Visible = True
Label12.Visible = True
Label37.Visible = True
Label13.Visible = True
Label40.Visible = True
Label43.Visible = True
Label44.Visible = True
Label41.Visible = True
Label45.Visible = True
Text_zero_T.Visible = True
Text_heater_pulse.Visible = True
Text_sensor_pulse.Visible = True
Text_zero_RH.Visible = True
Text_RH_rate.Visible = True
Text_RH_slope.Visible = True
Text_RC_filter.Visible = True
Label_zero_T.Visible = True
Label_sensor_pulse.Visible = True
Label_heater_pulse.Visible = True
Label_zero_RH.Visible = True
Label_RH_rate.Visible = True
Label_RH_slope.Visible = True
Label_RC_filter.Visible = True
Label24.Visible = True
Label30.Visible = True
End If
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
If Text_factory.Text = "0xA55A" Then
Text_hardware.Visible = True
Text_SN.Visible = True
Label12.Visible = True
Label37.Visible = True
Label13.Visible = True
Label40.Visible = True
Label43.Visible = True
Label44.Visible = True
Label41.Visible = True
Label45.Visible = True
Text_zero_T.Visible = True
Text_heater_pulse.Visible = True
Text_sensor_pulse.Visible = True
Text_zero_RH.Visible = True
Text_RH_rate.Visible = True
Text_RH_slope.Visible = True
Text_RC_filter.Visible = True
Label_zero_T.Visible = True
Label_sensor_pulse.Visible = True
Label_heater_pulse.Visible = True
Label_zero_RH.Visible = True
Label_RH_rate.Visible = True
Label_RH_slope.Visible = True
Label_RC_filter.Visible = True
Label24.Visible = True
Label30.Visible = True
End If
End If

End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = False
Timer1.Interval = 500
button_write.Enabled = True
button_read.Enabled = True
button_save.Enabled = True
button_custom.Enabled = True
button_scan.Enabled = True
End Sub
Private Sub buttons_disabled()
button_write.Enabled = False
button_read.Enabled = False
button_save.Enabled = False
button_custom.Enabled = False
button_scan.Enabled = False
End Sub
Private Sub button_scan_Click()
Call buttons_disabled
Call default_settings
End Sub
Private Sub default_settings() ' SET default: baud=9600, slaveID=1
Dim count_1 As Byte
Dim count_2 As Byte
Dim output_string As String
slave_id_global = 0
'TÄIENDAV OTSING, proovi erinevaid porte, erinevaid kiirusi:
For count_1 = 1 To 16

  If port_connected(count_1) = True Then ' Try to open port
      
      For count_2 = 0 To 6
      
         Combo_baud0.ListIndex = count_2
         
         output_string = write_no_respond(4, 1)
         
         output_string = write_no_respond(5, 9600)
         
         output_string = write_no_respond(17, 42330)
         
         output_string = write_no_respond(17, 42330)
    
      Next
   End If
Next

If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
End If
If Form1.Frame_E222X.Caption = "PVT100" Then
slave_id_global = 16
Text_slave_id0.Text = "16"
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
slave_id_global = 16
Text_slave_id0.Text = "16"
Else
slave_id_global = 1
Text_slave_id0.Text = "1"
End If
Combo_baud0.ListIndex = 3
Timer1.Enabled = True
End Sub
Private Sub port_scan() ' SCANNING PORTS
Dim count_1 As Byte
Dim output_string As String

Label_status.Caption = "NO PORT"

For count_1 = 1 To 16

   If port_connected(count_1) = True Then
   
      count_1 = count_1 - 1
      Combo_port_nr.ListIndex = CStr(count_1)
      count_1 = count_1 + 1
   
      Label_status.Caption = "NO DEVICE"
      output_string = modbus_read(1)
      
      If IsNumeric(output_string) Then
        If Form1.Frame_E222X.Caption = "E26XX" Then
         If Mid$(output_string, 1, 4) = "2608" Then
            Label_status.Caption = "CONNECTED"
         Else
            Label_status.Caption = "WRONG DEVICE"
         End If
        End If
        If Form1.Frame_E222X.Caption = "PVT100" Then
         If Mid$(output_string, 1, 2) = "20" Then
            Label_status.Caption = "CONNECTED"
         Else
            Label_status.Caption = "WRONG DEVICE"
         End If
        End If
        If Form1.Frame_E222X.Caption = "PVT10" Then
         If Mid$(output_string, 1, 2) = "20" Then
            Label_status.Caption = "CONNECTED"
         Else
            Label_status.Caption = "WRONG DEVICE"
         End If
        End If
        If Form1.Frame_E222X.Caption = "E22XX" Then
         If Mid$(output_string, 1, 2) = "22" Then
            Label_status.Caption = "CONNECTED"
         Else
            Label_status.Caption = "WRONG DEVICE"
         End If
        End If
         Exit Sub
         
      Else
         Label_status.Caption = output_string
      End If
      
   End If

Next

If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
End If


End Sub
Private Sub button_write_Click() ' WRITE ALL********************************************************

Call buttons_disabled
If Label_status.Caption = "NO PORT" Or Label_status.Caption = "NO DEVICE" Then
   Call port_scan
End If
If port_connected(Val(Combo_port_nr.Text)) = True Then
Label_status.Caption = "NO DEVICE"
MSComm1.InputLen = 8                                   ' 8 CHAR from RX buffer

   If Frame_E222X.Visible = True Then
      Label_status.Caption = E222x_write_all()
   ElseIf Frame_Custom.Visible = True Then
      Label_status.Caption = Custom_write
   End If
Else
   Label_status.Caption = "NO PORT"
End If

If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
End If

Timer1.Enabled = True

End Sub
Private Sub button_read_Click() ' READ ALL**********************************************************

Call buttons_disabled
If Label_status.Caption = "NO PORT" Or Label_status.Caption = "NO DEVICE" Then
   Call port_scan
End If
If port_connected(Val(Combo_port_nr.Text)) = True Then
Label_status.Caption = "NO DEVICE"
MSComm1.InputLen = 7                                   ' 7 CHAR from RX buffer

   If Frame_E222X.Visible = True Then
      Label_status.Caption = E222x_read_all()
   ElseIf Frame_Custom.Visible = True Then
      Label_status.Caption = Custom_read()
   End If
Else
   Label_status.Caption = "NO PORT"
End If

If MSComm1.PortOpen = True Then
   MSComm1.PortOpen = False
End If

Timer1.Enabled = True

End Sub
Private Function port_connected(comport As Byte) As Boolean ' Pordi kontroll >>>
On Error GoTo cannot_open_port
'close port
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
MSComm1.CommPort = comport
'try to open port
MSComm1.PortOpen = True
port_connected = True ' com port connected
Exit Function
cannot_open_port:
port_connected = False ' com port disconnected
End Function
Private Sub Combo_baud0_Click() ' Select baudrate
   Call port_settings
End Sub
Private Sub port_settings()

If Combo_baud0.Text = "1200" Then
   MSComm1.Settings = "1200,N,8,2"
   frame_delay = True
ElseIf Combo_baud0.Text = "2400" Then
   MSComm1.Settings = "2400,N,8,2"
   frame_delay = True
ElseIf Combo_baud0.Text = "4800" Then
   MSComm1.Settings = "4800,N,8,2"
   frame_delay = True
ElseIf Combo_baud0.Text = "9600" Then
   MSComm1.Settings = "9600,N,8,2"
   frame_delay = False
ElseIf Combo_baud0.Text = "19200" Then
   MSComm1.Settings = "19200,N,8,2"
   frame_delay = False
ElseIf Combo_baud0.Text = "38400" Then
   MSComm1.Settings = "38400,N,8,2"
   frame_delay = False
ElseIf Combo_baud0.Text = "57600" Then
   MSComm1.Settings = "57600,N,8,2"
   frame_delay = False
End If

End Sub
Private Sub button_custom_Click()
If button_custom.Caption = "E26XX" And Form1.Frame_E222X.Caption = "E26XX" Then
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   button_custom.Caption = "CUSTOM"
ElseIf button_custom.Caption = "CUSTOM" And Form1.Frame_E222X.Caption = "E26XX" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   button_custom.Caption = "E26XX"
ElseIf button_custom.Caption = "E22XX" And Form1.Frame_E222X.Caption = "E22XX" Then
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   button_custom.Caption = "CUSTOM"
ElseIf button_custom.Caption = "CUSTOM" And Form1.Frame_E222X.Caption = "E22XX" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   button_custom.Caption = "E22XX"
ElseIf button_custom.Caption = "PVT100" And Form1.Frame_E222X.Caption = "PVT100" Then
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   button_custom.Caption = "CUSTOM"
ElseIf button_custom.Caption = "CUSTOM" And Form1.Frame_E222X.Caption = "PVT100" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   button_custom.Caption = "PVT100"
ElseIf button_custom.Caption = "PVT10" And Form1.Frame_E222X.Caption = "PVT10" Then
   Frame_E222X.Visible = True
   Frame_Custom.Visible = False
   button_custom.Caption = "CUSTOM"
ElseIf button_custom.Caption = "CUSTOM" And Form1.Frame_E222X.Caption = "PVT10" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = True
   button_custom.Caption = "PVT10"
ElseIf button_custom.Caption = "HOME" Then
   Frame_E222X.Visible = False
   Frame_Custom.Visible = False
   Homepage.Visible = True
   button_custom.Caption = "CUSTOM"
End If
End Sub
'SEERIANUMBER
Private Sub Text_SN_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_SN.Text) Then
   temp = CDbl(Text_SN.Text)
   If temp > 0 And temp < 65536 Then
   Text_SN.Text = temp
   Else
   Text_SN.Text = "1"
   End If
Else
   Text_SN.Text = "1"
End If
End Sub
'HARDWARE
Private Sub Text_hardware_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_hardware.Text) Then
   temp = CDbl(Text_hardware.Text)
   If temp > 0 And temp < 65536 Then
   Text_hardware.Text = temp
   Else
   Text_hardware.Text = "2608"
   End If
Else
   Text_hardware.Text = "2608"
End If
End Sub
'SLAVE ID
Private Sub Text_slave_id0_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If Form1.Frame_E222X.Caption = "E26XX" Then
 If IsNumeric(Text_slave_id0.Text) Then
   temp = CDbl(Text_slave_id0.Text)
   If temp > 0 And temp < 248 Then
   slave_id_global = temp ' SET GLOBAL SLAVE ID
   Text_slave_id0.Text = temp
   Else
   Text_slave_id0.Text = "1"
   End If
 Else
   Text_slave_id0.Text = "1"
 End If
ElseIf Form1.Frame_E222X.Caption = "E26XX" Then
 If IsNumeric(Text_slave_id0.Text) Then
   temp = CDbl(Text_slave_id0.Text)
   If temp > 0 And temp < 248 Then
   slave_id_global = temp ' SET GLOBAL SLAVE ID
   Text_slave_id0.Text = temp
   Else
   Text_slave_id0.Text = "1"
   End If
 Else
   Text_slave_id0.Text = "1"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
 If IsNumeric(Text_slave_id0.Text) Then
   temp = CDbl(Text_slave_id0.Text)
   If temp > 0 And temp < 248 Then
   slave_id_global = temp ' SET GLOBAL SLAVE ID
   Text_slave_id0.Text = temp
   Else
   Text_slave_id0.Text = "16"
   End If
 Else
   Text_slave_id0.Text = "16"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
 If IsNumeric(Text_slave_id0.Text) Then
   temp = CDbl(Text_slave_id0.Text)
   If temp > 0 And temp < 248 Then
   slave_id_global = temp ' SET GLOBAL SLAVE ID
   Text_slave_id0.Text = temp
   Else
   Text_slave_id0.Text = "16"
   End If
 Else
   Text_slave_id0.Text = "16"
 End If
End If
End Sub
'RESPONSE DELAY
Private Sub Text_response_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_response.Text) Then
   temp = CDbl(Text_response.Text)
   If temp > 9 And temp < 256 Then
   Text_response.Text = temp
   Else
   Text_response.Text = "10"
   End If
Else
   Text_response.Text = "10"
End If
End Sub
'CONSTANT B
Private Sub Text_const_B_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_const_B.Text) Then
   temp = CDbl(Text_const_B.Text)
   If temp > -1 And temp < 65536 Then
   Text_const_B.Text = temp
   Else
   Text_const_B.Text = "100"
   End If
Else
   Text_const_B.Text = "100"
End If
End Sub
'CONSTANT C
Private Sub Text_const_C_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_const_C.Text) Then
   temp = CDbl(Text_const_C.Text)
   If temp > -1 And temp < 65536 Then
   Text_const_C.Text = temp
   Else
   Text_const_C.Text = "100"
   End If
Else
   Text_const_C.Text = "100"
End If
End Sub
'CONSTANT D
Private Sub Text_const_D_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_const_D.Text) Then
   temp = CDbl(Text_const_D.Text)
   If temp > -32769 And temp < 32768 Then
   Text_const_D.Text = temp
   Else
   Text_const_D.Text = "100"
   End If
Else
   Text_const_D.Text = "100"
End If
End Sub
'CONSTANT E
Private Sub Text_const_E_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_const_E.Text) Then
   temp = CDbl(Text_const_E.Text)
   If temp > -32769 And temp < 32768 Then
   Text_const_E.Text = temp
   Else
   Text_const_E.Text = "100"
   End If
Else
   Text_const_E.Text = "100"
End If
End Sub
'HEATER PULSE
Private Sub Text_heater_pulse_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_heater_pulse.Text) Then
   temp = CDbl(Text_heater_pulse.Text)
   If temp > -1 And temp < 1001 Then
   Text_heater_pulse.Text = temp
   Else
   Text_heater_pulse.Text = "0"
   End If
Else
   Text_heater_pulse.Text = "0"
End If
End Sub
'SENSOR PULSE
Private Sub Text_sensor_pulse_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_sensor_pulse.Text) Then
   temp = CDbl(Text_sensor_pulse.Text)
   If temp > -1 And temp < 1001 Then
   Text_sensor_pulse.Text = temp
   Else
   Text_sensor_pulse.Text = "0"
   End If
Else
   Text_sensor_pulse.Text = "0"
End If
End Sub
'TEMPERATURE ZERO ADJUSTMENT
Private Sub Text_zero_T_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Single
If IsNumeric(Text_zero_T.Text) Then
   temp = CDbl(Text_zero_T.Text)
   If temp >= -320 And temp <= 320 Then
   Text_zero_T.Text = temp
   Else
   Text_zero_T.Text = "0"
   End If
Else
   Text_zero_T.Text = "0"
End If
End Sub
'GAS ZERO ADJUSTMENT
Private Sub Text_zero_RH_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_zero_RH.Text) Then
   temp = CDbl(Text_zero_RH.Text)
   If temp > -32001 And temp < 32001 Then
   Text_zero_RH.Text = temp
   Else
   Text_zero_RH.Text = "0"
   End If
Else
   Text_zero_RH.Text = "0"
End If
End Sub
'GAS SLOPE
Private Sub Text_RH_slope_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If Form1.Frame_E222X.Caption = "E26XX" Then
 If IsNumeric(Text_RH_slope.Text) Then
   temp = CDbl(Text_RH_slope.Text)
   If temp > 0 And temp < 65536 Then
   Text_RH_slope.Text = temp
   Else
   Text_RH_slope.Text = "512"
   End If
 Else
   Text_RH_slope.Text = "512"
 End If
End If
End Sub
'CHANGE RATE LIMIT
Private Sub Text_RH_rate_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RH_rate.Text) Then
   temp = CDbl(Text_RH_rate.Text)
   If temp > -1 And temp < 32001 Then
   Text_RH_rate.Text = temp
   Else
   Text_RH_rate.Text = "0"
   End If
Else
   Text_RH_rate.Text = "0"
End If
End Sub
'INTEGRATING CONSTANT
Private Sub Text_RC_filter_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RC_filter.Text) Then
   temp = CDbl(Text_RC_filter.Text)
   If temp > 0 And temp < 32001 Then
   Text_RC_filter.Text = temp
   Else
   Text_RC_filter.Text = "0"
   End If
Else
   Text_RC_filter.Text = "0"
End If
End Sub
'ANALOG OUTPUT 1************************************************************************************
Private Sub Combo_AN1_I_U_Click()
If Combo_AN1_I_U.ListIndex = 0 Then
   Text_AN1_4ma.Text = "4"
   Text_AN1_20ma.Text = "20"
ElseIf Combo_AN1_I_U.ListIndex = 1 Then
   Text_AN1_4ma.Text = "0"
   Text_AN1_20ma.Text = "10"
End If
End Sub
'AN1 4ma
Private Sub Text_AN1_4ma_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN1_4ma.Text) Then
   temp = CDbl(Text_AN1_4ma.Text)
   If Combo_AN1_I_U.ListIndex = 0 Then
      If temp > 3 And temp < 21 Then
      Text_AN1_4ma.Text = temp
      Else
      Text_AN1_4ma.Text = "4"
      End If
   ElseIf Combo_AN1_I_U.ListIndex = 1 Then
      If temp > -1 And temp < 11 Then
      Text_AN1_4ma.Text = temp
      Else
      Text_AN1_4ma.Text = "4"
      End If
   End If
   
Else
   Text_AN1_4ma.Text = "4"
   Text_AN1_20ma.Text = "10"
End If

End Sub
'AN1 20ma
Private Sub Text_AN1_20ma_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN1_20ma.Text) Then
   temp = CDbl(Text_AN1_20ma.Text)
   If Combo_AN1_I_U.ListIndex = 0 Then
      If temp > 3 And temp < 21 Then
      Text_AN1_20ma.Text = temp
      Else
      Text_AN1_20ma.Text = "10"
      End If
   ElseIf Combo_AN1_I_U.ListIndex = 1 Then
      If temp > -1 And temp < 11 Then
      Text_AN1_20ma.Text = temp
      Else
      Text_AN1_20ma.Text = "10"
      End If
   End If
   
Else
   Text_AN1_4ma.Text = "4"
   Text_AN1_20ma.Text = "10"
End If

End Sub
Private Sub Text_AN1_4ma_Change()
If Text_AN1_4ma.Text = Text_AN1_20ma.Text Then
   Text_AN1_4ma.Text = "4"
   Text_AN1_20ma.Text = "10"
End If
End Sub
Private Sub Text_AN1_20ma_Change()
If Text_AN1_4ma.Text = Text_AN1_20ma.Text Then
   Text_AN1_4ma.Text = "4"
   Text_AN1_20ma.Text = "10"
End If
End Sub
'ANALOG OUTPUT 2************************************************************************************
Private Sub Combo_AN2_I_U_Click()
If Combo_AN2_I_U.ListIndex = 0 Then
   Text_AN2_4ma.Text = "4"
   Text_AN2_20ma.Text = "20"
ElseIf Combo_AN2_I_U.ListIndex = 1 Then
   Text_AN2_4ma.Text = "0"
   Text_AN2_20ma.Text = "10"
End If
End Sub
'AN2 4ma
Private Sub Text_AN2_4ma_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN2_4ma.Text) Then
   temp = CDbl(Text_AN2_4ma.Text)
   If Combo_AN2_I_U.ListIndex = 0 Then
      If temp > 3 And temp < 21 Then
      Text_AN2_4ma.Text = temp
      Else
      Text_AN2_4ma.Text = "4"
      End If
   ElseIf Combo_AN2_I_U.ListIndex = 1 Then
      If temp > -1 And temp < 11 Then
      Text_AN2_4ma.Text = temp
      Else
      Text_AN2_4ma.Text = "4"
      End If
   End If
   
Else
   Text_AN2_4ma.Text = "4"
   Text_AN2_20ma.Text = "10"
End If

End Sub
'AN2 20ma
Private Sub Text_AN2_20ma_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN2_20ma.Text) Then
   temp = CDbl(Text_AN2_20ma.Text)
   If Combo_AN2_I_U.ListIndex = 0 Then
      If temp > 3 And temp < 21 Then
      Text_AN2_20ma.Text = temp
      Else
      Text_AN2_20ma.Text = "10"
      End If
   ElseIf Combo_AN2_I_U.ListIndex = 1 Then
      If temp > -1 And temp < 11 Then
      Text_AN2_20ma.Text = temp
      Else
      Text_AN2_20ma.Text = "10"
      End If
   End If
   
Else
   Text_AN2_4ma.Text = "4"
   Text_AN2_20ma.Text = "10"
End If

End Sub
Private Sub Text_AN2_4ma_Change()
If Text_AN2_4ma.Text = Text_AN2_20ma.Text Then
   Text_AN2_4ma.Text = "4"
   Text_AN2_20ma.Text = "10"
End If
End Sub
Private Sub Text_AN2_20ma_Change()
If Text_AN2_4ma.Text = Text_AN2_20ma.Text Then
   Text_AN2_4ma.Text = "4"
   Text_AN2_20ma.Text = "10"
End If
End Sub
'ANALOG OUTPUT 1 SCALE******************************************************************************
Private Sub Combo_AN1_onoff_Click()
If Form1.Frame_E222X.Caption = "E26XX" Then
 If Combo_AN1_onoff.ListIndex = 0 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 1 Then
   Text_AN1_0deg.Text = "-40"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 2 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "32000"
 ElseIf Combo_AN1_onoff.ListIndex = 3 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
 If Combo_AN1_onoff.ListIndex = 0 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 1 Then
   Text_AN1_0deg.Text = "-40"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 2 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "100"
 ElseIf Combo_AN1_onoff.ListIndex = 3 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
 If Combo_AN1_onoff.ListIndex = 0 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 1 Then
   Text_AN1_0deg.Text = "-40"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 2 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "95"
 ElseIf Combo_AN1_onoff.ListIndex = 3 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
 If Combo_AN1_onoff.ListIndex = 0 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 1 Then
   Text_AN1_0deg.Text = "-40"
   Text_AN1_100deg.Text = "85"
 ElseIf Combo_AN1_onoff.ListIndex = 2 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "100"
 ElseIf Combo_AN1_onoff.ListIndex = 3 Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
 End If
End If
End Sub
Private Sub Text_AN1_0deg_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN1_0deg.Text) Then
   temp = CDbl(Text_AN1_0deg.Text)
   If Combo_AN1_onoff.ListIndex = 1 Then
      If temp > -41 And temp < 86 Then
      Text_AN1_0deg.Text = temp
      Else
      Text_AN1_0deg.Text = "0"
      End If
   ElseIf Combo_AN1_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_AN1_0deg.Text = temp
      Else
      Text_AN1_0deg.Text = "0"
      End If
   Else
      Text_AN1_0deg.Text = "0"
      Text_AN1_100deg.Text = "85"
   End If
   
Else
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
End If

End Sub
Private Sub Text_AN1_100deg_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN1_100deg.Text) Then
   temp = CDbl(Text_AN1_100deg.Text)
   If Combo_AN1_onoff.ListIndex = 1 Then
      If temp > -41 And temp < 86 Then
      Text_AN1_100deg.Text = temp
      Else
      Text_AN1_100deg.Text = "85"
      End If
   ElseIf Combo_AN1_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_AN1_100deg.Text = temp
      Else
      Text_AN1_100deg.Text = "85"
      End If
   Else
      Text_AN1_0deg.Text = "0"
      Text_AN1_100deg.Text = "85"
   End If
   
Else
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
End If

End Sub
Private Sub Text_AN1_0deg_Change()
If Text_AN1_0deg.Text = Text_AN1_100deg.Text Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
End If
End Sub
Private Sub Text_AN1_100deg_Change()
If Text_AN1_0deg.Text = Text_AN1_100deg.Text Then
   Text_AN1_0deg.Text = "0"
   Text_AN1_100deg.Text = "85"
End If
End Sub
'ANALOG OUTPUT 2 SCALE******************************************************************************
Private Sub Combo_AN2_onoff_Click()
If Form1.Frame_E222X.Caption = "E26XX" Then
 If Combo_AN2_onoff.ListIndex = 0 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 1 Then
   Text_AN2_0deg.Text = "-40"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 2 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "32000"
 ElseIf Combo_AN2_onoff.ListIndex = 3 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT10" Then
 If Combo_AN2_onoff.ListIndex = 0 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 1 Then
   Text_AN2_0deg.Text = "-20"
   Text_AN2_100deg.Text = "70"
 ElseIf Combo_AN2_onoff.ListIndex = 2 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 3 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "PVT100" Then
 If Combo_AN2_onoff.ListIndex = 0 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 1 Then
   Text_AN2_0deg.Text = "-40"
   Text_AN2_100deg.Text = "80"
 ElseIf Combo_AN2_onoff.ListIndex = 2 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 3 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 End If
ElseIf Form1.Frame_E222X.Caption = "E22XX" Then
 If Combo_AN2_onoff.ListIndex = 0 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 1 Then
   Text_AN2_0deg.Text = "-40"
   Text_AN2_100deg.Text = "80"
 ElseIf Combo_AN2_onoff.ListIndex = 2 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 ElseIf Combo_AN2_onoff.ListIndex = 3 Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
 End If
End If
End Sub
Private Sub Text_AN2_0deg_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN2_0deg.Text) Then
   temp = CDbl(Text_AN2_0deg.Text)
   If Combo_AN2_onoff.ListIndex = 1 Then
      If temp > -41 And temp < 86 Then
      Text_AN2_0deg.Text = temp
      Else
      Text_AN2_0deg.Text = "0"
      End If
   ElseIf Combo_AN2_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_AN2_0deg.Text = temp
      Else
      Text_AN2_0deg.Text = "0"
      End If
   Else
      Text_AN2_0deg.Text = "0"
      Text_AN2_100deg.Text = "85"
   End If
   
Else
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
End If

End Sub
Private Sub Text_AN2_100deg_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_AN2_100deg.Text) Then
   temp = CDbl(Text_AN2_100deg.Text)
   If Combo_AN2_onoff.ListIndex = 1 Then
      If temp > -41 And temp < 86 Then
      Text_AN2_100deg.Text = temp
      Else
      Text_AN2_100deg.Text = "85"
      End If
   ElseIf Combo_AN2_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_AN2_100deg.Text = temp
      Else
      Text_AN2_100deg.Text = "85"
      End If
   Else
      Text_AN2_0deg.Text = "0"
      Text_AN2_100deg.Text = "85"
   End If
   
Else
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
End If

End Sub
Private Sub Text_AN2_0deg_Change()
If Text_AN2_0deg.Text = Text_AN2_100deg.Text Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
End If
End Sub
Private Sub Text_AN2_100deg_Change()
If Text_AN2_0deg.Text = Text_AN2_100deg.Text Then
   Text_AN2_0deg.Text = "0"
   Text_AN2_100deg.Text = "85"
End If
End Sub
'RE1 DELAY
Private Sub Text_RE1_delay_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RE1_delay.Text) Then
   temp = CDbl(Text_RE1_delay.Text)
   If temp > -1 And temp < 1001 Then
   Text_RE1_delay.Text = temp
   Else
   Text_RE1_delay.Text = "0"
   End If
Else
   Text_RE1_delay.Text = "0"
End If
End Sub
'RE2 DELAY
Private Sub Text_RE2_delay_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RE2_delay.Text) Then
   temp = CDbl(Text_RE2_delay.Text)
   If temp > -1 And temp < 1001 Then
   Text_RE2_delay.Text = temp
   Else
   Text_RE2_delay.Text = "0"
   End If
Else
   Text_RE2_delay.Text = "0"
End If
End Sub
'RE1 ON OFF TIME
Private Sub Text_RE1_time_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RE1_time.Text) Then
   temp = CDbl(Text_RE1_time.Text)
   If temp > -1 And temp < 1001 Then
   Text_RE1_time.Text = temp
   Else
   Text_RE1_time.Text = "0"
   End If
Else
   Text_RE1_time.Text = "0"
End If
End Sub
'RE2 ON OFF TIME
Private Sub Text_RE2_time_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
If IsNumeric(Text_RE2_time.Text) Then
   temp = CDbl(Text_RE2_time.Text)
   If temp > -1 And temp < 1001 Then
   Text_RE2_time.Text = temp
   Else
   Text_RE2_time.Text = "0"
   End If
Else
   Text_RE2_time.Text = "0"
End If
End Sub
'RE1 SETPOINTS******************************************************************************
Private Sub Combo_RE1_onoff_Click()
If Combo_RE1_onoff.ListIndex = 0 Then
   Text_RE1_L.Text = "0"
   Text_RE1_H.Text = "85"
ElseIf Combo_RE1_onoff.ListIndex = 1 Then
   Text_RE1_L.Text = "-40"
   Text_RE1_H.Text = "85"
ElseIf Combo_RE1_onoff.ListIndex = 2 Then
   Text_RE1_L.Text = "0"
   Text_RE1_H.Text = "32000"
ElseIf Combo_RE1_onoff.ListIndex = 3 Then
   Text_RE1_L.Text = "0"
   Text_RE1_H.Text = "85"
End If

End Sub
Private Sub Text_RE1_L_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
Dim temp2 As Single
If IsNumeric(Text_RE1_L.Text) Then
   temp = CDbl(Text_RE1_L.Text)
   temp2 = CDbl(Text_RE1_L.Text)
   If Combo_RE1_onoff.ListIndex = 1 Then
      If temp2 >= -40 And temp2 <= 85 Then
      Text_RE1_L.Text = temp2
      Else
      Text_RE1_L.Text = "-40"
      End If
   ElseIf Combo_RE1_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_RE1_L.Text = temp
      Else
      Text_RE1_L.Text = "0"
      End If
   Else
      Text_RE1_L.Text = "0"
   End If
   
Else
   Text_RE1_L.Text = "0"
End If

End Sub
Private Sub Text_RE1_H_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
Dim temp2 As Single
If IsNumeric(Text_RE1_H.Text) Then
   temp = CDbl(Text_RE1_H.Text)
   temp2 = CDbl(Text_RE1_H.Text)
   If Combo_RE1_onoff.ListIndex = 1 Then
      If temp2 >= -40 And temp2 <= 85 Then
      Text_RE1_H.Text = temp2
      Else
      Text_RE1_H.Text = "85"
      End If
   ElseIf Combo_RE1_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_RE1_H.Text = temp
      Else
      Text_RE1_H.Text = "32000"
      End If
   Else
      Text_RE1_H.Text = "85"
   End If
   
Else
   Text_RE1_H.Text = "85"
End If

End Sub
Private Sub Text_RE1_L_Change()
If Text_RE1_L.Text = Text_RE1_H.Text Then
   Text_RE1_L.Text = "0"
   Text_RE1_H.Text = "85"
End If
End Sub
Private Sub Text_RE1_H_Change()
If Text_RE1_L = Text_RE1_H.Text Then
   Text_RE1_L.Text = "0"
   Text_RE1_H.Text = "85"
End If
End Sub
'RE2 SETPOINTS******************************************************************************
Private Sub Combo_RE2_onoff_Click()
If Combo_RE2_onoff.ListIndex = 0 Then
   Text_RE2_L.Text = "0"
   Text_RE2_H.Text = "85"
ElseIf Combo_RE2_onoff.ListIndex = 1 Then
   Text_RE2_L.Text = "-40"
   Text_RE2_H.Text = "85"
ElseIf Combo_RE2_onoff.ListIndex = 2 Then
   Text_RE2_L.Text = "0"
   Text_RE2_H.Text = "32000"
ElseIf Combo_RE2_onoff.ListIndex = 3 Then
   Text_RE2_L.Text = "0"
   Text_RE2_H.Text = "85"
End If

End Sub
Private Sub Text_RE2_L_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
Dim temp2 As Single
If IsNumeric(Text_RE2_L.Text) Then
   temp = CDbl(Text_RE2_L.Text)
   temp2 = CDbl(Text_RE2_L.Text)
   If Combo_RE2_onoff.ListIndex = 1 Then
      If temp2 >= -40 And temp2 <= 85 Then
      Text_RE2_L.Text = temp2
      Else
      Text_RE2_L.Text = "-40"
      End If
   ElseIf Combo_RE2_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_RE2_L.Text = temp
      Else
      Text_RE2_L.Text = "0"
      End If
   Else
      Text_RE2_L.Text = "0"
   End If
   
Else
   Text_RE2_L.Text = "0"
End If

End Sub
Private Sub Text_RE2_H_Validate(Cancel As Boolean) 'Texbox-de valideerimine: et oleksid numbrid ja et oleksid õiges vahemikus
Dim temp As Long
Dim temp2 As Single
If IsNumeric(Text_RE2_H.Text) Then
   temp = CDbl(Text_RE2_H.Text)
   temp2 = CDbl(Text_RE2_H.Text)
   If Combo_RE2_onoff.ListIndex = 1 Then
      If temp2 >= -40 And temp2 <= 85 Then
      Text_RE2_H.Text = temp2
      Else
      Text_RE2_H.Text = "85"
      End If
   ElseIf Combo_RE2_onoff.ListIndex = 2 Then
      If temp > -1 And temp < 32001 Then
      Text_RE2_H.Text = temp
      Else
      Text_RE2_H.Text = "32000"
      End If
   Else
      Text_RE2_H.Text = "85"
   End If
   
Else
   Text_RE2_H.Text = "85"
End If

End Sub
Private Sub Text_RE2_L_Change()
If Text_RE2_L.Text = Text_RE2_H.Text Then
   Text_RE2_L.Text = "0"
   Text_RE2_H.Text = "85"
End If
End Sub
Private Sub Text_RE2_H_Change()
If Text_RE2_L = Text_RE2_H.Text Then
   Text_RE2_L.Text = "0"
   Text_RE2_H.Text = "85"
End If
End Sub

'*************************************************************************HETKEL EI KASUTA>>>
'Private Sub Text_SN_GotFocus() 'tekstikasti ära märgistamine
'SelectAllText Text_SN
'End Sub
'ÜHINE ALAMPROGRAMM
'Private Sub SelectAllText(tb As TextBox) 'Vajalik et klikkides märgistatakse kogu tekst
'tb.SelStart = 0
'tb.SelLength = Len(tb.Text)
'End Sub
