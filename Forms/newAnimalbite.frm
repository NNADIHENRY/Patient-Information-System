VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnimalBite 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8835
   ClientLeft      =   2265
   ClientTop       =   1740
   ClientWidth     =   11505
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00C9D5BB&
      Height          =   3375
      Left            =   -600
      TabIndex        =   75
      Top             =   120
      Width           =   13695
      Begin MSComctlLib.ListView lstAnimalBite 
         Height          =   3135
         Left            =   720
         TabIndex        =   77
         Top             =   120
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   1920
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.Frame frlist 
      BackColor       =   &H00C9D5BB&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   -120
      TabIndex        =   78
      Top             =   3480
      Visible         =   0   'False
      Width           =   13695
      Begin VB.TextBox txtage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10200
         TabIndex        =   121
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10080
         TabIndex        =   80
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtaphycisian 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   99
         Top             =   4680
         Width           =   3855
      End
      Begin VB.TextBox txtatreatment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   98
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtanimalvac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox txtacondition 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   96
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtasite 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         TabIndex        =   95
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtanature 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         TabIndex        =   94
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtaheight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8880
         TabIndex        =   93
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtaweight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         TabIndex        =   92
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtabmi 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10200
         TabIndex        =   91
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtabrgy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   90
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtaaddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1905
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtamiddle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1905
         TabIndex        =   88
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtafirst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1905
         TabIndex        =   87
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtalast 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1905
         TabIndex        =   86
         Top             =   135
         Width           =   2895
      End
      Begin VB.ComboBox cboagender 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":0000
         Left            =   6600
         List            =   "newAnimalbite.frx":000A
         TabIndex        =   85
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboastatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":001C
         Left            =   6600
         List            =   "newAnimalbite.frx":002C
         TabIndex        =   84
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtaprescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Top             =   3360
         Width           =   3975
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   81
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   79
         Top             =   4560
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Date 
         Height          =   375
         Left            =   9360
         TabIndex        =   82
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88997889
         CurrentDate     =   40153
      End
      Begin MSComCtl2.DTPicker Date2 
         Height          =   375
         Left            =   9360
         TabIndex        =   100
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88997889
         CurrentDate     =   40153
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   122
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label88 
         BackStyle       =   0  'Transparent
         Caption         =   "Physician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   120
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label87 
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   119
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label86 
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccination of  Animal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   118
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label85 
         BackStyle       =   0  'Transparent
         Caption         =   "Condition of Animal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   117
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Site of Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   116
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label83 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   115
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   114
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   113
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   112
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "BML"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   111
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   110
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8040
         TabIndex        =   109
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   108
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   107
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   106
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   105
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   104
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "lbs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   103
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Prescription"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   102
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bite Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   101
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2520
      Top             =   600
   End
   Begin VB.Frame frameAnimalbite 
      BackColor       =   &H00404000&
      Caption         =   "Medical Records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   11295
      Begin VB.ComboBox cboYear 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Bite_Year"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":0053
         Left            =   9840
         List            =   "newAnimalbite.frx":0075
         TabIndex        =   68
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboDay 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Bite_Day"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":00B5
         Left            =   9000
         List            =   "newAnimalbite.frx":0119
         TabIndex        =   67
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cbomonth 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Bite_Month"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":0193
         Left            =   6960
         List            =   "newAnimalbite.frx":01BB
         TabIndex        =   66
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cboBite_Nat 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "newAnimalbite.frx":0220
         Left            =   2280
         List            =   "newAnimalbite.frx":0230
         TabIndex        =   65
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cboBiteSite 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "newAnimalbite.frx":0260
         Left            =   2280
         List            =   "newAnimalbite.frx":027C
         TabIndex        =   41
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox cboVac_Ani 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "newAnimalbite.frx":02B3
         Left            =   2280
         List            =   "newAnimalbite.frx":02BD
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtphysician 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Physician"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7800
         TabIndex        =   21
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txttreatment 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Treatment/Remarks"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtcondition 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Condition_of_Animal"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   19
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Index           =   0
         Left            =   9120
         TabIndex        =   72
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   5
         Left            =   9960
         TabIndex        =   71
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   70
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Bite:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   69
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   39
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Physician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   6
         Left            =   6720
         TabIndex        =   26
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment/ Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Index           =   5
         Left            =   5640
         TabIndex        =   25
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccination of Animal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Condition of    Animal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Site of Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404000&
      Caption         =   "Prescription"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   27
      Top             =   6480
      Width           =   11295
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   64
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearConsultation 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   63
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveConsultation 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   62
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewConsultationRecords 
         Caption         =   "View Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   61
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optPrescription1 
         BackColor       =   &H00404000&
         Caption         =   "Capsule/Tablets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optPrescription2 
         BackColor       =   &H00404000&
         Caption         =   "Suspension"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Frame frPrescription 
         BackColor       =   &H00404000&
         Caption         =   "Prescription"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Index           =   1
         Left            =   2520
         TabIndex        =   54
         Top             =   360
         Width           =   7935
         Begin VB.ComboBox cboMedicineMeasurementQTY 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "newAnimalbite.frx":02CA
            Left            =   1800
            List            =   "newAnimalbite.frx":02E9
            TabIndex        =   58
            Text            =   "Qty."
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboMedicineMeasurement 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "newAnimalbite.frx":0308
            Left            =   2760
            List            =   "newAnimalbite.frx":0312
            TabIndex        =   57
            Text            =   "Teaspoon"
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox cboPQTY 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "newAnimalbite.frx":032C
            Left            =   4320
            List            =   "newAnimalbite.frx":034B
            TabIndex        =   56
            Text            =   "Qty."
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboMedicine_x 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "newAnimalbite.frx":036A
            Left            =   6240
            List            =   "newAnimalbite.frx":0374
            TabIndex        =   55
            Text            =   "Medicine"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Suspension:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "X a Day"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   5280
            TabIndex        =   59
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame frPrescription 
         BackColor       =   &H00404000&
         Caption         =   "Prescription"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Index           =   0
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   7935
         Begin VB.ComboBox cboQTY 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            ItemData        =   "newAnimalbite.frx":0393
            Left            =   2160
            List            =   "newAnimalbite.frx":03A6
            TabIndex        =   31
            Text            =   "Qty."
            Top             =   480
            Width           =   855
         End
         Begin VB.ComboBox cboMedicine 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            ItemData        =   "newAnimalbite.frx":03C7
            Left            =   4200
            List            =   "newAnimalbite.frx":03D4
            TabIndex        =   30
            Text            =   "Medicine"
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox cboMedicineDays 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "newAnimalbite.frx":03F9
            Left            =   6240
            List            =   "newAnimalbite.frx":0406
            TabIndex        =   29
            Text            =   "No."
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Capsule/Tablets:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "X  a Day"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   34
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "for"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   33
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   7200
            TabIndex        =   32
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Prescription Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   11295
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5040
         TabIndex        =   73
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88997889
         CurrentDate     =   37257
      End
      Begin VB.TextBox txtpatientID 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtbml 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BML"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8880
         TabIndex        =   46
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtHeight 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Height"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5040
         TabIndex        =   45
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtWeight 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Weight"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7080
         TabIndex        =   44
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Status"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":041B
         Left            =   5040
         List            =   "newAnimalbite.frx":042B
         TabIndex        =   43
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cboGender 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Gender"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "newAnimalbite.frx":0452
         Left            =   5040
         List            =   "newAnimalbite.frx":045C
         TabIndex        =   42
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Barangay"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtBarangay 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Barangay"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5040
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Last_Name"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFC0&
         DataField       =   "First_Name"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtMiddleName 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Middle_Name"
         DataSource      =   "adoanimal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "B-Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   74
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   22
         Left            =   4320
         TabIndex        =   51
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   21
         Left            =   6360
         TabIndex        =   50
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BML"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   20
         Left            =   8400
         TabIndex        =   49
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   47
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   25
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   8
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.ComboBox cboCase 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "newAnimalbite.frx":046E
      Left            =   9720
      List            =   "newAnimalbite.frx":0490
      TabIndex        =   52
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Case No:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   53
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Registration:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   19
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "09/23/09"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 AM"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ANIMAL BITE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmAnimalBite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prescription As String
Dim str   As String
Dim txtids  As String

Private Function AutoID()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_animalbite_information Order By patient_id desc", databaseconnection, 3, 2
    
       If recordset.RecordCount = 0 Then
            txtpatientID.Text = "DR-0001"
        Else
            txtpatientID.Text = "DR-0000" + Format(Right(recordset!patient_id, 4) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtpatientID.Locked = True
End Function
Sub Animal_Bite()
            recordset!CaseNo = cboCase
            recordset!patient_id = txtpatientID
            recordset!lastname = txtLastName
            recordset!firstname = txtFirstName
            recordset!middlename = txtMiddleName
            recordset!Address = txtaddress
            recordset!bitedate = cboMonth & "-" & cboDay & "-" & cboYear
            recordset!barangay = txtbarangay
            recordset!Gender = cbogender
            recordset!Status = cbostatus
            recordset!Height = txtHeight
            recordset!bml = txtbml
            recordset!Site_of_Bite = cboBiteSite
            recordset!Nature_of_Bite = cboBite_Nat
            recordset!Animal_Conidtion = txtcondition
            recordset!Vaccination_of_Animal = cboVac_Ani
            recordset!Treatment = txttreatment
            recordset!Physician = txtphysician
            recordset!prescription = prescription & "-" & cboMedicineMeasurementQTY & "-" & cboMedicineMeasurement & "-" & cboPQTY & "-" & cboMedicine_x
            recordset!Birthdate = DTPicker1

End Sub
Private Sub LoadAnimalBiteRecords()
    On Error Resume Next
  Set recordset = New ADODB.recordset
    Dim intctr As Integer
    intctr = 0
    lstAnimalBite.ListItems.clear
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_animalbite_information", databaseconnection, adOpenDynamic, adLockPessimistic
    If Not recordset.BOF Then
        Do Until recordset.EOF
            Set gListItem = lstAnimalBite.ListItems.Add(, , recordset(0))
                gListItem.SubItems(1) = recordset(1)
                gListItem.SubItems(2) = recordset(2)
                gListItem.SubItems(3) = recordset(3)
                gListItem.SubItems(4) = recordset(4)
                gListItem.SubItems(5) = recordset(5)
                gListItem.SubItems(6) = recordset(6)
                gListItem.SubItems(7) = recordset(7)
                gListItem.SubItems(8) = recordset(8)
                gListItem.SubItems(9) = recordset(11)
                gListItem.SubItems(10) = recordset(12)
                gListItem.SubItems(11) = recordset(10)
                gListItem.SubItems(12) = recordset(20)
                gListItem.SubItems(13) = recordset(13)
                gListItem.SubItems(14) = recordset(15)
                gListItem.SubItems(15) = recordset(14)
                gListItem.SubItems(16) = recordset(18)
                gListItem.SubItems(17) = recordset(9)
                gListItem.SubItems(18) = recordset(16)
                gListItem.SubItems(19) = recordset(17)
                'a.SubItems(20) = recordset(17)
               
               recordset.MoveNext
               
               intctr = intctr + 1
               
               
        Loop
    End If
End Sub


Private Sub cmdclose_Click()
frlist.Visible = False
Frame3.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
Set recordset3 = New ADODB.recordset
        recordset3.Open "SELECT * FROM patient_animalbite_information where Patient_Id='" & lstAnimalBite.SelectedItem & " '", databaseconnection, 3, 3
        recordset3("lastname") = txtalast.Text
        recordset3("firstname") = txtafirst.Text
        recordset3("middlename") = txtamiddle.Text
        recordset3("address") = txtaaddress.Text
        recordset3("barangay") = txtabrgy.Text
        recordset3("Animal_Conidtion") = txtacondition.Text
        recordset3("Vaccination_of_Animal") = txtanimalvac.Text
        recordset3("Physician") = txtaphycisian.Text
        recordset3("Weight") = txtaweight.Text
        recordset3("Height") = txtaheight
        recordset3("bml") = txtabmi.Text
        recordset3("gender") = cboagender.Text
        recordset3("Status") = cboastatus.Text
        recordset3("birthdate") = Date
        recordset3("Nature_of_Bite") = txtanature.Text
        recordset3("bitedate") = Date2.Value
        recordset3("Site_of_Bite") = txtasite.Text
        recordset3("Treatment") = txtatreatment.Text
        'recordset3("Prescription") = txtaprescription.Text
        recordset3.Update
        lstAnimalBite.Refresh
MsgBox "Record Successfully Updated!", vbInformation
LoadAnimalBiteRecords

End Sub
Private Sub cmdprint_Click()
Set recordset = New ADODB.recordset
If (txtids = "") Then
 MsgBox "Please choose patient record.", vbCritical, "No data to be printed"
Exit Sub
Else:
str = "SELECT * FROM patient_animalbite_information WHERE Patient_Id ='" & txtids & "'"
recordset.Open str, databaseconnection, 3, 3
'recordset2.Close
Set rptanimalbite.DataSource = recordset

rptanimalbite.Show vbModal
End If
End Sub


Private Sub cmdSaveConsultation_Click()
    
 Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_animalbite_information", databaseconnection, adOpenDynamic, adLockPessimistic

    If optPrescription1.Value = True Then
        prescription = optPrescription1.Caption
    Else
        prescription = optPrescription2.Caption
    End If
    
    recordset.AddNew
   ' On Error Resume Next
    Animal_Bite
    recordset.Update
    recordset.Close
    
    '_____________________________________
        Set recordset = New ADODB.recordset
    
    
        recordset.Open "Select * From animalbite_chart Where (Year = '" & cboYear.Text & "'" & "And Month='" & cboMonth.Text & "')", databaseconnection, adOpenStatic, adLockOptimistic
         'recordset.Open "Select * From animalbite_chart Where mONTH = '" & cbomonth.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
        'recordset.Open "Select * From animalbite_chart Where Year = '" & cboYear.Text & "'" & "Where Month = '" & cbomonth.Text & " '", databaseconnection, adOpenStatic, adLockOptimistic
        
        If cboYear.Text = 2000 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2001 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2002 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2003 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2004 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2005 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
         ElseIf cboYear.Text = 2006 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
          ElseIf cboYear.Text = 2007 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
        ElseIf cboYear.Text = 2008 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
        ElseIf cboYear.Text = 2009 Then
                    recordset!NumberBites = recordset!NumberBites + 1
                    
                If cboBite_Nat.Text = "DogBite" Then
                    recordset!DogBite = recordset!DogBite + 1
                ElseIf cboBite_Nat.Text = "InsectBite" Then
                    recordset!InsectBite = recordset!InsectBite + 1
                ElseIf cboBite_Nat.Text = "MonkeyBite" Then
                    recordset!MOnkeyBite = recordset!MOnkeyBite + 1
                ElseIf cboBite_Nat.Text = "SnakeBite" Then
                    recordset!SnakeBite = recordset!SnakeBite + 1
                End If
        End If
        
        
        
        recordset.Update
        
        recordset.Close
    '_________________________________
    
        MsgBox "Data Has Been Successfully Saved!", vbInformation, "Congratulations"
    
End Sub


Private Sub cmdViewConsultationRecords_Click()
frlist.Visible = True
frlist.Visible = True
Frame3.Visible = True
LoadAnimalBiteRecords

End Sub
Private Sub LvA_Initialize()
  Lv_SetView lstAnimalBite, "PATIENT_ID|LAST NAME|FIRST NAME|MIDDLE NAME|ADDRESS|BARANGAY|HEIGHT|WEIGHT|BML|GENDER|STATUS|AGE|BIRTHDATE|ANIMAL_CONDITION|NATURE_OF_BITE|VACCINATION_OF_ANIMAL|PHYSICIAN|BITE_DATE|SITE_OF_BITE|TREATMENT"
  With lstAnimalBite
    .ColumnHeaders(1).Width = 1500
    .ColumnHeaders(2).Width = 1500
    .ColumnHeaders(3).Width = 1500
    .ColumnHeaders(4).Width = 1500
    .ColumnHeaders(5).Width = 1500
    .ColumnHeaders(6).Width = 1500
    .ColumnHeaders(7).Width = 2500
    .ColumnHeaders(8).Width = 1500
    .ColumnHeaders(9).Width = 1500
    .ColumnHeaders(10).Width = 1500
    .ColumnHeaders(11).Width = 1500
    .ColumnHeaders(12).Width = 1500
    .ColumnHeaders(13).Width = 1500
    .ColumnHeaders(14).Width = 1500
    .ColumnHeaders(15).Width = 2500
    .ColumnHeaders(16).Width = 2500
    .ColumnHeaders(17).Width = 2500
    .ColumnHeaders(18).Width = 2500
    .ColumnHeaders(19).Width = 2500
    .ColumnHeaders(20).Width = 2500
    .ColumnHeaders(21).Width = 2500
 '   .ColumnHeaders(22).Width = 1000


  End With
End Sub
Private Sub Form_Load()
Dim a As Integer
 Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_animalbite_information", databaseconnection, 3, 3
     If recordset.RecordCount > 0 Then
        Do Until recordset.EOF
           ' a = recordset!patient_id
            recordset.MoveNext
        Loop
    End If
    On Error Resume Next
LvA_Initialize
Frame3.Visible = False
    AutoID
End Sub

Private Sub lstanimalbite_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_animalbite_information where Patient_Id='" & lstAnimalBite.SelectedItem & " '", databaseconnection, 3, 3
             txtids = recordset!patient_id
             txtalast.Text = recordset!lastname
             txtafirst.Text = recordset!firstname
             txtamiddle.Text = recordset!middlename
             txtaaddress.Text = recordset!Address
             txtabrgy.Text = recordset!barangay
             txtacondition.Text = recordset!Animal_Conidtion
             txtanimalvac.Text = recordset!Vaccination_of_Animal
             txtaphycisian.Text = recordset!Physician
             txtaweight.Text = recordset!Weight
             txtaheight.Text = recordset!Height
             txtabmi.Text = recordset!bml
             cboagender.Text = recordset!Gender
             cboastatus.Text = recordset!Status
             txtage.Text = recordset!Age
             Date.Value = recordset!Birthdate
             txtanature.Text = recordset!Nature_of_Bite
             Date2.Value = recordset!bitedate
             txtasite.Text = recordset!Site_of_Bite
             txtatreatment.Text = recordset!Treatment
             txtaprescription.Text = recordset!prescription
            
             cmdsave.Enabled = True

End Sub

Private Sub txtHeight_Change()
    txtbml = Val(txtHeight)
End Sub
Private Sub txtWeight_Change()
    txtbml = Val(txtHeight) * Val(txtWeight)
End Sub
