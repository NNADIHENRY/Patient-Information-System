VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDressing 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8475
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   12570
   ClipControls    =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frdressingframe 
      BackColor       =   &H00C9D5BB&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   12375
      Begin VB.TextBox txtprescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   94
         Top             =   5760
         Width           =   6615
      End
      Begin VB.TextBox txtbarangay1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   79
         Top             =   6840
         Width           =   3495
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
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
         Height          =   510
         Left            =   9120
         TabIndex        =   20
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
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
         Height          =   510
         Left            =   7560
         TabIndex        =   19
         Top             =   7320
         Width           =   1575
      End
      Begin VB.CommandButton cmdCloserecordsdressing 
         Caption         =   "&Close "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   10560
         TabIndex        =   18
         Top             =   7320
         Width           =   1695
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6000
         TabIndex        =   17
         Top             =   7320
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstDressingInfo 
         Height          =   3855
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C9D5BB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Left            =   120
         TabIndex        =   42
         Top             =   4080
         Width           =   12135
         Begin VB.TextBox txtaddress 
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
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox txtmiddle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1440
            TabIndex        =   50
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txtlast 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1440
            TabIndex        =   49
            Top             =   0
            Width           =   3495
         End
         Begin VB.ComboBox cbogender 
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
            ItemData        =   "newdressing.frx":0000
            Left            =   6600
            List            =   "newdressing.frx":000A
            TabIndex        =   48
            Top             =   0
            Width           =   1455
         End
         Begin VB.ComboBox cbostatus 
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
            ItemData        =   "newdressing.frx":001C
            Left            =   10320
            List            =   "newdressing.frx":002C
            TabIndex        =   47
            Top             =   0
            Width           =   1575
         End
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
            Height          =   405
            Left            =   8640
            TabIndex        =   46
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox cbotype 
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
            ItemData        =   "newdressing.frx":0053
            Left            =   1920
            List            =   "newdressing.frx":0069
            TabIndex        =   45
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtfirst 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1440
            TabIndex        =   44
            Top             =   480
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker Date2 
            Height          =   375
            Left            =   6600
            TabIndex        =   43
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
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
            CurrentDate     =   32143
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   10320
            TabIndex        =   95
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
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
            CurrentDate     =   32143
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "DRESSING Date"
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
            Left            =   8520
            TabIndex        =   96
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label71 
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
            Height          =   495
            Left            =   480
            TabIndex        =   61
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Types of Wounds"
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
            Left            =   0
            TabIndex        =   60
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
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
            Left            =   5520
            TabIndex        =   59
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
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
            Left            =   8160
            TabIndex        =   58
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label67 
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
            Height          =   375
            Left            =   9480
            TabIndex        =   57
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label66 
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
            Height          =   375
            Left            =   5520
            TabIndex        =   56
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label65 
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
            Height          =   495
            Left            =   0
            TabIndex        =   55
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label64 
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
            Height          =   495
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label63 
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
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Prescription"
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
            Left            =   5520
            TabIndex        =   52
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.Label Label16 
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
         Height          =   495
         Left            =   600
         TabIndex        =   80
         Top             =   6840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   7815
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   12375
      Begin VB.ComboBox cboMonth 
         BackColor       =   &H00FFFFC0&
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
         ItemData        =   "newdressing.frx":00CA
         Left            =   7560
         List            =   "newdressing.frx":00F2
         TabIndex        =   90
         Top             =   4800
         Width           =   1335
      End
      Begin VB.ComboBox cboDay 
         BackColor       =   &H00FFFFC0&
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
         ItemData        =   "newdressing.frx":0159
         Left            =   9000
         List            =   "newdressing.frx":01BD
         TabIndex        =   89
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtCaseNo 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Age"
         DataSource      =   "adodressing"
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
         Left            =   10320
         MaxLength       =   3
         TabIndex        =   88
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboDressingTypesofWounds 
         BackColor       =   &H00FFFFC0&
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
         ItemData        =   "newdressing.frx":0237
         Left            =   3000
         List            =   "newdressing.frx":0247
         TabIndex        =   85
         Top             =   4680
         Width           =   2535
      End
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
         ItemData        =   "newdressing.frx":027A
         Left            =   10200
         List            =   "newdressing.frx":029C
         TabIndex        =   84
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox txtid 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3120
         TabIndex        =   83
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "G O "
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   82
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
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
         Left            =   1680
         TabIndex        =   81
         Top             =   7200
         Width           =   1455
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
         Height          =   1935
         Left            =   960
         TabIndex        =   62
         Top             =   5160
         Width           =   10575
         Begin VB.OptionButton optcapsule 
            BackColor       =   &H00404000&
            Caption         =   "Capsule/Tablets"
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
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton optsuspension 
            BackColor       =   &H00404000&
            Caption         =   "Suspension"
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
            Left            =   120
            TabIndex        =   74
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Frame frcapsule 
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
            Height          =   1095
            Left            =   2520
            TabIndex        =   63
            Top             =   360
            Width           =   7935
            Begin VB.ComboBox cboQTY2 
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
               ItemData        =   "newdressing.frx":02DC
               Left            =   2160
               List            =   "newdressing.frx":02FE
               TabIndex        =   66
               Text            =   "Qty."
               Top             =   480
               Width           =   855
            End
            Begin VB.ComboBox cboMedicine1 
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
               ItemData        =   "newdressing.frx":033D
               Left            =   4200
               List            =   "newdressing.frx":034A
               TabIndex        =   65
               Text            =   "Medicine"
               Top             =   480
               Width           =   1455
            End
            Begin VB.ComboBox cboDays 
               Appearance      =   0  'Flat
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
               ItemData        =   "newdressing.frx":036F
               Left            =   6240
               List            =   "newdressing.frx":0391
               TabIndex        =   64
               Text            =   "No."
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Capsule/Tablets:"
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
               Height          =   495
               Left            =   240
               TabIndex        =   70
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "X  a Day"
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
               Index           =   0
               Left            =   3120
               TabIndex        =   69
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "for"
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
               Height          =   255
               Index           =   0
               Left            =   5760
               TabIndex        =   68
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
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
               Left            =   7200
               TabIndex        =   67
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.Frame frsuspension 
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
            Height          =   1095
            Left            =   2520
            TabIndex        =   71
            Top             =   360
            Width           =   7935
            Begin VB.ComboBox cboMedicine 
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
               ItemData        =   "newdressing.frx":03D0
               Left            =   6240
               List            =   "newdressing.frx":03E0
               TabIndex        =   13
               Text            =   "Medicine"
               Top             =   480
               Width           =   1455
            End
            Begin VB.ComboBox cboQTY1 
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
               ItemData        =   "newdressing.frx":0409
               Left            =   4320
               List            =   "newdressing.frx":042B
               TabIndex        =   12
               Text            =   "Qty."
               Top             =   480
               Width           =   855
            End
            Begin VB.ComboBox cbomeasurement 
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
               ItemData        =   "newdressing.frx":044E
               Left            =   2640
               List            =   "newdressing.frx":0458
               TabIndex        =   11
               Text            =   "Measurement"
               Top             =   480
               Width           =   1575
            End
            Begin VB.ComboBox cboQTY 
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
               ItemData        =   "newdressing.frx":0472
               Left            =   1680
               List            =   "newdressing.frx":0494
               TabIndex        =   10
               Text            =   "Qty."
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "X a Day"
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
               Left            =   5280
               TabIndex        =   73
               Top             =   520
               Width           =   855
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Suspension:"
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
               Index           =   3
               Left            =   240
               TabIndex        =   72
               Top             =   480
               Width           =   1455
            End
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Prescription Types"
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
            Height          =   495
            Left            =   240
            TabIndex        =   75
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3240
         Top             =   240
      End
      Begin VB.CommandButton cmdViewDressingRecords 
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
         Height          =   375
         Left            =   7200
         TabIndex        =   35
         Top             =   7200
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddDressingRecord 
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   3120
         TabIndex        =   14
         Top             =   7200
         Width           =   2055
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
         Height          =   375
         Left            =   5160
         TabIndex        =   34
         Top             =   7200
         Width           =   2055
      End
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
         Height          =   375
         Left            =   9240
         TabIndex        =   33
         Top             =   7200
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00404000&
         Caption         =   "Patient's Information"
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
         Height          =   3735
         Left            =   960
         TabIndex        =   23
         Top             =   720
         Width           =   10575
         Begin VB.TextBox txtbarangay 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Age"
            DataSource      =   "adodressing"
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
            Left            =   2160
            MaxLength       =   150
            TabIndex        =   4
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtDressingAge 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Age"
            DataSource      =   "adodressing"
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
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   6
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox txtDressingAddress 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Address"
            DataSource      =   "adodressing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1920
            Width           =   5415
         End
         Begin VB.ComboBox cboDressingGen 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Gender"
            DataSource      =   "adodressing"
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
            ItemData        =   "newdressing.frx":04B7
            Left            =   6480
            List            =   "newdressing.frx":04C1
            TabIndex        =   7
            Top             =   3120
            Width           =   1335
         End
         Begin VB.ComboBox cboDressingStat 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Status"
            DataSource      =   "adodressing"
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
            ItemData        =   "newdressing.frx":04D3
            Left            =   2160
            List            =   "newdressing.frx":04E3
            TabIndex        =   5
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtDressingLN 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Last_Name"
            DataSource      =   "adodressing"
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
            Left            =   2160
            TabIndex        =   0
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtDressingFN 
            BackColor       =   &H00FFFFC0&
            DataField       =   "First_Name"
            DataSource      =   "adodressing"
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
            Left            =   4920
            TabIndex        =   1
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtDressingMN 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Middle_Name"
            DataSource      =   "adodressing"
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
            Left            =   7680
            TabIndex        =   2
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00404000&
            Caption         =   "Date of Birth"
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
            Height          =   1095
            Left            =   7920
            TabIndex        =   24
            Top             =   2520
            Width           =   2295
            Begin MSComCtl2.DTPicker Date1 
               Height          =   375
               Left            =   120
               TabIndex        =   8
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
               CurrentDate     =   32143
            End
         End
         Begin VB.Label Label11 
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
            Height          =   495
            Left            =   960
            TabIndex        =   78
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
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
            Left            =   3600
            TabIndex        =   77
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Patient's Name"
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
            Left            =   480
            TabIndex        =   32
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   1080
            TabIndex        =   31
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   5280
            TabIndex        =   30
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   960
            TabIndex        =   29
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label49 
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
            ForeColor       =   &H0080FFFF&
            Height          =   495
            Left            =   2880
            TabIndex        =   28
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label50 
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
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   5640
            TabIndex        =   27
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label52 
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
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   8160
            TabIndex        =   26
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Patient's ID  No."
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
            Index           =   1
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   8760
         TabIndex        =   93
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   10080
         TabIndex        =   92
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   7680
         TabIndex        =   91
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Types of Wounds"
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
         Height          =   495
         Left            =   960
         TabIndex        =   87
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Dressing"
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
         Left            =   5640
         TabIndex        =   86
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 AM"
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
         Height          =   495
         Left            =   5880
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME:"
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
         Left            =   5040
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "09/23/09"
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
         Left            =   3840
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Dressing:"
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
         Height          =   615
         Left            =   1320
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Case Number"
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
         Height          =   495
         Left            =   8520
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1320
      TabIndex        =   76
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   41
      Top             =   -4920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DRESSING FORM"
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
      Left            =   4320
      TabIndex        =   15
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmDressing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPrescription As String, intPrescriptionType As Integer, arrPrescription() As String
Dim PrescriptionFin   As String
Dim str As String
Private Sub cmdAddDressingRecord_Click()
'On Error Resume Next
If txtDressingLN.Text = "" Then
        MsgBox "LAST NAME Required!", vbInformation
        ElseIf txtDressingFN.Text = "" Then
        MsgBox "FIRST NAME Required!", vbInformation
        ElseIf txtDressingMN.Text = "" Then
        MsgBox "MIDDLE NAME Required!", vbInformation
        ElseIf txtDressingAddress.Text = "" Then
        MsgBox "ADDRESS Required!", vbInformation
        ElseIf cboDressingGen.Text = "" Then
        MsgBox "GENDER Required!", vbInformation
        ElseIf cboDressingStat.Text = "" Then
        MsgBox "STATUS Required!", vbInformation
        ElseIf Date1.Value = "" Then
        MsgBox "BIRTHDATE Required!", vbInformation
        ElseIf txtDressingAge.Text = "" Then
        MsgBox "AGE Required!", vbInformation
        ElseIf cboDressingTypesofWounds.Text = "" Then
        MsgBox "TYPE OF WOUND Required!", vbInformation
Else
        PrescriptionInfo
        Set recordset = New ADODB.recordset
        recordset.Open "SELECT * FROM patient_dressing_information", databaseconnection, adOpenDynamic, adLockPessimistic
        recordset.AddNew
        recordset("patient_id") = txtid.Text
        recordset("Dressing_Time") = Time
        recordset("LastName") = txtDressingLN.Text
        recordset("FirstName") = txtDressingFN.Text
        recordset("MiddleName") = txtDressingMN
        recordset("Address") = txtDressingAddress
        recordset("Gender") = cboDressingGen
        recordset("Status") = cboDressingStat
        recordset("Birthdate") = Date1
        recordset("Age") = txtDressingAge
        recordset("Types_of_Wounds") = cboDressingTypesofWounds
        recordset("prescription") = PrescriptionFin
        recordset("Dressing_Date") = cboMonth & "-" & cboDay & "-" & cboYear
        recordset("barangay") = txtbarangay.Text
        recordset("CaseNo") = txtCaseNo.Text
        
        recordset.Update
        recordset.Close
        MsgBox "Record Successfully Added!", vbInformation
End If
'---------------------data will be added to the chart-----------------------
        Set recordset = New ADODB.recordset
        recordset.Open "Select * From Dressing_Chart Where Month = '" & cboMonth.Text & "' and Year = '" & cboYear.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
        
      If recordset.RecordCount > 0 Then
         recordset!Wounds_Num = recordset!Wounds_Num + 1
        If cboDressingTypesofWounds.Text = "Lacerations" Then
            recordset!Lacerations = recordset!Lacerations + 1
        ElseIf cboDressingTypesofWounds.Text = "Abrasions" Then
            recordset!Abrasions = recordset!Abrasions + 1
        ElseIf cboDressingTypesofWounds.Text = "Contusions" Then
            recordset!Contusions = recordset!Contusions + 1
        ElseIf cboDressingTypesofWounds.Text = "Avulsions" Then
            recordset!Avulsions = recordset!Avulsions + 1
        End If
     Else
        recordset.AddNew
        recordset!Wounds_Num = recordset!Wounds_Num + 1
        If cboDressingTypesofWounds.Text = "Lacerations" Then
            recordset!Lacerations = recordset!Lacerations + 1
        ElseIf cboDressingTypesofWounds.Text = "Abrasions" Then
            recordset!Abrasions = recordset!Abrasions + 1
        ElseIf cboDressingTypesofWounds.Text = "Contusions" Then
            recordset!Contusions = recordset!Contusions + 1
        ElseIf cboDressingTypesofWounds.Text = "Avulsions" Then
            recordset!Avulsions = recordset!Avulsions + 1
        End If
            recordset!Year = cboYear.Text
            recordset!Month = cboMonth.Text
     End If
        
        
        recordset.Update
        
        recordset.Close
    '_________________________________
  cmdNew.Enabled = True

End Sub


Private Sub cmdClearConsultation_Click()
txtDressingLN.Text = ""
txtDressingFN.Text = ""
txtDressingMN.Text = ""
txtDressingAddress.Text = ""
cboDressingGen.Text = ""
cboDressingStat.Text = ""
txtDressingAge.Text = ""
cboDressingTypesofWounds.Text = ""
optcapsule.Value = False
optsuspension.Value = False
cboMedicineMeasurementQTY.Text = "QTY"
cboQTY2.Text = "QTY"
cboMedicine1.Text = "Medicine"
cboDays.Text = "No."
cboQTY.Text = "QTY"
cbomeasurement.Text = "Measurement"
cboMedicine.Text = "Medicine"
cboQTY1.Text = "QTY"
txtbarangay.Text = ""
frsuspension.Visible = True
frcapsule.Visible = False
strPrescription = ""
End Sub

Private Sub cmdCloserecordsdressing_Click()
frdressingframe.Visible = False
End Sub

Private Sub cmdedit_Click()
Frame3.Enabled = True
txtlast.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdNew_Click()
     
     cmdsearch.Enabled = False
     Frame5.Enabled = True
     AutoID
     Autocaseno
     cmdNew.Enabled = False
     cmdAddDressingRecord.Enabled = True
End Sub

Private Sub cmdprint_Click()
Set recordset = New ADODB.recordset
If (txtid.Text = "") Then
 MsgBox "Please choose patient record.", vbCritical, "No data to be printed"
Exit Sub
Else:
str = "SELECT * FROM patient_dressing_information WHERE patient_id ='" & txtid.Text & "'"
recordset.Open str, databaseconnection, 3, 3
'recordset2.Close
Set rptDressing.DataSource = recordset2

rptDressing.Show vbModal
End If

End Sub

Private Sub cmdsave_Click()
On Error Resume Next
    If txtlast.Text = "" Then
    MsgBox "LAST NAME Required!", vbInformation
    ElseIf txtfirst.Text = "" Then
    MsgBox "FIRST NAME Required!", vbInformation
    ElseIf txtmiddle.Text = "" Then
    MsgBox "MIDDLE NAME Required!", vbInformation
    ElseIf txtaddress.Text = "" Then
    MsgBox "ADDRESS Required!", vbInformation
    ElseIf cbogender.Text = "" Then
    MsgBox "GENDER Required!", vbInformation
    ElseIf cbostatus.Text = "" Then
    MsgBox "STATUS Required!", vbInformation
    ElseIf txtage.Text = "" Then
    MsgBox "AGE Required!", vbInformation
    ElseIf cbotype.Text = "" Then
    MsgBox "TYPE OF WOUND Required!", vbInformation
Else
    PrescriptionInfo
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_dressing_information where patient_id='" & lstDressingInfo.SelectedItem & " '", databaseconnection, 3, 3
     If newP = True Then
         recordset.AddNew
     End If
    recordset("LastName") = txtlast.Text
    recordset("FirstNAme") = txtfirst.Text
    recordset("MiddleName") = txtmiddle.Text
    recordset("Address") = txtaddress.Text
    recordset("Gender") = cbogender.Text
    recordset("Age") = txtage.Text
    recordset("Status") = cbostatus.Text
    recordset("Birthdate") = Date2
    recordset("Types_of_Wounds") = cbotype.Text
    recordset("prescription") = txtprescription.Text
    recordset("patient_id") = txtid.Text
    recordset.Update
    
      lstDressingInfo.Refresh
MsgBox "Record Successfully Updated!", vbInformation
LoadDressingRecords
End If

End Sub

Private Sub cmdViewDressingRecords_Click()
frdressingframe.Visible = True
LoadDressingRecords
End Sub

Private Sub Lv_Initialize()
  Lv_SetView lstDressingInfo, "PATIENT_ID|DATE|TIME|LAST NAME|FIRST NAME|MIDDLE NAME|ADDRESS|GENDER|AGE|STATUS|BIRTHDATE|TYPES_OF_WOUNDS|PRESCRIPTION|BARANGAY"
  With lstDressingInfo
    .ColumnHeaders(1).Width = 1000
    .ColumnHeaders(2).Width = 2000
    .ColumnHeaders(3).Width = 2000
    .ColumnHeaders(4).Width = 2000
    .ColumnHeaders(5).Width = 2000
    .ColumnHeaders(6).Width = 2000
    .ColumnHeaders(7).Width = 3000
    .ColumnHeaders(8).Width = 1000
    .ColumnHeaders(9).Width = 1000
    .ColumnHeaders(10).Width = 1000
    .ColumnHeaders(11).Width = 2000
    .ColumnHeaders(12).Width = 3000
    .ColumnHeaders(13).Width = 5000
    .ColumnHeaders(14).Width = 2000
  End With
End Sub
Private Sub cmdsearch_Click()
'AutocaseNo
  Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_dressing_information where patient_id='" & txtid.Text & "' order by caseNo desc", databaseconnection, 3, 3
        
       If recordset.RecordCount > 0 Then
            With Me
             .txtDressingLN.Text = recordset!lastname
             .txtDressingFN.Text = recordset!firstname
             .txtDressingMN.Text = recordset!middlename
             .txtDressingAddress.Text = recordset!Address
             .cboDressingGen.Text = recordset!Gender
             .txtDressingAge.Text = recordset!Age
             .cboDressingStat.Text = recordset!Status
             .txtbarangay.Text = recordset!barangay
             .txtCaseNo.Text = recordset!CaseNo + 1
            End With
            Frame5.Enabled = False
            cmdNew.Enabled = False
            cmdAddDressingRecord.Enabled = True
        End If
End Sub

Private Sub Command2_Click()
AutoID
cmdClearConsultation_Click
End Sub
Private Sub Form_Load()
On Error Resume Next
Lv_Initialize
Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM patient_dressing_information", databaseconnection, adOpenDynamic, adLockPessimistic
recordset.MoveLast
'txtid.Text = IIf(IsNull(recordset!patient_id), "", recordset!patient_id)
'LoadDressingRecords
frPrescription.Visible = False
frsuspension.Visible = False
frsuspension.Enabled = False
frcapsule.Enabled = False
strPrescription = ""
End Sub

Private Sub LoadDressingRecords()
    Dim intctr As Integer
    intctr = 0
    lstDressingInfo.ListItems.clear
    Set recordset2 = New ADODB.recordset
    recordset2.Open "SELECT * FROM patient_dressing_information", databaseconnection, adOpenDynamic, adLockPessimistic
    If Not recordset2.BOF Then
        Do Until recordset2.EOF
            Set a = lstDressingInfo.ListItems.Add(, , recordset2(0))
                a.SubItems(1) = recordset2(1)
                a.SubItems(2) = recordset2(2)
                a.SubItems(3) = recordset2(3)
                a.SubItems(4) = recordset2(4)
                a.SubItems(5) = recordset2(5)
                a.SubItems(6) = recordset2(6)
                a.SubItems(7) = recordset2(7)
                a.SubItems(8) = recordset2(10)
                a.SubItems(9) = recordset2(8)
                a.SubItems(10) = recordset2(9)
                a.SubItems(11) = recordset2(11)
                a.SubItems(12) = recordset2(12)
                a.SubItems(13) = recordset2(14)
                
               recordset2.MoveNext
               
               intctr = intctr + 1
               
               
        Loop
    End If
End Sub

Private Function AutoID()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_dressing_information Order By patient_id desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            txtid.Text = "DR-0001"
        Else
            txtid.Text = "DR-000" + Format(Right(recordset!patient_id, 4) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtid.Locked = True
End Function
Private Function Autocaseno()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_dressing_information where patient_id= '" & txtid.Text & "' Order By caseNo desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            txtCaseNo.Text = "01"
        Else
            txtCaseNo.Text = "0" + Format(Right(recordset!CaseNo, 2) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtid.Locked = True
End Function







Private Sub lstDressingInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)

 Set recordset2 = New ADODB.recordset
    recordset2.Open "SELECT * FROM patient_dressing_information where patient_id='" & lstDressingInfo.SelectedItem & " '", databaseconnection, 3, 3
             
             txtlast.Text = recordset2!lastname
             txtfirst.Text = recordset2!firstname
             txtmiddle.Text = recordset2!middlename
             txtaddress.Text = recordset2!Address
             cbogender.Text = recordset2!Gender
             txtage.Text = recordset2!Age
             cbostatus.Text = recordset2!Status
             Date2.Value = recordset2!Birthdate
             cbotype.Text = recordset2!Types_of_Wounds
             txtprescription.Text = recordset2!prescription
             txtbarangay1.Text = recordset2!barangay
             DTPicker1.Value = recordset2!Dressing_Date
             cmdedit.Enabled = True
             cmdsave.Enabled = True
                      
End Sub

Private Sub optcapsule_Click()
frcapsule.Visible = True
frsuspension.Visible = False
frsuspension.Enabled = False
frcapsule.Enabled = True
End Sub

Private Sub optsuspension_Click()
frsuspension.Visible = True
frcapsule.Visible = False
frsuspension.Enabled = True
frcapsule.Enabled = False
End Sub

Private Sub Timer2_Timer()
Dim today As Variant
today = Now
Label41.Caption = Format(today, "hh:mm:ss ampm")
Label8.Caption = Format(today, "mm/dd/yy")
End Sub
Sub PrescriptionType()
frcapsule.Visible = False
frsuspension.Visible = False
If Me.optcapsule.Value = True Then
    frcapsule.Visible = True
    frsuspension.Visible = False
Else
If Me.optsuspension.Value = True Then
    frcapsule.Visible = False
    frsuspension.Visible = True
End If
End If
End Sub
Private Sub optPrescriptionType_Click(Index As Integer)
PrescriptionType
intPrescriptionType = Index
End Sub
Sub PrescriptionInfo()
If optcapsule.Value = True Then
    strPrescription = "Prescription Type: Capsule/Tablets" & " , "
    strPrescription = strPrescription & "Qty:" & cboQTY2.Text & " , " ' &'' 'vbCrLf
    strPrescription = strPrescription & "Medicine:" & cboMedicine1.Text & " , " '& vbCrLf
    PrescriptionFin = strPrescription & "No Of Days to be taken:" & cboDays.Text '& vbCrLf & vbNewLine"
Else
    strPrescription = "Prescription Type: Suspension" & " , "  '& 'vbCrLf & vbNewLine
    strPrescription = strPrescription & "QTY: " & cboQTY.Text & " , "  '& ' vbCrLf & vbNewLine
    strPrescription = strPrescription & "Measurement: " & cbomeasurement.Text & " , " '& vbCrLf & vbNewLine
    strPrescription = strPrescription & "Medicine Type: " & cboMedicine.Text & " , " '& vbCrLf & vbNewLine
    PrescriptionFin = strPrescription & "X a Days: " & cboQTY1.Text '& vbCrLf & vbNewLine
End If

End Sub


