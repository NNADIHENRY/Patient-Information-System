VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImmunization 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8385
   ClientLeft      =   105
   ClientTop       =   285
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2880
      Top             =   600
   End
   Begin VB.Frame frimmunization 
      BackColor       =   &H00C9D5BB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   11175
      Begin VB.TextBox txtbarangay1 
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
         Left            =   1800
         TabIndex        =   79
         Top             =   6480
         Width           =   3015
      End
      Begin VB.TextBox txtage1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   75
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdCloserecordsImmunization 
         Caption         =   "Close Record"
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
         Left            =   8760
         TabIndex        =   46
         Top             =   6960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6720
         TabIndex        =   74
         Top             =   5040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88997889
         CurrentDate     =   32143
      End
      Begin MSComctlLib.ListView lstimmunization 
         Height          =   3615
         Left            =   240
         TabIndex        =   53
         Top             =   120
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6376
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.ComboBox cboblood 
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
         ItemData        =   "immunization_form.frx":0000
         Left            =   6720
         List            =   "immunization_form.frx":000D
         TabIndex        =   69
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CommandButton cmdsaverec 
         Caption         =   "Update Record"
         Enabled         =   0   'False
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
         Left            =   7080
         TabIndex        =   68
         Top             =   6960
         Width           =   1695
      End
      Begin VB.ComboBox cbovaccine 
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
         ItemData        =   "immunization_form.frx":001B
         Left            =   6720
         List            =   "immunization_form.frx":002E
         TabIndex        =   66
         Top             =   5520
         Width           =   1695
      End
      Begin VB.ComboBox cbogender 
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
         ItemData        =   "immunization_form.frx":0054
         Left            =   1800
         List            =   "immunization_form.frx":005E
         TabIndex        =   64
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtlast 
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
         Left            =   1800
         TabIndex        =   52
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox txtfirst 
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
         Left            =   1800
         TabIndex        =   51
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox txtmiddle 
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
         Left            =   1800
         TabIndex        =   50
         Top             =   5040
         Width           =   3015
      End
      Begin VB.TextBox txtweight 
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
         Left            =   6720
         TabIndex        =   49
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox txtrate 
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
         Left            =   1800
         TabIndex        =   48
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txtadd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   6960
         Width           =   3015
      End
      Begin VB.CommandButton cmdprintrecord 
         Caption         =   "Print Record"
         Enabled         =   0   'False
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
         Left            =   5400
         TabIndex        =   45
         Top             =   6960
         Width           =   1695
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
         Height          =   495
         Left            =   840
         TabIndex        =   80
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   6240
         TabIndex        =   76
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "lbs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   67
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   2640
         TabIndex        =   65
         Top             =   6120
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         TabIndex        =   63
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         TabIndex        =   62
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
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
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label17 
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
         Left            =   960
         TabIndex        =   60
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label45 
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
         Height          =   375
         Left            =   5880
         TabIndex        =   59
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
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
         TabIndex        =   58
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Fatal Heart           Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   57
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label58 
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
         Left            =   840
         TabIndex        =   56
         Top             =   6960
         Width           =   1095
      End
      Begin VB.Label Label61 
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
         Left            =   5640
         TabIndex        =   55
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Vaccine Type"
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
         Left            =   5280
         TabIndex        =   54
         Top             =   5640
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   6480
      Width           =   11175
      Begin VB.CommandButton cmdExit2 
         Caption         =   "Exit"
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
         Left            =   8880
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdviewimmunizationrecords 
         Caption         =   "View Records"
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
         Left            =   6720
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdimmunizationclear 
         Cancel          =   -1  'True
         Caption         =   "Clear"
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
         Left            =   4560
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdAddImmunizationRecord 
         Caption         =   "Save"
         Enabled         =   0   'False
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
         Left            =   2400
         TabIndex        =   40
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "New"
         Default         =   -1  'True
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
         Left            =   240
         TabIndex        =   81
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Caption         =   "Patient's Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   11175
      Begin VB.Frame Frame5 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   615
         Left            =   1560
         TabIndex        =   82
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtpatientid 
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
            Left            =   120
            TabIndex        =   84
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton cmdS 
            Caption         =   "G O "
            Height          =   375
            Left            =   2280
            TabIndex        =   83
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txtbarangay 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Middle_Name"
         DataSource      =   "adoimmuniz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   77
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox txtage 
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
         Left            =   5280
         TabIndex        =   72
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtcasenum 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtimmunizationmn 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Middle_Name"
         DataSource      =   "adoimmuniz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtimmunizationfn 
         BackColor       =   &H00FFFFC0&
         DataField       =   "First_Name"
         DataSource      =   "adoimmuniz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtimmunizationln 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Last_Name"
         DataSource      =   "adoimmuniz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtimmunizationaddress 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Address"
         DataSource      =   "adoimmuniz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4080
         TabIndex        =   15
         Top             =   2400
         Width           =   6975
         Begin VB.TextBox txtimmunizationfatalheartrate 
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
            Height          =   390
            Left            =   4680
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboimmunizationBloodtype 
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
            ItemData        =   "immunization_form.frx":0070
            Left            =   4680
            List            =   "immunization_form.frx":0080
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtimmunizationweight 
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
            Height          =   390
            Left            =   1200
            TabIndex        =   17
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboimmunizationGender 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Gender"
            DataSource      =   "adoimmuniz"
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
            ItemData        =   "immunization_form.frx":0091
            Left            =   1200
            List            =   "immunization_form.frx":009B
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Fatal Heart Rate"
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
            Left            =   2640
            TabIndex        =   23
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Blood Type"
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
            Left            =   3240
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label18 
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
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label16 
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
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404000&
         Caption         =   "Immunization Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   3855
         Begin VB.ComboBox cboimmunizationdate 
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
            ItemData        =   "immunization_form.frx":00AD
            Left            =   1200
            List            =   "immunization_form.frx":0111
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cboimmunizationyear 
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
            ItemData        =   "immunization_form.frx":018B
            Left            =   2280
            List            =   "immunization_form.frx":01AD
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cboimmunizationmonth 
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
            ItemData        =   "immunization_form.frx":01ED
            Left            =   120
            List            =   "immunization_form.frx":0215
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label19 
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
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   1440
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label21 
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
            Left            =   2520
            TabIndex        =   12
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404000&
         Caption         =   "Types of Vaccine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   4080
         TabIndex        =   2
         Top             =   3840
         Width           =   6975
         Begin VB.OptionButton OptHEPAB 
            BackColor       =   &H00404000&
            Caption         =   "HEPA B"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   4800
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optDPT 
            BackColor       =   &H00404000&
            Caption         =   "DPT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   3840
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optMEASLES 
            BackColor       =   &H00404000&
            Caption         =   "MEASLES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optPOLIO 
            BackColor       =   &H00404000&
            Caption         =   "POLIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optBCG 
            BackColor       =   &H00404000&
            Caption         =   "BCG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   78
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label22 
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
         Left            =   4080
         TabIndex        =   73
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   240
         TabIndex        =   71
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Case Number"
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
         Left            =   7560
         TabIndex        =   42
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's ID"
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
         TabIndex        =   41
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label48 
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
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label46 
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
         Left            =   -1200
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label13 
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
         Left            =   4080
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IMMUNIZATION FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1080
      TabIndex        =   70
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Immunization:"
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
      Left            =   360
      TabIndex        =   35
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "09/23/09"
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
      Left            =   3360
      TabIndex        =   34
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label36 
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
      Height          =   375
      Left            =   4680
      TabIndex        =   33
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label37 
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
      Left            =   5520
      TabIndex        =   32
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmImmunization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As String
Dim intctr As Integer
Dim patientid  As String


Private Function AutoID()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_immunization_information Order By patient_id desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            txtpatientID.Text = "IM-0001"
        Else
            txtpatientID.Text = "IM-000" + Format(Right(recordset!patient_id, 4) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtpatientID.Locked = True
End Function
Private Function Autocaseno()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_immunization_information Order By caseNo desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            txtcasenum.Text = "IM-0001"
        Else
            txtcasenum.Text = "IM-000" + Format(Right(recordset!patient_id, 4) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtcasenum.Locked = True
End Function


Private Sub cmdAddImmunizationRecord_Click()

Set recordset = New ADODB.recordset
recordset.Open "select * From patient_immunization_information", databaseconnection, 3, 3
'txtpatientid.Enabled = False

With recordset
    recordset.AddNew
       ' recordset!immunization_date = Label37.Caption
        'recordset!immunization_time = Label10.Caption
        recordset!patient_id = txtpatientID.Text
        recordset!patient_id = txtage.Text
        recordset!CaseNo = txtcasenum.Text
        recordset!lastname = txtimmunizationln.Text
        recordset!firstname = txtimmunizationfn.Text
        recordset!middlename = txtimmunizationmn.Text
        recordset!Gender = cboimmunizationGender.Text
        recordset!Address = txtimmunizationaddress.Text
        recordset!Birthdate = cboimmunizationmonth.Text + " " + cboimmunizationdate.Text + "," + cboimmunizationyear.Text
        recordset!Weight = txtimmunizationweight.Text
        recordset!bloodtype = cboimmunizationBloodtype.Text
        recordset!fatalheartrate = txtimmunizationfatalheartrate.Text
        recordset!vaccine = answer
        recordset!Age = txtage.Text
        recordset!barangay = txtbarangay.Text
           With Me
If txtimmunizationln.Text = "" Or txtimmunizationfn.Text = "" Or txtimmunizationmn.Text = "" _
                    Or cboimmunizationGender.Text = "" Or txtimmunizationaddress.Text = "" _
                    Or txtimmunizationfatalheartrate.Text = "" Or txtimmunizationweight.Text = "" _
                    Or cboimmunizationBloodtype.Text = "" Or txtbarangay.Text = "" _
                    Then
                        MsgBox "Required fields must be complete", vbInformation, "Information"
         
Else:
                recordset.Update
    ' ================== data to be added in the chart =====================
      Set recordset = New ADODB.recordset
      recordset.Open "Select * From immunization Where Month = '" & cboimmunizationmonth.Text & "' and Year = '" & cboimmunizationyear.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
        
      If recordset.RecordCount > 0 Then
           recordset!numberofpatient = recordset!numberofpatient + 1
            If answer = "BCG" Then
                recordset!BCG = recordset!BCG + 1
            ElseIf answer = "POLIO" Then
                recordset!Polio = recordset!Polio + 1
            ElseIf anwswer = "DPT" Then
                recordset!Dpt = recordset!Dpt + 1
            ElseIf answer = "HEPA B" Then
                recordset!HepaB = recordset!HepaB + 1
            Else
                recordset!Measles = recordset!Measles + 1
            End If
     Else
            recordset.AddNew
            recordset!numberofpatient = recordset!numberofpatient + 1
            If answer = "BCG" Then
                recordset!BCG = recordset!BCG + 1
            ElseIf answer = "POLIO" Then
                recordset!Polio = recordset!Polio + 1
            ElseIf anwswer = "DPT" Then
                recordset!Dpt = recordset!Dpt + 1
            ElseIf answer = "HEPA B" Then
                recordset!HepaB = recordset!HepaB + 1
            Else
                recordset!Measles = recordset!Measles + 1
            End If
        recordset!Year = cboimmunizationyear.Text
        recordset!Month = cboimmunizationmonth.Text
        recordset.Update
        recordset.Close
                MsgBox "Data has been successfully Saved!"
       End If
       End If
       End With
    End With
        
End Sub

Private Sub cmdCloserecordsImmunization_Click()
    Frame3.Visible = True
    Frame6.Visible = True
    frimmunization.Visible = False
End Sub


Private Sub cmdExit2_Click()
Unload Me
End Sub

Private Sub cmdimmunizationclear_Click()
        
        txtcasenum.Text = ""
        txtimmunizationln.Text = ""
        txtimmunizationfn.Text = ""
        txtimmunizationmn.Text = ""
        cboimmunizationGender.Text = ""
        txtimmunizationaddress.Text = ""
        cboimmunizationmonth.Text = ""
        cboimmunizationdate.Text = ""
        cboimmunizationyear.Text = ""
        txtimmunizationweight.Text = ""
        cboimmunizationBloodtype.Text = ""
        txtimmunizationfatalheartrate.Text = ""
        txtpatientID.Text = ""
        txtage.Text = ""
        optBCG.Value = False
        optPOLIO.Value = False
        optMEASLES.Value = False
        optDPT.Value = False
        OptHEPAB.Value = False
        txtbarangay.Text = ""
        
        

End Sub

Private Sub cmdNew_Click()
      cmdAddImmunizationRecord.Enabled = True
      cmdS.Enabled = False
      AutoID
End Sub

Private Sub cmdprintrecord_Click()
   Set recordset = New ADODB.recordset
     recordset.Open "select * From patient_immunization_information where patient_id = '" & lstimmunization.SelectedItem & "'", databaseconnection, 3, 3
        Set immun.DataSource = recordset
            immun.Show
End Sub

Private Sub cmdS_Click()
     
    Set recordset = New ADODB.recordset
    recordset.Open "select * From patient_immunization_information where patient_id = '" & txtpatientID.Text & "'", databaseconnection, 3, 3
    If recordset.RecordCount > 0 Then
    txtcasenum.Text = ""
        txtimmunizationln.Text = recordset!lastname
        txtimmunizationfn.Text = recordset!firstname
        txtimmunizationmn.Text = recordset!middlename
        cboimmunizationGender.Text = recordset!Gender
        txtimmunizationaddress.Text = recordset!lastname
        cboimmunizationmonth.Text = recordset!lastname
        cboimmunizationdate.Text = recordset!lastname
        cboimmunizationyear.Text = recordset!lastname
        txtimmunizationweight.Text = recordset!lastname
        cboimmunizationBloodtype.Text = recordset!lastname
        txtimmunizationfatalheartrate.Text = recordset!lastname
        txtage.Text = recordset!Age
        
Else
 MsgBox "Patient ID not found", vbInformation
 
End If
 cmdAddImmunizationRecord.Enabled = True
End Sub

Private Sub cmdsaverec_Click()
Set recordset = New ADODB.recordset
recordset.Open "select * From patient_immunization_information where patient_id = '" & lstimmunization.SelectedItem & "'", databaseconnection, 3, 3
  If recordset.RecordCount > 0 Then
With recordset
    .Clone
        recordset!CaseNo = txtcasenum.Text
        recordset!lastname = txtlast.Text
        recordset!firstname = txtfirst.Text
        recordset!middlename = txtmiddle.Text
        recordset!Gender = cbogender.Text
        recordset!Address = txtadd.Text
        recordset!Birthdate = DTPicker1.Value
        recordset!bloodtype = cboblood.Text
        recordset!fatalheartrate = txtrate.Text
        recordset!Weight = txtWeight.Text
        recordset!vaccine = answer
        recordset!Age = txtage1.Text
      .Update
      LoadImmunizationRecords
      End With
 End If
      With Me
        .cmdprintrecord.Enabled = True
        .cmdsaverec.Enabled = False
    End With
End Sub
      


Private Sub ClearImmunizationEntries()
       
        txtlast.Text = ""
        txtfirst.Text = ""
        txtmiddle.Text = ""
        cbogender.Text = ""
        txtadd.Text = ""
        cboMonth.Text = ""
        cboDay.Text = ""
        cboYear.Text = ""
        txtWeight.Text = ""
        cboblood.Text = ""
        txtrate.Text = ""
        optBCG.Value = False
        optPOLIO.Value = False
        optMEASLES.Value = False
        optDPT.Value = False
        OptHEPAB.Value = False
        txtage.Text = ""
        
End Sub

Private Sub cmdviewimmunizationrecords_Click()
Frame3.Visible = False
Frame6.Visible = False
frimmunization.Visible = True
LoadImmunizationRecords
'Lv_Initialize
End Sub

Private Sub Command1_Click()
'answer = MsgBox("Are you sure to exit?", vbExclamation + vbYesNo, "Confirm")
'If answer = vbYes Then
'frmMain.Show
'Me.Hide
'Else
'MsgBox "Action canceled", vbInformation, "Confirm"

'End If
End Sub

Private Sub LoadImmunizationRecords()
On Error Resume Next
Dim str As String
Dim intimr As Integer
   intimr = 0
    lstimmunization.ListItems.clear
    Set recordset2 = New ADODB.recordset
    str = "SELECT * FROM patient_immunization_information"
    recordset2.Open str, databaseconnection, 3, 3
    If Not recordset2.BOF Then
       Do Until recordset2.EOF
            Set gListItem = lstimmunization.ListItems.Add(, , recordset2!patient_id)
                gListItem.SubItems(1) = recordset2!immunization_date
                gListItem.SubItems(2) = recordset2!CaseNo
                gListItem.SubItems(3) = recordset2!lastname
                gListItem.SubItems(4) = recordset2!firstname
                gListItem.SubItems(5) = recordset2!middlename
                gListItem.SubItems(6) = recordset2!Gender
                gListItem.SubItems(7) = recordset2!Address
                gListItem.SubItems(8) = recordset2!Birthdate
                gListItem.SubItems(9) = recordset2!bloodtype
                gListItem.SubItems(10) = recordset2!Weight
                gListItem.SubItems(11) = recordset2!fatalheartrate
                gListItem.SubItems(12) = recordset2!vaccine
                gListItem.SubItems(13) = recordset2!Age
                gListItem.SubItems(14) = recordset2!barangay
                 recordset2.MoveNext
              intimr = intimr + 1
        Loop
   End If

End Sub

Private Sub Command2_Click()
  
End Sub

Private Sub Form_Load()
    Lv_Initialize
End Sub
Private Sub lstimmunization_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error Resume Next

patientid = Item
'txtpatientid(1) = Item.ListSubItems(1)
txtcasenum.Text = Item.ListSubItems(2)
txtlast.Text = Item.ListSubItems(3)
txtfirst.Text = Item.ListSubItems(4)
txtmiddle.Text = Item.ListSubItems(5)
cbogender.Text = Item.ListSubItems(6)
txtadd.Text = Item.ListSubItems(7)
DTPicker1.Value = Item.ListSubItems(8)
cboblood.Text = Item.ListSubItems(9)
txtWeight.Text = Item.ListSubItems(10)
txtrate.Text = Item.ListSubItems(11)
cbovaccine.Text = Item.ListSubItems(12)
txtage1.Text = Item.ListSubItems(13)
txtbarangay1.Text = Item.ListSubItems(14)
    With Me
        .cmdprintrecord.Enabled = True
        .cmdsaverec.Enabled = True
    End With
End Sub

Private Sub Lv_Initialize()
On Error Resume Next

Lv_SetView lstimmunization, "Patient_ID|Immunization Date|Case No.|Last Name|First Name|Middle Name|Gender|Address|Birth Date|Blood Type|Weight|Fatal Heart Rate|Vaccine|Age |Barangay"
With lstimmunization

    .ColumnHeaders(1).Width = 1800
    .ColumnHeaders(2).Width = 2500
    .ColumnHeaders(3).Width = 2000
    .ColumnHeaders(4).Width = 2000
    .ColumnHeaders(5).Width = 2000
    .ColumnHeaders(6).Width = 1800
    .ColumnHeaders(7).Width = 2000
    .ColumnHeaders(8).Width = 1800
    .ColumnHeaders(9).Width = 1500
    .ColumnHeaders(10).Width = 1500
    .ColumnHeaders(11).Width = 1500
    .ColumnHeaders(12).Width = 1500
    .ColumnHeaders(13).Width = 1500
    .ColumnHeaders(13).Width = 1500
    
End With

End Sub


Private Sub optBCG_Click()
answer = "BCG"
End Sub

Private Sub optDPT_Click()
answer = "DPT"
End Sub

Private Sub OptHEPAB_Click()
answer = "HEPA B"
End Sub

Private Sub optMEASLES_Click()
answer = "MEASLES"
End Sub

Private Sub optPOLIO_Click()
answer = "POLIO"
End Sub

Private Sub Timer3_Timer()
Dim today As Variant
today = Now
Label37.Caption = Format(today, "hh:mm:ss ampm")
Label10.Caption = Format(today, "mm/dd/yy")

End Sub

