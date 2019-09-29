VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConsultation 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8685
   ClientLeft      =   2115
   ClientTop       =   2370
   ClientWidth     =   11490
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frConsultationRecords 
      BackColor       =   &H00C9D5BB&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   11295
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1320
         TabIndex        =   141
         Top             =   5760
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   32143
      End
      Begin VB.TextBox txtage1 
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
         Index           =   3
         Left            =   1200
         TabIndex        =   132
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdCloseConsultationRecords 
         Caption         =   "Close Record"
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
         Left            =   8400
         TabIndex        =   39
         Top             =   7200
         Width           =   1695
      End
      Begin VB.TextBox txtid 
         Height          =   405
         Left            =   1800
         TabIndex        =   123
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Left            =   7680
         TabIndex        =   114
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox cboconstatus 
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
         ItemData        =   "newConsultation.frx":0000
         Left            =   3240
         List            =   "newConsultation.frx":0010
         TabIndex        =   111
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ComboBox cbocongender 
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
         ItemData        =   "newConsultation.frx":0037
         Left            =   1200
         List            =   "newConsultation.frx":0041
         TabIndex        =   110
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtConsultation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   20
         Left            =   4560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   81
         Top             =   6120
         Width           =   6495
      End
      Begin VB.CommandButton cmdUpdateConsultation 
         Caption         =   "Update Record"
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
         Left            =   6240
         TabIndex        =   40
         Top             =   7200
         Width           =   2175
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   0
         Left            =   1800
         TabIndex        =   38
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   1
         Left            =   1800
         TabIndex        =   37
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   2
         Left            =   1800
         TabIndex        =   36
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   5
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   6240
         Width           =   3255
      End
      Begin VB.TextBox txtConsultation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   6
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   6720
         Width           =   3255
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   8
         Left            =   5040
         TabIndex        =   33
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   9
         Left            =   6840
         TabIndex        =   32
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   10
         Left            =   5040
         TabIndex        =   31
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   11
         Left            =   7080
         TabIndex        =   30
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   12
         Left            =   5040
         TabIndex        =   29
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   13
         Left            =   7080
         TabIndex        =   28
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   14
         Left            =   5040
         TabIndex        =   27
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   15
         Left            =   7080
         TabIndex        =   26
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtConsultation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   16
         Left            =   4560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   4680
         Width           =   3735
      End
      Begin VB.TextBox txtConsultation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Index           =   17
         Left            =   8400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtConsultation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   18
         Left            =   8400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtConsultation 
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
         Index           =   19
         Left            =   1200
         TabIndex        =   22
         Top             =   7440
         Width           =   3255
      End
      Begin VB.CommandButton cmdconPrint 
         Caption         =   "Print Record"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   7200
         Width           =   1695
      End
      Begin MSComctlLib.ListView lstConsultationRecords 
         Height          =   2295
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   25
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Physician"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Patient"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "6"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "7"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "8"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Gender"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Civil Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Brgy"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "BDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Weight"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Height"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "BMI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "CR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "PR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "RR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "BP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Temperature"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Complaints"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Action"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "prescription"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   600
         TabIndex        =   131
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   124
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   115
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   113
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "lbs"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   112
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Prescription:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   4680
         TabIndex        =   82
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   41
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   60
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   59
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   58
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Ht:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   6480
         TabIndex        =   56
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Complaints:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4680
         TabIndex        =   54
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2520
         TabIndex        =   53
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Wt:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4560
         TabIndex        =   51
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "CR:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   6600
         TabIndex        =   50
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "BML:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   49
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "RR:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   6600
         TabIndex        =   48
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "PR:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4560
         TabIndex        =   47
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Temp:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   6360
         TabIndex        =   46
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "BP:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   4560
         TabIndex        =   45
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosis:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   8520
         TabIndex        =   44
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan of Action:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   8520
         TabIndex        =   43
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Physician:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   120
         TabIndex        =   42
         Top             =   7440
         Width           =   1335
      End
   End
   Begin VB.TextBox Text3 
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
      Height          =   405
      Left            =   2160
      TabIndex        =   140
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO "
      Height          =   495
      Left            =   3840
      TabIndex        =   139
      Top             =   720
      Width           =   975
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
      TabIndex        =   62
      Top             =   6240
      Width           =   11295
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   134
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddConsultationRecord 
         Caption         =   "Save"
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
         Left            =   3840
         TabIndex        =   128
         Top             =   1560
         Width           =   1815
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
         Left            =   5640
         TabIndex        =   127
         Top             =   1560
         Width           =   1815
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
         Left            =   7440
         TabIndex        =   126
         Top             =   1560
         Width           =   1815
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
         Height          =   495
         Left            =   9240
         TabIndex        =   125
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Frame frPrescription 
         BackColor       =   &H00404000&
         Caption         =   "Prescription"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Index           =   0
         Left            =   1920
         TabIndex        =   73
         Top             =   360
         Width           =   7935
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
            ItemData        =   "newConsultation.frx":0053
            Left            =   6240
            List            =   "newConsultation.frx":0069
            TabIndex        =   76
            Text            =   "1"
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
            ItemData        =   "newConsultation.frx":007F
            Left            =   4200
            List            =   "newConsultation.frx":00A1
            TabIndex        =   75
            Text            =   "Kremil-S"
            Top             =   480
            Width           =   1455
         End
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
            ItemData        =   "newConsultation.frx":010B
            Left            =   2160
            List            =   "newConsultation.frx":012D
            TabIndex        =   74
            Text            =   "1"
            Top             =   480
            Width           =   855
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
            TabIndex        =   80
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
            Index           =   0
            Left            =   5760
            TabIndex        =   79
            Top             =   480
            Width           =   615
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
            Index           =   0
            Left            =   3120
            TabIndex        =   78
            Top             =   480
            Width           =   975
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
            TabIndex        =   77
            Top             =   480
            Width           =   1815
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
         Height          =   1095
         Index           =   1
         Left            =   1920
         TabIndex        =   66
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
            ItemData        =   "newConsultation.frx":0150
            Left            =   1680
            List            =   "newConsultation.frx":0172
            TabIndex        =   70
            Text            =   "Qty."
            Top             =   480
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
            ItemData        =   "newConsultation.frx":01A2
            Left            =   2640
            List            =   "newConsultation.frx":01AC
            TabIndex        =   69
            Text            =   "Teaspoon"
            Top             =   480
            Width           =   1455
         End
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
            Index           =   1
            ItemData        =   "newConsultation.frx":01C6
            Left            =   4320
            List            =   "newConsultation.frx":01E8
            TabIndex        =   68
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
            Index           =   1
            ItemData        =   "newConsultation.frx":0218
            Left            =   6240
            List            =   "newConsultation.frx":023A
            TabIndex        =   67
            Text            =   "Medicine"
            Top             =   480
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
            TabIndex        =   72
            Top             =   480
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
            TabIndex        =   71
            Top             =   520
            Width           =   855
         End
      End
      Begin VB.OptionButton optPrescriptionType 
         BackColor       =   &H00404000&
         Caption         =   "Capsule/Tablets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
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
         TabIndex        =   64
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optPrescriptionType 
         BackColor       =   &H00404000&
         Caption         =   "Suspension"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Prescription Types"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2760
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   1215
      Left            =   120
      TabIndex        =   83
      Top             =   5160
      Width           =   11295
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   15
         Left            =   6720
         TabIndex        =   87
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   14
         Left            =   6720
         TabIndex        =   86
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   13
         Left            =   1560
         TabIndex        =   85
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cboConsultationDiagnosis 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "newConsultation.frx":02A4
         Left            =   1560
         List            =   "newConsultation.frx":02AE
         TabIndex        =   84
         Text            =   "Select Diagnosis"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Complaints"
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
         Index           =   4
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan of Action"
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
         Left            =   5280
         TabIndex        =   90
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Diagnosis"
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
         Left            =   240
         TabIndex        =   89
         Top             =   720
         Width           =   1335
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
         Index           =   3
         Left            =   5640
         TabIndex        =   88
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404000&
      Height          =   1215
      Left            =   120
      TabIndex        =   92
      Top             =   3960
      Width           =   11295
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   10
         Left            =   5400
         TabIndex        =   100
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "CR"
         DataSource      =   "adoKulai"
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
         Index           =   8
         Left            =   720
         TabIndex        =   99
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   9
         Left            =   3120
         TabIndex        =   98
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BP"
         DataSource      =   "adoKulai"
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
         Index           =   11
         Left            =   7800
         TabIndex        =   97
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   12
         Left            =   7800
         TabIndex        =   96
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   7
         Left            =   5400
         TabIndex        =   95
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
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
         Index           =   6
         Left            =   3120
         TabIndex        =   94
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Wt"
         DataSource      =   "adoKulai"
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
         Index           =   5
         Left            =   720
         TabIndex        =   93
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Temp"
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
         Left            =   7080
         TabIndex        =   108
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CR"
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
         Left            =   240
         TabIndex        =   107
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PR"
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
         Left            =   2640
         TabIndex        =   106
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Wt"
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
         Left            =   240
         TabIndex        =   105
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BP"
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
         Index           =   9
         Left            =   7200
         TabIndex        =   104
         Top             =   240
         Width           =   495
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
         Height          =   375
         Index           =   10
         Left            =   4800
         TabIndex        =   103
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RR"
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
         Index           =   11
         Left            =   4920
         TabIndex        =   102
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ht"
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
         Index           =   12
         Left            =   2640
         TabIndex        =   101
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "consultation date"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4935
      Begin VB.ComboBox cboConsultationYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Year"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":02C8
         Left            =   2640
         List            =   "newConsultation.frx":02F0
         TabIndex        =   5
         Text            =   "2009"
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboConsultationDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Day"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":033C
         Left            =   1440
         List            =   "newConsultation.frx":03A0
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboConsultationMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Month"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":041A
         Left            =   240
         List            =   "newConsultation.frx":0442
         TabIndex        =   3
         Text            =   "Jan."
         Top             =   720
         Width           =   1095
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
         Index           =   16
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   975
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
         Height          =   375
         Index           =   17
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   735
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
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   11295
      Begin VB.TextBox txtcaseno 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         TabIndex        =   138
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Barangay"
         DataSource      =   "adoanimal"
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
         Index           =   3
         Left            =   6360
         TabIndex        =   17
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Barangay"
         DataSource      =   "adoanimal"
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
         Index           =   4
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Last_Name"
         DataSource      =   "adoanimal"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "First_Name"
         DataSource      =   "adoanimal"
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
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtNewConsultation 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Middle_Name"
         DataSource      =   "adoanimal"
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
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label9 
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
         Height          =   375
         Left            =   7800
         TabIndex        =   117
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID Number"
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
         TabIndex        =   116
         Top             =   240
         Width           =   1815
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
         Index           =   14
         Left            =   5400
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
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
         Index           =   13
         Left            =   5400
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label43 
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
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         Left            =   840
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00404000&
      Height          =   1455
      Left            =   5040
      TabIndex        =   118
      Top             =   2640
      Width           =   6375
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   135
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   32143
      End
      Begin VB.ComboBox cboage 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Year"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":048D
         Left            =   5040
         List            =   "newConsultation.frx":0497
         TabIndex        =   133
         Text            =   "Year"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtage 
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
         Index           =   16
         Left            =   3840
         TabIndex        =   129
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboConsultationGender 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Year"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":04A8
         Left            =   360
         List            =   "newConsultation.frx":04B2
         TabIndex        =   120
         Text            =   "Male"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboConsultationCivilStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Birth_Year"
         DataSource      =   "adoKulai"
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
         ItemData        =   "newConsultation.frx":04C4
         Left            =   2040
         List            =   "newConsultation.frx":04CE
         TabIndex        =   119
         Text            =   "Single"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth date"
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
         Left            =   480
         TabIndex        =   136
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   130
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   122
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
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
         Index           =   18
         Left            =   1800
         TabIndex        =   121
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "09/23/09"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   137
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTATION FORM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7680
      TabIndex        =   109
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   " 00:00:00 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME:"
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
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPrescription As String, intPrescriptionType As Integer, arrPrescription() As String
Dim str As String

Private Sub cmdAddConsultationRecord_Click()
    PrescriptionInfo
     ' checking of value
    no_null_val
'----------------------------------------------------------------------------------------'
  ' adding record to te consultation
   
    '------------------------------------------------------------------------------'

End Sub
Private Sub cmdClearConsultation_Click()
    'clearing of text
        ClearConsultationRecords
End Sub

Private Sub cmdCloseConsultationRecords_Click()
     frConsultationRecords.Visible = False
End Sub
Private Sub cmdconPrint_Click()
On Error Resume Next
Set recordset2 = New ADODB.recordset
        If (txtid.Text = "") Then
                MsgBox "Please choose patient record  .", vbCritical, "No data to be printed"
                Exit Sub
        Else:
                str = "SELECT * FROM patient_consultation_information WHERE patient_id ='" & txtid.Text & "'"
                recordset2.Open str, databaseconnection, 3, 3
                'recordset2.Close
                Set rpt_consultation.DataSource = recordset2
                rpt_consultation.Show vbModal
        End If
End Sub
Private Sub cmdExit_Click()
'frmMain.Show
'frmLogin.Show
    Me.Hide
    Unload Me
End Sub

Private Sub cmdgo_Click()
cmdAddConsultationRecord.Enabled = True
 Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_consultation_information where patient_id='" & UCase(Text3.Text) & "' order by caseNo", databaseconnection, 3, 3
    If Text3.Text = "" Then
       MsgBox "field is empty", vbInformation
    
   ElseIf recordset.RecordCount > 0 Then
         txtCaseNo = recordset("caseNo") + 1
         txtNewConsultation(0) = recordset("lastname")
         txtNewConsultation(1) = recordset("firstname")
         txtNewConsultation(2) = recordset("middlename")
        cboConsultationGender.Text = recordset("gender")
         cboConsultationCivilStatus.Text = recordset("civil_status")
         txtNewConsultation(3) = recordset("barangay")
        txtNewConsultation(4) = recordset("address")
         DTPicker1.Value = recordset("birthdate")
         txtage(16).Text = recordset("Age")
    Else
        MsgBox " No records Found", vbInformation
  End If
  Frame5.Enabled = False
  Frame2.Enabled = False
  Frame3.Enabled = False
End Sub

Private Sub cmdUpdateConsultation_Click()
        UpdateConsultationRecord (consultationID)
End Sub

Private Sub cmdViewConsultationRecords_Click()
    frConsultationRecords.Visible = True
    LoadConsultationRecords
End Sub

Private Sub Command1_Click()
ClearConsultationRecords
cmdgo.Enabled = False
cmdAddConsultationRecord.Enabled = True
Command1.Enabled = False
Call AutoID
  Frame5.Enabled = True
  Frame2.Enabled = True
  Frame3.Enabled = True
  Autocaseno

End Sub

Private Sub Form_Load()
LoadConsultationRecords
frPrescription(0).Visible = True
frPrescription(1).Visible = False
frConsultationRecords.Height = 11415
strPrescription = ""
intPrescriptionType = 0

End Sub
Private Sub AddNewConsultationRecord()
Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM patient_consultation_information", databaseconnection, adOpenDynamic, adLockPessimistic
recordset.AddNew
recordset("consultation_date") = Date
recordset("consultation_time") = Time
recordset("caseNo") = txtCaseNo.Text
recordset("lastname") = txtNewConsultation(0)
recordset("firstname") = txtNewConsultation(1)
recordset("middlename") = txtNewConsultation(2)
recordset("gender") = cboConsultationGender.Text
recordset("civil_status") = cboConsultationCivilStatus.Text
recordset("barangay") = txtNewConsultation(3)
recordset("address") = txtNewConsultation(4)
recordset("birthdate") = DTPicker1.Value
recordset("weight") = txtNewConsultation(5)
recordset("height") = txtNewConsultation(6)
recordset("bmi") = txtNewConsultation(7)
recordset("cr") = txtNewConsultation(8)
recordset("pr") = txtNewConsultation(9)
recordset("rr") = txtNewConsultation(10)
recordset("bp") = txtNewConsultation(11)
recordset("temperature") = txtNewConsultation(12)
recordset("complaints") = txtNewConsultation(13)
recordset("diagnosis") = cboConsultationDiagnosis.Text
recordset("plan_of_action") = txtNewConsultation(14)
recordset("physician") = txtNewConsultation(15)
recordset("prescription") = strPrescription
recordset("Age") = CInt(txtage(16).Text)
recordset("patient_id") = Text3.Text
recordset.Update

 Set recordset = New ADODB.recordset
        recordset.Open "Select * From Consultation Where Month = '" & cboConsultationMonth.Text & "' and Year = '" & cboConsultationYear.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
        
    If recordset.RecordCount > 0 Then
        If cboage.Text = "Month" Then
            If txtage(16).Text >= 1 And txtage(16).Text <= 12 Then
             recordset!oneToOneYear = recordset!oneToOneYear + 1
            Else
               MsgBox "Please Enter a Value from 1 to 12. or choose the Year option for age more than 12 months", vbInformation
               txtage(16).Text = ""
            End If
            
        Else
            If txtage(16).Text >= 2 And txtage(16).Text <= 10 Then
            recordset!twoToTenyears = recordset!twoToTenyears + 1
            ElseIf txtage(16).Text >= 11 And txtage(16).Text <= 30 Then
            recordset!ElevenToThirty = recordset!ElevenToThirty + 1
            Else
            recordset!thirtyOneUp = recordset!thirtyOneUp + 1
            End If
        End If
            recordset!totalpatient = recordset!totalpatient + 1
     Else
        recordset.AddNew
        If cboage.Text = "Month" Then
            If txtage(16).Text >= 1 And txtage(16).Text <= 12 Then
             recordset!oneToOneYear = recordset!oneToOneYear + 1
            Else
               MsgBox "Please Enter a Value from 1 to 12. or choose the Year option for age more than 12 months", vbInformation
               txtage(16).Text = ""
            End If
            
        Else
            If txtage(16).Text >= 2 And txtage(16).Text <= 10 Then
            recordset!twoToTenyears = recordset!twoToTenyears + 1
            ElseIf txtage(16).Text >= 11 And txtage(16).Text <= 30 Then
            recordset!ElevenToThirty = recordset!ElevenToThirty + 1
            Else
            recordset!thirtyOneUp = recordset!thirtyOneUp + 1
            End If
        End If
            recordset!Year = cboConsultationYear.Text
            recordset!Month = cboConsultationMonth.Text
            recordset!totalpatient = recordset!totalpatient + 1
     End If
        recordset.Update
        recordset.Close
         'clearing of text
        ClearConsultationRecords
End Sub
Private Sub no_null_val()

If txtNewConsultation(0) = "" Or txtNewConsultation(1) = "" Or txtNewConsultation(2) = "" Or _
txtNewConsultation(3) = "" Or txtNewConsultation(4) = "" Or txtNewConsultation(5) = "" _
Or txtNewConsultation(6) = "" Or txtNewConsultation(8) = "" Or txtNewConsultation(9) = "" Or _
txtNewConsultation(11) = "" Or txtNewConsultation(12) = "" Or txtNewConsultation(13) = "" _
Or txtNewConsultation(14) = "" Or txtNewConsultation(15) = "" Or cboConsultationGender.Text = "" _
Or cboConsultationCivilStatus.Text = "" Or cboConsultationMonth.Text = "" Or cboConsultationDate.Text = "" Or _
cboConsultationYear.Text = "" Or cboConsultationDiagnosis.Text = "" Or txtCaseNo.Text = "" Then
MsgBox "Please fill in missing details", vbInformation, "Missing Info"

Else: AddNewConsultationRecord
    
MsgBox "Patient data has been Successfully added.", vbInformation, "Data Added"
End If

End Sub
Private Sub LoadConsultationRecords()
    Dim intctr As Integer
    intctr = 0
    lstConsultationRecords.ListItems.clear
    Set recordset2 = New ADODB.recordset
    recordset2.Open "SELECT * FROM patient_consultation_information", databaseconnection, adOpenDynamic, adLockPessimistic
    If Not recordset2.BOF Then
        Do Until recordset2.EOF
            Set a = lstConsultationRecords.ListItems.Add(, , recordset2(0))
               ' a.SubItems(25) = recordset2("patient_id")
                a.SubItems(1) = recordset2(1)
                a.SubItems(2) = recordset2(2)
                a.SubItems(3) = recordset2(22)
                a.SubItems(4) = recordset2(3) & " " & recordset2(4) & ", " & recordset2(5)
                a.SubItems(5) = recordset2(3)
                a.SubItems(6) = recordset2(4)
                a.SubItems(7) = recordset2(5)
                a.SubItems(8) = recordset2(6)
                a.SubItems(9) = recordset2(7)
                a.SubItems(10) = recordset2(8)
                a.SubItems(11) = recordset2(9)
                a.SubItems(12) = recordset2("Birthdate")
                a.SubItems(13) = recordset2(11)
                a.SubItems(14) = recordset2(12)
                a.SubItems(15) = recordset2(13)
                a.SubItems(16) = recordset2(14)
                a.SubItems(17) = recordset2(15)
                a.SubItems(18) = recordset2(16)
                a.SubItems(19) = recordset2(17)
                a.SubItems(20) = recordset2(18)
                a.SubItems(21) = recordset2(19)
                a.SubItems(22) = recordset2(20)
                a.SubItems(23) = recordset2(21)
                a.SubItems(24) = recordset2("Age")
                ReDim Preserve arrPrescription(intctr) As String
                If Not IsNull(recordset2(23)) Then
                    arrPrescription(intctr) = recordset2(23)
                Else
                    arrPrescription(intctr) = ""
                End If
               recordset2.MoveNext
               intctr = intctr + 1
        Loop
    End If

End Sub

Private Function AutoID()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_consultation_information Order By patient_id desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            Text3.Text = "CO-0001"
        Else
            Text3.Text = "CO-000" + Format(Right(recordset!patient_id, 4) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        Text3.Locked = True
End Function

Private Function Autocaseno()
 
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM patient_consultation_information  where patient_Id='" & Text3.Text & "' Order By caseNo desc", databaseconnection, 3, 2
     '  "select * from PO Order By POID DESC"
       If recordset.RecordCount = 0 Then
            txtCaseNo.Text = "01"
        Else
            txtCaseNo.Text = "0" + Format(Right(recordset!patient_id, 2) + 1)
        End If
        recordset.Close
        Set recordset = Nothing
        txtCaseNo.Locked = True
End Function







Private Sub lstConsultationRecords_ItemClick(ByVal Item As MSComctlLib.ListItem)

Set recordset2 = New ADODB.recordset
    recordset2.Open "SELECT * FROM patient_consultation_information where patient_id='" & lstConsultationRecords.SelectedItem & " '", databaseconnection, 3, 3
 
'txtConsultation(0) = Item
    consultationID = Item
    txtid = recordset2("patient_id")
    txtConsultation(0) = Item.ListSubItems(5)
    txtConsultation(1) = Item.ListSubItems(6)
    txtConsultation(2) = Item.ListSubItems(7)
    cbocongender = Item.ListSubItems(8)
    cboconstatus = Item.ListSubItems(9)
    txtConsultation(5) = Item.ListSubItems(10)
    txtConsultation(6) = Item.ListSubItems(11)
    cboconyear = Item.ListSubItems(12)
    DTPicker2.Value = Item.ListSubItems(12)
    txtConsultation(8) = Item.ListSubItems(13)
    txtConsultation(9) = Item.ListSubItems(14)
    txtConsultation(10) = Item.ListSubItems(16)
    txtConsultation(11) = Item.ListSubItems(17)
    txtConsultation(12) = Item.ListSubItems(15)
    txtConsultation(13) = Item.ListSubItems(18)
    txtConsultation(14) = Item.ListSubItems(19)
    txtConsultation(15) = Item.ListSubItems(20)
    txtConsultation(16) = Item.ListSubItems(21)
    txtConsultation(17) = Item.ListSubItems(22)
    txtConsultation(18) = Item.ListSubItems(23)
    txtConsultation(19) = Item.ListSubItems(3)
    txtage1(3).Text = Item.ListSubItems(24)
    txtConsultation(20) = arrPrescription(Item.Index - 1)
End Sub

Private Sub UpdateConsultationRecord(ByVal id)
    Set recordset2 = New ADODB.recordset
    'recordset2.Open "SELECT * FROM patient_consultation_information WHERE lastname='" & txtConsultation(0).Text & "' ", databaseconnection, adOpenDynamic, adLockPessimistic
    recordset2.Open "SELECT * FROM patient_consultation_information WHERE patient_id='" & txtid.Text & "' ", databaseconnection, adOpenDynamic, adLockPessimistic
  If Not recordset2.BOF Then
                                        
        recordset2(3) = txtConsultation(0)
        recordset2(4) = txtConsultation(1)
        recordset2(5) = txtConsultation(2)
        recordset2(6) = cbocongender
        recordset2(7) = cboconstatus
        recordset2(8) = txtConsultation(5)
        recordset2(9) = txtConsultation(6)
        recordset2(10) = cboconyear
        recordset2(11) = txtConsultation(8)
        recordset2(12) = txtConsultation(9)
        recordset2(13) = txtConsultation(10)
        recordset2(17) = txtConsultation(14)
        recordset2(18) = txtConsultation(15)
        recordset2(19) = txtConsultation(16)
        recordset2(20) = txtConsultation(17)
        recordset2(21) = txtConsultation(18)
        recordset2(22) = txtConsultation(19)
        recordset2(23) = txtConsultation(20)
        recordset2.Update
        ClearConsultationEntries
        LoadConsultationRecords
   End If
End Sub

Private Sub ClearConsultationEntries()
For i = 0 To 2
    txtConsultation(i) = ""
Next i
For i = 5 To 6
    txtConsultation(i) = ""
Next i
For i = 8 To 19
    txtConsultation(i) = ""
Next i
txtid.Text = ""
Text3.Text = ""
End Sub
Private Sub ClearConsultationRecords()
AutoID
For i = 0 To 15
txtNewConsultation(i) = ""
Next i
cboConsultationMonth.Text = ""
cboConsultationDate.Text = ""
cboConsultationYear.Text = ""
cboConsultationGender.Text = ""
cboConsultationCivilStatus.Text = ""
cboConsultationDiagnosis.Text = ""
optPrescriptionType(0).Value = False
optPrescriptionType(1).Value = False
cboMedicineMeasurementQTY.Text = "QTY"
cboQTY(0).Text = "QTY"
cboQTY(1).Text = "QTY"
cboMedicineDays.Text = "No."
cboMedicine(0).Text = "Medicine"
cboMedicine(1).Text = "Medicine"
cboMedicineMeasurement.Text = "Teaspoon"
frPrescription(0).Visible = False
frPrescription(1).Visible = False
strPrescription = ""
End Sub

Private Sub optPrescriptionType_Click(Index As Integer)
PrescriptionType
intPrescriptionType = Index
End Sub

Private Sub Timer4_Timer()
Dim today As Variant
today = Now
Label34.Caption = Format(today, "hh:mm:ss ampm")
Label11.Caption = Format(today, "mm/dd/yy")
End Sub
Sub PrescriptionType()
frPrescription(0).Visible = False
frPrescription(1).Visible = False
If Me.optPrescriptionType(0).Value = True Then
    frPrescription(0).Visible = True
    frPrescription(1).Visible = False
Else
If Me.optPrescriptionType(1).Value = True Then
    frPrescription(0).Visible = False
    frPrescription(1).Visible = True
End If
End If
End Sub
Sub PrescriptionInfo()
If intPrescriptionType = 0 Then
    strPrescription = strPrescription & "Prescription Type: Capsule/Tablets" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "Prescription: " & cboMedicine(0).Text & vbCrLf & vbNewLine
    strPrescription = strPrescription & "Prescription Qty: " & cboQTY(0).Text & vbCrLf & vbNewLine
    strPrescription = strPrescription & "No Of Days to be taken: " & cboMedicineDays.Text & vbCrLf & vbNewLine
Else
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
    strPrescription = strPrescription & "" & vbCrLf & vbNewLine
End If
End Sub

Private Sub txtConsultation_Change(Index As Integer)
txtConsultation(12).Text = Val(txtConsultation(8).Text) * Val(txtConsultation(9).Text)
txtConsultation(13).Text = Val(txtConsultation(10).Text) * Val(txtConsultation(11).Text)
End Sub

Private Sub txtid_Change()
txtid.Locked = True
End Sub

Private Sub txtNewConsultation_Change(Index As Integer)
txtNewConsultation(7).Text = Val(txtNewConsultation(5).Text) * Val(txtNewConsultation(6).Text)
txtNewConsultation(10).Text = Val(txtNewConsultation(8).Text) * Val(txtNewConsultation(9).Text)
End Sub
