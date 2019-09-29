VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form chartImm 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ChartImmunization"
   ClientHeight    =   8895
   ClientLeft      =   690
   ClientTop       =   960
   ClientWidth     =   12120
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Form1.frx":0000
      Left            =   2160
      List            =   "Form1.frx":000A
      TabIndex        =   87
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
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
      Height          =   3015
      Left            =   600
      TabIndex        =   13
      Top             =   5040
      Width           =   10815
      Begin VB.TextBox Text72 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   78
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text71 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   77
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text70 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   76
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text69 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   75
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text68 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   74
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text67 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   73
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text66 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   72
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text65 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   71
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text64 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   70
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text63 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   69
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text62 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   68
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text61 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   67
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text60 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   66
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text59 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   65
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text58 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   64
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text57 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   63
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text56 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   62
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text55 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   61
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text54 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   60
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text53 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   59
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text52 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   58
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text51 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   57
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text50 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   56
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text49 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   55
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text48 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   54
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text47 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   53
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text46 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   52
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text45 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   51
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text44 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   50
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text43 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   49
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text42 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text41 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   46
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text39 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   45
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text38 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   44
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text37 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   43
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   42
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   41
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text34 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   40
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text33 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   39
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   38
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text31 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   37
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   35
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   34
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text26 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         MaxLength       =   4
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   21
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   18
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "DPT Drop-out  Percentage **"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "DPT1 Drop-out  Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   10800
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   10800
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Immunized DPT3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   10800
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Immunized DPT1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   10800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   10800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Fully Immunized Children for the Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   8280
      Width           =   1575
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4095
      Left            =   840
      OleObjectBlob   =   "Form1.frx":001A
      TabIndex        =   88
      Top             =   600
      Width           =   10935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   3600
      ScaleHeight     =   1275
      ScaleWidth      =   4155
      TabIndex        =   111
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Nov"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   110
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   11040
      TabIndex        =   109
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Sep"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   9000
      TabIndex        =   108
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Oct"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   9720
      TabIndex        =   107
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Jul"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   7800
      TabIndex        =   106
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Aug"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   8280
      TabIndex        =   105
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "May"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   6480
      TabIndex        =   104
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Jun"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   7200
      TabIndex        =   103
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Mar"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   102
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Apr"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   101
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Feb"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   100
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Jan"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   99
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Hepa B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   10320
      TabIndex        =   98
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H00FF00FF&
      Height          =   420
      Left            =   9720
      TabIndex        =   97
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "DPT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   9000
      TabIndex        =   96
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Measles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7320
      TabIndex        =   95
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Polio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6000
      TabIndex        =   94
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "BCG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4680
      TabIndex        =   93
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label30 
      BackColor       =   &H0080FFFF&
      Height          =   420
      Left            =   8400
      TabIndex        =   92
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label29 
      BackColor       =   &H0000FF00&
      Height          =   420
      Left            =   6720
      TabIndex        =   91
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H000040C0&
      Height          =   420
      Left            =   5400
      TabIndex        =   90
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label26 
      Height          =   420
      Left            =   4080
      TabIndex        =   89
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
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
      Left            =   1320
      TabIndex        =   86
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "X  100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   85
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "DPT1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   84
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "="
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
      Left            =   5280
      TabIndex        =   83
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "DPT1 - DPT3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   82
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "** DPT Drpo-out                 Percentage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   3960
      TabIndex        =   81
      Top             =   8280
      Width           =   2055
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "= PT1 - DPT3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   80
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "* DPT Drop-out               Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   720
      TabIndex        =   79
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mar."
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
      Left            =   5040
      TabIndex        =   12
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Jun."
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
      TabIndex        =   11
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Jul."
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
      Left            =   7680
      TabIndex        =   10
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Oct."
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
      Left            =   9720
      TabIndex        =   9
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nov."
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
      Left            =   10440
      TabIndex        =   8
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "May"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Aug."
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
      Left            =   8400
      TabIndex        =   6
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sep."
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
      Left            =   9120
      TabIndex        =   5
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Feb."
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
      TabIndex        =   4
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apr."
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
      Left            =   5640
      TabIndex        =   3
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jan."
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
      Left            =   3720
      TabIndex        =   2
      Top             =   5280
      Width           =   615
   End
End
Attribute VB_Name = "chartImm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X As Integer
    Dim a As Integer
    Dim irow   As Integer
    Dim rsTotall As Double
    Dim rsTotalw As Double
    Dim rsTotala As Double
    Dim rsTotalc As Double
    Dim rsTotalav As Double
    Dim rsTotalam As Double

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
Set recordset = New ADODB.recordset
        recordset.Open "Select * From immunization Where Year = '" & Combo1.Text & "'", databaseconnection, 3, 3
MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData(vbCFMetafile)
    SavePicture Picture1.Picture, App.Path & "\Image1.wmf"
    Set immuchart.DataSource = recordset
Set immuchart.Sections("section2").Controls.Item("Image1").Picture = LoadPicture((App.Path & "\Image1.wmf"))
    immuchart.Show
Clipboard.clear
End Sub

Private Sub Combo1_Click()
Dim i   As Integer
   'On Error Resume Next
Set recordset = New ADODB.recordset
        recordset.Open "Select * From immunization Where Year = '" & Combo1.Text & "'", databaseconnection, 3, 3
        If recordset.RecordCount > 0 Then
       
        ReDim ArrayChart(1 To recordset.RecordCount, 1 To 6) ' Array
        'Puting Records from Database to Array
        irow = 1
        For X = 1 To recordset.RecordCount
            ArrayChart(1, irow) = "Jan"
            ArrayChart(2, irow) = "Feb"
            ArrayChart(3, irow) = "Mar"
            ArrayChart(4, irow) = "Apr"
            ArrayChart(5, irow) = "May"
            ArrayChart(6, irow) = "Jun"
            ArrayChart(7, irow) = "Jul"
            ArrayChart(8, irow) = "Aug"
            ArrayChart(9, irow) = "Sep"
            ArrayChart(10, irow) = "Oct"
            ArrayChart(11, irow) = "Nov"
            ArrayChart(12, irow) = "Dec"
            
            ArrayChart(X, 2) = recordset!BCG
            ArrayChart(X, 3) = recordset!HepaB
            ArrayChart(X, 4) = recordset!Polio
            ArrayChart(X, 5) = recordset!Dpt
            ArrayChart(X, 6) = recordset!Measles
               'calculating the sum of all dressing
                   ' rsTotall = recordset!BCG + recordset!BCG
                    'rsTotalc = recordset!HepaB + recordset!HepaB
                    'rsTotala = recordset!Polio + recordset!Polio
                    'rsTotalav = recordset!Dpt + recordset!Dpt
                   ' rsTotalam = recordset!Measles + recordset!Measles
                    'rsTotalw = recordset!Wounds_Num + recordset!Wounds_Num
                recordset.MoveNext
                
        Next X
               
            MSChart1.ChartData = ArrayChart
            MSChart1.Refresh
        Else ' If no Record in Database, then Show an Error Msg and Exit the Sub
        MsgBox "No Data to Show on Chart!!!", vbCritical
        Exit Sub
'# Assigns our array to the MSChart control #
      
    End If
    i = 1
    For i = 1 To recordset.RecordCount
    Set recordset = New ADODB.recordset
        recordset.Open "Select * From immunization where Month='" & ArrayChart(i, irow) & "'", databaseconnection, 3, 3
        If recordset.RecordCount > 0 Then
          Select Case ArrayChart(i, irow)
              Case "Jan":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text1.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                    Text25.Text = Text1.Text
               Case "Feb":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text2.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
               Case "Mar":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text3.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
               Case "Apr":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text4.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
               Case "May":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text5.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Jun":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text6.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Jul":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text7.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Aug":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text8.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Sep":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text9.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Oct":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text10.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Nov":
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text11.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                Case "Dec":
                   
                    rsTotall = recordset!BCG
                    rsTotalc = recordset!HepaB
                    rsTotala = recordset!Polio
                    rsTotalav = recordset!Dpt
                    rsTotalam = recordset!Measles
                    Text12.Text = rsTotall + rsTotalc + rsTotala + rsTotalav + rsTotalam
                
          End Select
          End If
    Next i
        
        
        
    
End Sub

Private Sub Form_Load()
   Dim numSeries As Integer
   Dim icount     As Integer
            MSChart1.ToDefaults
    With MSChart1
          .chartType = VtChChartType2dBar
          '  Establish the number of items in the group
          numSeries = .Plot.SeriesCollection.Count
          ' Add a black line border of each of the shapes
        For icount = 1 To numSeries
          .Plot.SeriesCollection(icount).DataPoints(-1).EdgePen.VtColor.Set 0, 0, 0
        Next icount
          ' Turn off the background grids
              .Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
              .Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
              .Plot.Axis(VtChAxisIdY2).AxisGrid.MajorPen.Style = VtPenStyleNull
              .Plot.Wall.Pen.Style = VtPenStyleNull
          '  Define the background color to white
              .Backdrop.Fill.Brush.FillColor.Set 255, 255, 255
              .Backdrop.Fill.Style = VtFillStyleBrush
    End With
           
End Sub

