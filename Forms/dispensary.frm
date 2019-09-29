VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404000&
   Caption         =   "Form3"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13695
   LinkTopic       =   "Form3"
   ScaleHeight     =   9615
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   3360
      TabIndex        =   51
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   525
      Left            =   840
      TabIndex        =   50
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   6240
      TabIndex        =   49
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      Appearance      =   0  'Flat
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   48
      Top             =   9000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   47
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   12
      Left            =   6240
      TabIndex        =   46
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   11
      Left            =   6240
      TabIndex        =   45
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   10
      Left            =   840
      TabIndex        =   44
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   9
      Left            =   3360
      TabIndex        =   43
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   8
      Left            =   840
      TabIndex        =   42
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   7
      Left            =   3360
      TabIndex        =   41
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   40
      Top             =   8280
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   5
      Left            =   2640
      TabIndex        =   39
      Top             =   6480
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   38
      Top             =   7080
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   37
      Top             =   7680
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   36
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   35
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   34
      Top             =   2160
      Width           =   3855
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   30
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   29
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   28
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   0
      Left            =   9360
      TabIndex        =   23
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   0
      Left            =   10440
      TabIndex        =   22
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   0
      Left            =   11520
      TabIndex        =   21
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MONTH"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DAY"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   17
      Left            =   4320
      TabIndex        =   32
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   16
      Left            =   5640
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   15
      Left            =   10440
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MONTH"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      Left            =   9360
      TabIndex        =   26
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DAY"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   14
      Left            =   10680
      TabIndex        =   25
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   13
      Left            =   11640
      TabIndex        =   24
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Physician"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      TabIndex        =   20
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
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
      TabIndex        =   19
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan of Action"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      Left            =   240
      TabIndex        =   18
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Complaints"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ht"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   12
      Left            =   5640
      TabIndex        =   16
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   11
      Left            =   5640
      TabIndex        =   15
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BMI"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   10
      Left            =   5640
      TabIndex        =   14
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BP"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   13
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Temp"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   8
      Left            =   2640
      TabIndex        =   12
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Wt"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   10
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Head of the Family"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
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
      TabIndex        =   6
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MAIN DISPENSARY"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Republic of the Philippines"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CITY HEALTH DEPARTMENT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bacolod City"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Index           =   2
      Left            =   5640
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As Integer

Private Sub Command2_Click()
answer = MsgBox("Do you want to exit now?", vbExclamation + vbYesNo, "Confirm")
If answer = vbYes Then
Form4.Show
Me.Hide
Else
MsgBox "Action canceled", vbInformation, "Confirm"

End If
End Sub
