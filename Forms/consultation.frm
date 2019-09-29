VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00404000&
   Caption         =   "Form9"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9450
   LinkTopic       =   "Form9"
   ScaleHeight     =   7020
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   22
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   21
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   20
      Top             =   6000
      Width           =   1575
   End
   Begin VB.ComboBox cYear 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   5040
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cDate 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   3840
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cMonth 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox cStatus 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   1815
   End
   Begin VB.ComboBox cGender 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.ComboBox cAge 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   4800
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Syndrome"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Name"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
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
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultation Form"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As Integer

Private Sub Command1_Click()
Text1.Text = ""
Text5.Text = ""
Text4.Text = ""
End Sub

Private Sub Command2_Click()
answer = MsgBox("Are you sure to exit now?", vbExclamation + vbYesNo, "Confirm")
If answer = vbYes Then
Form4.Show
Me.Hide
Else
MsgBox "Action canceled", vbInformation, "Confirm"

End If
End Sub

