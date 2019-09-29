VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00404000&
   Caption         =   "Form7"
   ClientHeight    =   9030
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14025
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form7"
   ScaleHeight     =   9030
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   11040
      TabIndex        =   11
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   615
      Left            =   5760
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   7320
      TabIndex        =   9
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.PictureBox MSChart1 
      Height          =   5775
      Left            =   1560
      ScaleHeight     =   5715
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   2280
      Width           =   10455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ELIGIBLE POPULATION"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "HEALTH CENTER"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL POPULATION"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EPI MONITORING CHART"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrData(12, 1 To 12)
Dim answer As Integer

Private Sub Command1_Click()
Form4.Show
Me.Hide

End Sub

Private Sub Command2_Click()
answer = MsgBox("Do you want to exit now?", vbExclamation + vbYesNo, "Confirm")
If answer = vbYes Then
Form4.Show
Me.Hide
Else
MsgBox "Action canceled", vbInformation, "Confirm"

End If
End Sub

Private Sub Form_Load()
 
   arrData(1, 1) = "Jan"
   arrData(2, 1) = "Feb"
   arrData(3, 1) = "Mar"
   arrData(4, 1) = "Apr"
   arrData(5, 1) = "May"
   arrData(6, 1) = "Jun"
   arrData(7, 1) = "Jul"
   arrData(8, 1) = "Aug"
   arrData(9, 1) = "Sep"
   arrData(10, 1) = "Oct"
   arrData(11, 1) = "Nov"
   arrData(12, 1) = "Dec"
   

   arrData(1, 2) = 10 'dengue red
   arrData(2, 3) = 8 'tb red
   arrData(3, 4) = 7 ' nose bleed red
   arrData(4, 5) = 6 ' nose bleed red
   arrData(5, 6) = 5.5 ' nose bleed red
   arrData(6, 7) = 4 ' nose bleed red
   arrData(7, 8) = 4 ' nose bleed red
   arrData(8, 9) = 4 ' nose bleed red
   arrData(9, 10) = 4 ' nose bleed red
   arrData(10, 8) = 4 ' nose bleed red
   arrData(11, 9) = 4 ' nose bleed red
   arrData(12, 10) = 4 ' nose bleed red
   'MSChart1.ChartData = arrData

End Sub

