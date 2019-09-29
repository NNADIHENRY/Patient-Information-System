VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00404000&
   Caption         =   "Form6"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14145
   LinkTopic       =   "Form6"
   ScaleHeight     =   9120
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   8880
      TabIndex        =   45
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ADD"
      Height          =   495
      Left            =   7440
      TabIndex        =   43
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   11880
      TabIndex        =   42
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   8400
      TabIndex        =   41
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "individual.frx":0000
      Left            =   12600
      List            =   "individual.frx":0040
      TabIndex        =   37
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "individual.frx":00BC
      Left            =   11640
      List            =   "individual.frx":0120
      TabIndex        =   36
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "individual.frx":019A
      Left            =   10680
      List            =   "individual.frx":01C2
      TabIndex        =   35
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "individual.frx":0228
      Left            =   8400
      List            =   "individual.frx":0232
      TabIndex        =   34
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   10440
      TabIndex        =   33
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   1800
      TabIndex        =   32
      Top             =   6840
      Width           =   7335
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   1800
      TabIndex        =   31
      Top             =   6120
      Width           =   7335
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   1800
      TabIndex        =   30
      Top             =   5400
      Width           =   7335
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   1800
      TabIndex        =   29
      Top             =   4680
      Width           =   7335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   4
      Left            =   8040
      TabIndex        =   28
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   27
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   26
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   25
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   24
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1800
      TabIndex        =   23
      Top             =   3360
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "individual.frx":0244
      Left            =   1800
      List            =   "individual.frx":024E
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1800
      TabIndex        =   19
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   525
      Left            =   1800
      TabIndex        =   18
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   8400
      TabIndex        =   44
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DAY"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   11880
      TabIndex        =   40
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   12720
      TabIndex        =   39
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MONTH"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   38
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "WALK-IN"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "REMARKS"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAN OF ACTION"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DIAGNOSIS"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "P.E FINDINGS"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BR"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   13
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PR"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "WT"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   11
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AGE"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPERATURE"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RR"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPLAINTS"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FAMILY SERIAL NO."
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BARANGAY"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   11400
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   7200
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HEAD OF FAMILY"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INDIVIDUAL TREATMENT RECORD"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As Integer
Private Sub Command2_Click()
Dim remove As Integer
remove = lstName.ListIndex
If remove < 0 Then
MsgBox "No names is selected", vbInformation, "Error"
Else
answer = MsgBox("Are you sure you want to delete ", vbExclamation & vbCrLf & "the selected name?", vbCritical + vbYesNo, "Warning")
If answer = vbYes Then
If remove >= 0 Then
lstName.RemoveItem remove
txtName.SetFocus
MsgBox "Selected name was deleted", vbInformation, "Delete Confirm"

End If
End If
End If
End Sub

Private Sub Command3_Click()
answer = MsgBox("Are you sure to exit?", vbExclamation + vbYesNo, "Confirm")
If answer = vbYes Then
    Me.Hide
    frmMain.Show
End If
End Sub


Private Sub Command4_Click()
answer = MsgBox("Do you want to add record?", vbExclamation + vbYesNo, "Add Confirm")
If answer = vbYes Then
lstName.AddItem txtName.Text
txtName.Text = ""
txtName.SetFocus
CmdAdd.Enabled = False
End If

End Sub

Private Sub Timer1_Timer()

End Sub
