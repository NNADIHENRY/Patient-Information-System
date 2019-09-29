VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help Assistant"
   ClientHeight    =   4215
   ClientLeft      =   105
   ClientTop       =   285
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtguide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2775
      Left            =   3600
      TabIndex        =   7
      Top             =   720
      Width           =   5295
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00404000&
      Caption         =   "How to generate reports"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00404000&
      Caption         =   "How to register patients "
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
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00404000&
      Caption         =   "How to search patient Information"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00404000&
      Caption         =   "How to update patient"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404000&
      Caption         =   " About the System"
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      Height          =   2775
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM GUIDELINES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
        Me.Hide
End Sub

Private Sub Option1_Click()
        txtguide = "Bacolod City Health Out - Patients Information System" & " " & vbNewLine & "is a type of system that can support ideas and information" & vbNewLine & " accessing and traceable information and document." & vbNewLine & " The sytem can generate reports and it can provide an animal bite analyzer that help the medical professional to easily diagnose the condition of the patients."

End Sub

Private Sub Option2_Click()
        txtguide = "To Log In.Enter the username and password in the Log in form.Point the mouse pointer to PROCEED button then click.Main form will appear."

End Sub

Private Sub Option3_Click()
        txtguide = "First click File. Then click Transaction Files. The patients registration will appear then choose the services and update the data of the patients."
End Sub

Private Sub Option4_Click()
        txtguide = "Point the mouse in Query Services and click search patients record. The patient form record will appear."
End Sub

Private Sub Option5_Click()
        txtguide = "Point the mouse in File and then click File and go to Transaction Files."
End Sub
