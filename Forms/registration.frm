VERSION 5.00
Begin VB.Form frmRegistration 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7545
   ClientLeft      =   3825
   ClientTop       =   2700
   ClientWidth     =   8820
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdRegExit 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "Register"
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
         Left            =   2760
         TabIndex        =   14
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox txtRegAge 
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
         Left            =   2760
         TabIndex        =   13
         Top             =   4080
         Width           =   975
      End
      Begin VB.ComboBox cboRegGender 
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
         ItemData        =   "registration.frx":0000
         Left            =   2760
         List            =   "registration.frx":000A
         TabIndex        =   12
         Text            =   "Male"
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txtRegPosition 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox txtRegLN 
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
         Left            =   2760
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.ComboBox cboUsers 
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
         ItemData        =   "registration.frx":001C
         Left            =   2760
         List            =   "registration.frx":002F
         TabIndex        =   9
         Text            =   "Doctor"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtRegUsername 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox txtRegPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox txtRegCofirm 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txtRegFN 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtRegMN 
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
         Left            =   2760
         TabIndex        =   4
         Top             =   3120
         Width           =   3015
      End
      Begin VB.ComboBox cboRegMonth 
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
         ItemData        =   "registration.frx":007E
         Left            =   2760
         List            =   "registration.frx":00A6
         TabIndex        =   3
         Text            =   "Jan."
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboRegDay 
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
         ItemData        =   "registration.frx":00F1
         Left            =   3960
         List            =   "registration.frx":0152
         TabIndex        =   2
         Text            =   "1"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboRegYear 
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
         ItemData        =   "registration.frx":01C9
         Left            =   5160
         List            =   "registration.frx":0227
         TabIndex        =   1
         Text            =   "2019"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRATION FORM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
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
         Left            =   960
         TabIndex        =   29
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         Left            =   960
         TabIndex        =   28
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
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
         Left            =   960
         TabIndex        =   27
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   960
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Types of User"
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
         Height          =   495
         Left            =   960
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
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
         Left            =   960
         TabIndex        =   24
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   960
         TabIndex        =   23
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
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
         Height          =   495
         Left            =   840
         TabIndex        =   22
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Date"
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
         Height          =   495
         Left            =   960
         TabIndex        =   19
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label16 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim str     As String
Private Sub cmdRegExit_Click()
       Unload Me
End Sub

Private Sub cmdRegister_Click()
       Call AddNewRegistrationRecord
End Sub

Private Sub AddNewRegistrationRecord()
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM usertypes where usertype='" & cboUsers & "'", databaseconnection, 3, 3
           str = recordset!usertype_id
          
    'adding new record
    Set recordset = New ADODB.recordset
    recordset.Open "SELECT * FROM users", databaseconnection, 3, 3
    recordset.AddNew
           'trapping of errors
        If txtRegUsername.Text <> "" Or txtRegPassword <> "" Or txtRegCofirm.Text <> "" Then
              If txtRegPassword.Text = txtRegCofirm.Text Then
                    recordset("usertype_id") = str
                    recordset("DateofRegistration") = cboRegMonth.Text & "/" & cboRegDay.Text & "/" & cboRegMonth.Text
                    recordset("Lastname") = txtRegLN.Text
                    recordset("Firstname") = txtRegFN.Text
                    recordset("Middlename") = txtRegMN.Text
                    recordset("Position") = txtRegPosition.Text
                    recordset("Age") = txtRegAge.Text
                    recordset("Gender") = cboRegGender.Text
                    recordset("username") = txtRegUsername.Text
                    recordset("password") = txtRegPassword.Text
                    recordset.Update
                    MsgBox "Data HAS BEEN added", vbInformation, "REGISTRATION"
                    Call clear
              Else
                   MsgBox "CONFIRMATION OF PASSWORD IS INCORRECT", vbInformation, "REGISTRATION"
              End If
        Else
                 MsgBox "you must filled up the fields", vbInformation, "REGISTRATION"
        End If
End Sub


Public Sub clear()
                    txtRegAge.Text = ""
                    txtRegLN.Text = ""
                    txtRegFN.Text = ""
                    txtRegMN.Text = ""
                    txtRegPosition.Text = ""
                    txtRegAge.Text = ""
                    txtRegUsername.Text = ""
                    txtRegPassword.Text = ""
End Sub
