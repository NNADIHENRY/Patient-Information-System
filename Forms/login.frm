VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   15
   ClientTop       =   -75
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "login.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Username"
      DataSource      =   "adoLog"
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
      Left            =   2400
      TabIndex        =   4
      Text            =   "admin"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "Password"
      DataSource      =   "adoLog"
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "admin"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "CLOSE"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   960
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "LOG IN"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      DrawMode        =   8  'Xor Pen
      FillColor       =   &H008080FF&
      Height          =   2535
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
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
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "09/18/2009"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "TIME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BACOLOD CITY HEALTH OUT-PATIENT INFORMATION SYSTEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   -840
      TabIndex        =   0
      Top             =   240
      Width           =   8535
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rs As ADODB.recordset
'Dim db As ADODB.Connection
Dim usertype As String
Dim typeofuser  As Integer

Private Sub cmdclear_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmRegistration.Show vbModal
End Sub

Private Sub cmdLogin_Click()

Set recordset = New ADODB.recordset
'recordset.Open "select * From users ", databaseconnection, 3, 3
usertype = "Select * from users where username = '" & Trim(txtUsername.Text) & "' and password = '" & Trim(txtPassword.Text) & "'"
              recordset.Open usertype, databaseconnection, 3, 3
                   
With frmMain
    
If (recordset(0) = 0) Then
                       MsgBox "Invalid password or username", vbInformation, "Access Denied"
                       txtUsername.Text = ""
                       txtPassword.Text = ""
                       txtUsername.SetFocus
                       recordset.Close
                       Exit Sub
Else
                    
                 
        typeofuser = recordset!usertype_id
        frmMain.Enabled = True
        frmMain.Show
                
            If typeofuser = "1" Then
                    .mnuFile.Enabled = True
                    .mnuConsultation_Form.Enabled = False
                    .Immunization_Form.Enabled = False
                    .Dressing_Form.Enabled = False
                    .Animal_Bite_Form.Enabled = False
                    .Database_Backup.Enabled = False
                    .Back_up_Files.Enabled = False
                    
'                    Select Case button.Index
'                        Case 2: Call Animal_Bite_Form_Click
'                        Case 3: Call mnuConsultation_Form_Click
'                        Case 4: Call Dressing_Form_Click
'                        Case 5: Call Immunization_Form_Click
'                        Case 6: Call Search_Patients_Record__Click
'                        Case 8: Call mnuHelp_Assistant_Click
'                        Case 9: Call Exit_Click
'                    End Select

                    With frmMain.Toolbar2
                        .Buttons(2).Enabled = False
                        .Buttons(3).Enabled = False
                        .Buttons(4).Enabled = False
                        .Buttons(5).Enabled = False
                    End With
            ElseIf typeofuser = "2" Then
                    .mnuFile.Enabled = True
                    .mnuConsultation_Form.Enabled = True
                    .Immunization_Form.Enabled = True
                    .Dressing_Form.Enabled = True
                    .Animal_Bite_Form.Enabled = True
                    .Database_Backup.Enabled = False
                    .Back_up_Files.Enabled = False
                    .mnuReport.Enabled = False
                    .mnuQuery_Services.Enabled = True
                     With frmMain.Toolbar2
                        .Buttons(2).Enabled = True
                        .Buttons(3).Enabled = True
                        .Buttons(4).Enabled = True
                        .Buttons(5).Enabled = True
                        .Buttons(6).Enabled = False
                    End With
            ElseIf typeofuser = "3" Then
                    .mnuConsultation_Form.Enabled = False
                    .Immunization_Form.Enabled = False
                    .Dressing_Form.Enabled = False
                    .Animal_Bite_Form.Enabled = False
                    .Database_Backup.Enabled = False
                    .Back_up_Files.Enabled = False
                    With frmMain.Toolbar2
                        .Buttons(2).Enabled = False
                        .Buttons(3).Enabled = False
                        .Buttons(4).Enabled = False
                        .Buttons(5).Enabled = False
                        .Buttons(6).Enabled = True
                    End With
            ElseIf typeofuser = "4" Then
                    .mnuConsultation_Form.Enabled = False
                    .Immunization_Form.Enabled = False
                    .Dressing_Form.Enabled = False
                    .Animal_Bite_Form.Enabled = False
                    .mnuReport.Enabled = False
                    .mnuQuery_Services.Enabled = False
                    .mnuHelp_Assistant.Enabled = False
                    With frmMain.Toolbar2
                        .Buttons(2).Enabled = False
                        .Buttons(3).Enabled = False
                        .Buttons(4).Enabled = False
                        .Buttons(5).Enabled = False
                        .Buttons(8).Enabled = False
                    End With
           Else
                    .mnuConsultation_Form.Enabled = True
                    .Immunization_Form.Enabled = True
                    .Dressing_Form.Enabled = True
                    .Animal_Bite_Form.Enabled = True
                    .mnuReport.Enabled = True
                    .mnuQuery_Services.Enabled = True
                    .mnuHelp_Assistant.Enabled = True
                    With frmMain.Toolbar2
                        .Buttons(2).Enabled = True
                        .Buttons(3).Enabled = True
                        .Buttons(4).Enabled = True
                        .Buttons(5).Enabled = True
                        .Buttons(8).Enabled = True
                    End With
            End If
End If

   Me.Hide
  End With
End Sub




Private Sub Timer1_Timer()
Dim today As Variant
today = Now
Label4.Caption = Format(today, "hh:mm:ss ampm")
Label5.Caption = Format(today, "mm/dd/yy")
End Sub

