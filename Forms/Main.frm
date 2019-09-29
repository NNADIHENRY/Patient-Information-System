VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9915
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   4210816
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   64
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EE044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EE376
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EE782
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EEB2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EEF1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2EF350
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2F7F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2FCC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":30110B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   1005
      ButtonWidth     =   1852
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Animal Bite"
            Object.ToolTipText     =   "Animal Bite"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultation"
            Object.ToolTipText     =   "Consultation"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dressing"
            Object.ToolTipText     =   "Dressing"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Immunization"
            Object.ToolTipText     =   "Immunization"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Object.ToolTipText     =   "Search"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "About"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Main.frx":301DEF
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6120
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTransaction_Files 
         Caption         =   "&Transaction Files"
         Begin VB.Menu mnuConsultation_Form 
            Caption         =   "Consultation Form"
            Shortcut        =   {F2}
         End
         Begin VB.Menu Immunization_Form 
            Caption         =   "&Immunization Form"
            Shortcut        =   {F3}
         End
         Begin VB.Menu Dressing_Form 
            Caption         =   "&Dressing Form"
            Shortcut        =   {F4}
         End
         Begin VB.Menu Animal_Bite_Form 
            Caption         =   "Animal Bite Form"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu mnuNewUserAccount 
         Caption         =   "New User Account"
         Shortcut        =   ^N
      End
      Begin VB.Menu Database_Backup 
         Caption         =   "Database Back Up"
         Shortcut        =   ^D
      End
      Begin VB.Menu Back_up_Files 
         Caption         =   "Back Up Files"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "mnuLogOut"
         Shortcut        =   ^L
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu Statistical_Reports 
         Caption         =   "Statistical Reports"
         Begin VB.Menu Animal_Bite 
            Caption         =   "Animal Bite"
            Shortcut        =   {F6}
         End
         Begin VB.Menu Immunization 
            Caption         =   "Immunization"
            Shortcut        =   {F7}
         End
         Begin VB.Menu Consultation 
            Caption         =   "Consultation"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuDressing 
            Caption         =   "Dressing"
            Shortcut        =   {F9}
         End
      End
   End
   Begin VB.Menu mnuQuery_Services 
      Caption         =   "Query Services"
      Begin VB.Menu Search_Patients_Record_ 
         Caption         =   "Search Patient's Record"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuHelp_Assistant 
      Caption         =   "Help Assistant"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Animal_Bite_Form_Click()
frmAnimalBite.Show
End Sub
Private Sub Animal_Bite_Click()
frmChartAnimalBite.Show
End Sub



Private Sub Consultation_Click()
frmChartConsultation.Show
End Sub
Private Sub Consultation_Form_Click()
Form8.Show
Me.Hide
End Sub
Private Sub Dressing_Form_Click()
frmDressing.Show
End Sub
Private Sub Exit_Click()
 
Unload Me
End Sub

Private Sub Immunization_Click()
chartImm.Show
End Sub
Private Sub Patients_Medical_Form_Click()
p.Show
End Sub
Private Sub Immunization_Form_Click()
frmImmunization.Show
End Sub

Private Sub mnuLogOut_Click()
  frmMain.Enabled = False
  frmLogin.Show
End Sub

Private Sub Toolbar2_ButtonClick(ByVal button As MSComctlLib.button)
Select Case button.Index
    Case 2: Call Animal_Bite_Form_Click
    Case 3: Call mnuConsultation_Form_Click
    Case 4: Call Dressing_Form_Click
    Case 5: Call Immunization_Form_Click
    Case 6: Call Search_Patients_Record__Click
    Case 8: Call mnuHelp_Assistant_Click
    Case 9: Call Exit_Click
End Select
End Sub
Private Sub MDIForm_Unload(cancel As Integer)
On Error Resume Next
If MsgBox("Are You Sure you want to Quit ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
    Unload frmMain
Else
cancel = True
End If
End Sub

Private Sub MDIForm_Load()
With Toolbar2
Set .ImageList = ImageList1
.Buttons(2).Image = 9
.Buttons(3).Image = 6
.Buttons(4).Image = 7
.Buttons(5).Image = 8
.Buttons(6).Image = 1
.Buttons(8).Image = 5
.Buttons(9).Image = 4


End With
End Sub

Private Sub mnuConsultation_Form_Click()
frmConsultation.Show
End Sub
Private Sub mnuDressing_Click()
frmChartDressing.Show
End Sub
Private Sub mnuHelp_Assistant_Click()
frmHelp.Show
End Sub
Private Sub mnuNewUserAccount_Click()
frmRegistration.Show
End Sub
Private Sub Search_Patients_Record__Click()
frmSearch.Show
End Sub

