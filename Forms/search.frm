VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7770
   ClientLeft      =   105
   ClientTop       =   285
   ClientWidth     =   11070
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11070
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
      ItemData        =   "search.frx":0000
      Left            =   1440
      List            =   "search.frx":0010
      TabIndex        =   26
      Text            =   "barangay"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404000&
      Height          =   6495
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   4455
      Begin MSComctlLib.ListView PatientList 
         Height          =   5895
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   10398
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label17 
         BackColor       =   &H00404000&
         Caption         =   "Choose from the list below to display the patients information"
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
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "PRINT"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
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
      Left            =   6840
      TabIndex        =   12
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   5895
      Left            =   4680
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   4680
         Width           =   3375
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   18
         Top             =   4200
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   17
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   13
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   8
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   7
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT'S INFORMATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2040
         TabIndex        =   20
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Medical History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Types of Services"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "                         STATUS:"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
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
      Left            =   9120
      TabIndex        =   33
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404000&
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   10815
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "search.frx":0042
         Left            =   6360
         List            =   "search.frx":0052
         TabIndex        =   30
         Text            =   "Consultation"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404000&
         Caption         =   "Select Transaction Services"
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
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   3120
      TabIndex        =   32
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ictr As Integer
Dim str   As String
Private Sub Check2_Click()
Check1.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
End Sub

Private Sub Check3_Click()
Check1.Enabled = False
Check2.Enabled = False
Check4.Enabled = False
End Sub

Private Sub Check4_Click()
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
End Sub

Private Sub cmdclear_Click()
             Text5.Text = ""
             Text9.Text = ""
             Text8.Text = ""
             Text2.Text = ""
             Text4.Text = ""
             Text3.Text = ""
             Text12.Text = ""
             Text10.Text = ""
             Text11.Text = ""
End Sub

Private Sub cmdExit_Click()
Me.Hide
End Sub

Private Sub Command3_Click()
'Command6.Enabled = False
'cmdExit.Enabled = True
End Sub
Private Sub Command6_Click()
Me.Refresh
End Sub

Private Sub Lv_Initialize()
  Lv_SetView PatientList, "RECORDS LIST"
  With PatientList
    .ColumnHeaders(1).Width = 4200
   
  End With
End Sub

Private Sub cmdprint_Click()
        Set recordset = New ADODB.recordset
   Select Case Combo2.Text
        Case "Animal Bite"
          recordset.Open "select * from patient_animalbite_information where patient_Id='" & Text5.Text & "'", databaseconnection, 3, 3
          ' recordset.Open "select * from patient_animalbite_information ", databaseconnection, adOpenStatic, adLockOptimistic
         Case "Consultation"
            recordset.Open "select * from patient_consultation_information where patient_Id='" & Text5.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
         Case "Dressing"
            recordset.Open "select * from patient_dressing_information where patient_Id='" & Text5.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
         Case Else
            recordset.Open "select * from patient_immunization_information where patient_Id='" & Text5.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
        
   End Select
    Set DataReport1.DataSource = recordset
             DataReport1.Show vbModal
             recordset.Close
End Sub
Private Sub Combo1_Click()
   PatientList.Enabled = False
End Sub
Private Sub Combo2_Click()
   PatientList.Enabled = False
End Sub

Private Sub Command1_Click()

On Error Resume Next
  PatientList.Enabled = True
Set rs = New ADODB.recordset
   Select Case Combo2.Text
        Case "Consultation":
           rs.Open "select * from patient_consultation_information ", databaseconnection, 3, 3
           PatientList.ListItems.clear
           If Not rs.EOF Then
                    Select Case Combo1.Text
                            Case "barangay":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!barangay)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Last Name":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!lastname)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Case Number":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!CaseNo)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                        Case Else
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!patient_id)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                        End Select
        
                    
                    End If
        Case "Dressing":
           rs.Open "select * from patient_dressing_information ", databaseconnection, 3, 3
           PatientList.ListItems.clear
           If Not rs.EOF Then
                    Select Case Combo1.Text
                       Case "barangay":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!barangay)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Last Name":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!lastname)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Case Number":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!CaseNo)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                        Case Else
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!patient_id)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                    End Select
                End If
           Case "Animal Bite":
           rs.Open "select * from patient_animalbite_information ", databaseconnection, 3, 3
           PatientList.ListItems.clear
           If Not rs.EOF Then
                    Select Case Combo1.Text
                       Case "barangay":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!barangay)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Last Name":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!lastname)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Case Number":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!CaseNo)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                        Case Else
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!patient_id)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                    End Select
                End If
            Case Else
           rs.Open "select * from patient_immunization_information ", databaseconnection, 3, 3
           PatientList.ListItems.clear
           If Not rs.EOF Then
                    Select Case Combo1.Text
                        Case "barangay":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!barangay)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Last Name":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!lastname)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                         Case "Case Number":
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!CaseNo)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                        Case Else
                            For ictr = 1 To rs.RecordCount
                                Set gListItem = PatientList.ListItems.Add(, , rs!patient_id)
                                Set gListItem = Nothing
                                rs.MoveNext
                            Next
                    End Select
                End If
   End Select
End Sub

Private Sub Form_Load()
Lv_Initialize
End Sub
Private Sub PatientList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
  Set rs = New ADODB.recordset
  If Combo2.Text = "Consultation" Then
      If Combo1.Text = "barangay" Then
        rs.Open "Select * from patient_consultation_information where barangay='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Last Name" Then
        rs.Open "Select * from patient_consultation_information where lastname='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Case Number" Then
        rs.Open "Select * from patient_consultation_information where caseNo='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      Else
        rs.Open "Select * from patient_consultation_information where patient_id='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      End If
  ElseIf Combo2.Text = "Dressing" Then
      If Combo1.Text = "barangay" Then
        rs.Open "Select * from patient_dressing_information where barangay='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Last Name" Then
        rs.Open "Select * from patient_dressing_information where lastname='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Case Number" Then
        rs.Open "Select * from patient_dressing_information where caseNo='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      Else
        rs.Open "Select * from patient_dressing_information where patient_id='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      End If
  ElseIf Combo2.Text = "Animal Bite" Then
      If Combo1.Text = "barangay" Then
        rs.Open "Select * from patient_animalbite_information where barangay='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Last Name" Then
        rs.Open "Select * from patient_animalbite_information where lastname='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Case Number" Then
        rs.Open "Select * from patient_animalbite_information where caseNo='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      Else
        rs.Open "Select * from patient_animalbite_information where patient_id='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      End If
  Else
      If Combo1.Text = "barangay" Then
        rs.Open "Select * from patient_immunization_information where barangay='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Last Name" Then
        rs.Open "Select * from patient_immunization_information where lastname='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      ElseIf Combo1.Text = "Case Number" Then
        rs.Open "Select * from patient_immunization_information where caseNo='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      Else
        rs.Open "Select * from patient_immunization_information where patient_id='" & PatientList.SelectedItem & "'", databaseconnection, 3, 3
      End If
   End If   'Set gListItem = PatientList.ListItems.Item(1)
             Text5.Text = rs!patient_id
             Text9.Text = rs!CaseNo
             Text8.Text = rs!lastname & " " & rs!firstname
             Text2.Text = rs!barangay
             Text4.Text = rs!Birthdate
             Text3.Text = rs!Age
             Text12.Text = rs!barangay
             Text10.Text = rs!Gender
             Text11.Text = Combo2.Text
             
             
             
End Sub

