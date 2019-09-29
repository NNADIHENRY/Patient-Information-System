VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChartConsultation 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8415
   ClientLeft      =   2085
   ClientTop       =   2685
   ClientWidth     =   11160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Report"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   7080
      Width           =   1815
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
      Left            =   6000
      TabIndex        =   14
      Top             =   7080
      Width           =   1815
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "chart_syndrome.frx":0000
      Left            =   8760
      List            =   "chart_syndrome.frx":001C
      TabIndex        =   13
      Text            =   "2009"
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2d1"
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2d"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3d"
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   10
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      TabIndex        =   9
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6480
      TabIndex        =   5
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "Legend"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   3375
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "1 month - 1 year old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "2 yrs. - 10 yrs. old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "11yrs. - 30 yrs. old"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "31yrs. old and up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   1800
         Width           =   1815
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4695
      Left            =   240
      OleObjectBlob   =   "chart_syndrome.frx":0050
      TabIndex        =   16
      Top             =   240
      Width           =   10575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   360
      ScaleHeight     =   4635
      ScaleWidth      =   10395
      TabIndex        =   23
      Top             =   240
      Width           =   10455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0FFC0&
      Height          =   2055
      Left            =   3600
      Top             =   5760
      Width           =   7455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H00C0FFC0&
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Chart for the year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5760
      TabIndex        =   22
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "11 to  30 yrs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1 to 12 Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "31 to up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "2 to 10 yrs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NUMBER OF CONSULTATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   6000
      Width           =   2895
   End
End
Attribute VB_Name = "frmChartConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim answer As Integer
Dim X As Integer
Dim a As Integer
Dim irow   As Integer
Dim rsTotall As Double
Dim rsTotalw As Double
Dim rsTotala As Double
Dim rsTotalc As Double
Dim rsTotalav As Double

Private Sub cboYear_CLick()
   On Error Resume Next
Set recordset = New ADODB.recordset
        recordset.Open "Select * From Consultation Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3
        If recordset.RecordCount > 0 Then
       
        ReDim ArrayChart(1 To recordset.RecordCount, 1 To 5) ' Array
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
            ArrayChart(9, irow) = "Sept"
            ArrayChart(10, irow) = "Oct"
            ArrayChart(11, irow) = "Nov"
            ArrayChart(12, irow) = "Dec"
            
            ArrayChart(X, 2) = recordset!oneToOneYear
            ArrayChart(X, 3) = recordset!twoToTenyears
            ArrayChart(X, 4) = recordset!ElevenToThirty
            ArrayChart(X, 5) = recordset!thirtyOneUp
               'calculating the sum of all dressing
                    rsTotall = rsTotall + recordset!oneToOneYear
                    rsTotalc = rsTotalc + recordset!twoToTenyears
                    rsTotala = rsTotala + recordset!ElevenToThirty
                    rsTotalav = rsTotalav + recordset!thirtyOneUp
                    'rsTotalw = recordset!Wounds_Num + recordset!Wounds_Num
                recordset.MoveNext
                
        Next X
               
            MSChart1.ChartData = ArrayChart
            MSChart1.Refresh
            Text4.Text = rsTotala
            Text1.Text = rsTotalav
            Text3.Text = rsTotall
            Text2.Text = rsTotalc
            Text5.Text = rsTotall + rsTotalc + rsTotala + rsTotalav

        Else ' If no Record in Database, then Show an Error Msg and Exit the Sub
               MsgBox "No Data to Show on Chart!!!", vbInformation, " PATIENT INFORMATION SYSTEM"
               Exit Sub
'# Assigns our array to the MSChart control #

    End If
                    
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
With MSChart1
 chartStr = Command1(Index).Caption
  Select Case chartStr
    Case Is = "2d1"
      .chartType = VtChChartType2dArea
      .Stacking = True
    Case Is = "2d"
      .chartType = VtChChartType2dBar
      .Stacking = False
    Case Is = "3d"
      .chartType = VtChSeriesType3dBar
      .Plot.Projection = VtProjectionTypeOblique
      .Stacking = True
  End Select
End With
End Sub

Private Sub Command3_Click()
Set recordset = New ADODB.recordset
        recordset.Open "Select * From Consultation Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3

MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData(vbCFMetafile)
    SavePicture Picture1.Picture, App.Path & "\Image1.wmf"
Set consult.DataSource = recordset
Set consult.Sections("section2").Controls.Item("Image1").Picture = LoadPicture((App.Path & "\Image1.wmf"))
consult.Show
Clipboard.clear

End Sub

Private Sub Form_Load()
'   If recordset.RecordCount > 0 Then
'       Do Until recordset.EOF
'           cboYear.AddItem
'        recordset.MoveNext
'        Loop
'   End If
  
            Dim numSeries As Integer
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




