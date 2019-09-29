VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChartDressing 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8220
   ClientLeft      =   1890
   ClientTop       =   2490
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
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
      Left            =   6840
      TabIndex        =   21
      Top             =   6120
      Width           =   2175
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
      Left            =   10080
      TabIndex        =   16
      Top             =   7080
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
      Left            =   10080
      TabIndex        =   15
      Top             =   6120
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
      Left            =   10080
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
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
      Left            =   10080
      TabIndex        =   13
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3d"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   12
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2d"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   11
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2d1"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   7080
      Width           =   975
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
      ItemData        =   "frmChartDressing.frx":0000
      Left            =   9240
      List            =   "frmChartDressing.frx":0028
      TabIndex        =   8
      Text            =   "2009"
      Top             =   5520
      Width           =   2175
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
      Left            =   6840
      TabIndex        =   2
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   3615
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Avulsions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Contusions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "Abrasions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         Caption         =   "Lacerations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1920
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1920
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   495
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4815
      Left            =   480
      OleObjectBlob   =   "frmChartDressing.frx":0074
      TabIndex        =   7
      Top             =   240
      Width           =   10935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   480
      ScaleHeight     =   4635
      ScaleWidth      =   10635
      TabIndex        =   23
      Top             =   360
      Width           =   10695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL WOUNDS FOR THIS YEAR"
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "AVULSIONS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTUSIONS"
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
      Left            =   9000
      TabIndex        =   19
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ABRASIONS"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "LACERATION"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C0FFC0&
      Height          =   2055
      Left            =   4200
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0FFC0&
      Height          =   2775
      Left            =   240
      Top             =   5400
      Width           =   11295
   End
   Begin VB.Label Label5 
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
      Left            =   6360
      TabIndex        =   9
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0FFC0&
      FillColor       =   &H000000FF&
      Height          =   5175
      Left            =   240
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "frmChartDressing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim a As Integer
Dim irow   As Integer
Dim rsTotall As Double
Dim rsTotalw As Double
Dim rsTotala As Double
Dim rsTotalc As Double
Dim rsTotalav As Double

Private Sub cboYear_CLick()
 
  ' On Error Resume Next
Set recordset = New ADODB.recordset
        recordset.Open "Select * From Dressing_Chart Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3
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
            
            ArrayChart(X, 2) = recordset!Lacerations
            ArrayChart(X, 3) = recordset!Abrasions
            ArrayChart(X, 4) = recordset!Contusions
            ArrayChart(X, 5) = recordset!Avulsions
               'calculating the sum of all dressing
                    rsTotall = rsTotall + recordset!Lacerations
                    rsTotalav = rsTotalav + recordset!Contusions
                    rsTotala = rsTotala + recordset!Abrasions
                    rsTotalc = rsTotalc + recordset!Avulsions
                    rsTotalw = rsTotalw + recordset!Wounds_Num
                recordset.MoveNext
        Next X
               
            MSChart1.ChartData = ArrayChart
            MSChart1.Refresh
            Text1.Text = rsTotalc
            Text2.Text = rsTotalav
            Text3.Text = rsTotala
            Text4.Text = rsTotall
            Text5.Text = rsTotalw
        Else ' If no Record in Database, then Show an Error Msg and Exit the Sub
        MsgBox "No Data to Show on Chart!!!", vbCritical
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

Private Sub Command2_Click()
MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData(vbCFMetafile)
    SavePicture Picture1.Picture, App.Path & "\Image1.wmf"
Set dressingchart.DataSource = recordset
Set dressingchart.Sections("section2").Controls.Item("Image1").Picture = LoadPicture((App.Path & "\Image1.wmf"))
dressingchart.Show
Clipboard.clear
End Sub


Private Sub Form_Load()
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
    '-----------------------
    
   
End Sub
