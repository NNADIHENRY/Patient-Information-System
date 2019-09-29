VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChartAnimalBite 
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmChartAnimalBite"
   ClientHeight    =   8205
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   11190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11190
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
      Left            =   6480
      TabIndex        =   15
      Top             =   6120
      Width           =   2055
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
      TabIndex        =   14
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
      Left            =   9720
      TabIndex        =   13
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
      Left            =   9720
      TabIndex        =   12
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
      Left            =   9720
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmChartAnimalBite.frx":0000
      Left            =   8880
      List            =   "frmChartAnimalBite.frx":0022
      TabIndex        =   10
      Text            =   "2000"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   2775
      Begin VB.Label Label3 
         BackColor       =   &H00404000&
         Caption         =   "Monkey Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         Caption         =   "Dog Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Top             =   1680
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Top             =   720
         Width           =   375
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Top             =   1200
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   240
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         Caption         =   "Snake Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         Caption         =   "Insect Bite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4080
      TabIndex        =   1
      Top             =   7320
      Width           =   1935
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
      Left            =   6120
      TabIndex        =   0
      Top             =   7320
      Width           =   1575
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4335
      Left            =   840
      OleObjectBlob   =   "frmChartAnimalBite.frx":0062
      TabIndex        =   7
      Top             =   840
      Width           =   9975
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   840
      ScaleHeight     =   3915
      ScaleWidth      =   9795
      TabIndex        =   22
      Top             =   1080
      Width           =   9855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NUMBER OF BITE PER YEAR"
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
      TabIndex        =   21
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Monkey Bite"
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
      TabIndex        =   20
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Snake Bite"
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
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Dog Bite"
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
      Left            =   8760
      TabIndex        =   18
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insect Bite"
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
      Left            =   8760
      TabIndex        =   17
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0FFC0&
      Height          =   2055
      Left            =   3600
      Top             =   6000
      Width           =   7455
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0FFC0&
      Height          =   4815
      Left            =   480
      Top             =   480
      Width           =   10575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      Caption         =   "This Year Animal Bite Chart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "N o        o  f            B i   t  e s"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   4095
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "frmChartAnimalBite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim icount  As Integer
Dim rsTotall As Double
Dim rsTotalw As Double
Dim rsTotala As Double
Dim rsTotalc As Double
Dim rsTotalav As Double
Dim irow As Integer

Private Sub cboYear_CLick()
Dim X As Integer
Dim a As Integer

    Set recordset = New ADODB.recordset
    'recordset.Open "SELECT * FROM animalbite_chart", databaseconnection, 3, 3
    recordset.Open "Select * From animalbite_chart Where Year = '" & cboYear.Text & "'", databaseconnection, adOpenStatic, adLockOptimistic
     
    If recordset.RecordCount = 0 Then
        MsgBox "No Data to Show on Chart!!!", vbCritical, "MSChart Demo": Exit Sub
    Else
        ReDim ArrayChart(1 To recordset.RecordCount, 1 To 5)
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
                 ArrayChart(X, 2) = recordset!DogBite
                 ArrayChart(X, 3) = recordset!MOnkeyBite
                 ArrayChart(X, 4) = recordset!InsectBite
                 ArrayChart(X, 5) = recordset!SnakeBite
                ' ArrayChart(X, 6) = recordset!Year
                'calculating the sum of all dressing
                    rsTotall = rsTotall + recordset!DogBite
                    rsTotalc = rsTotalc + recordset!MOnkeyBite
                    rsTotala = rsTotala + recordset!InsectBite
                    rsTotalav = rsTotalav + recordset!SnakeBite
                    'Text5.Text = rsTotalav + recordset!Wounds_Num
                 recordset.MoveNext
                 Next X

            MSChart1.ChartData = ArrayChart
            MSChart1.Refresh
                    Text4.Text = rsTotalc
                    Text1.Text = rsTotalav
                    Text3.Text = rsTotall
                    Text2.Text = rsTotala
                    Text5.Text = rsTotall + rsTotalc + rsTotala + rsTotalav
    End If
                    
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub Command2_Click()
        Dim X As Integer
        Dim a As Integer
        
        
            Set recordset = New ADODB.recordset
            'recordset.Open "SELECT * FROM animalbite_chart", databaseconnection, 3, 3
            recordset.Open "Select * From animalbite_chart Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3
            
            recordset.Filter = ""
             
            If recordset.RecordCount = 0 Then
                MsgBox "No Data to Show on Chart!!!", vbCritical, "MSChart Demo": Exit Sub
            Else
                ReDim ArrayChart(1 To recordset.RecordCount, 1 To 5)
                
                         For X = 1 To recordset.RecordCount
                         ArrayChart(X, 1) = recordset!Month
                         ArrayChart(X, 2) = recordset!DogBite
                         ArrayChart(X, 3) = recordset!MOnkeyBite
                         ArrayChart(X, 4) = recordset!InsectBite
                         ArrayChart(X, 5) = recordset!SnakeBite
                        ' ArrayChart(X, 6) = recordset!Year
                         
                         recordset.MoveNext
                         Next X
        
                    MSChart1.ChartData = ArrayChart
                    MSChart1.Refresh
            End If
End Sub

Private Sub Command1_Click()
Set recordset = New ADODB.recordset
        recordset.Open "Select * From animalbite_chart Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3

MSChart1.EditCopy
    Picture1.Picture = Clipboard.GetData(vbCFMetafile)
    SavePicture Picture1.Picture, App.Path & "\Image1.wmf"
Set animal.DataSource = recordset
Set animal.Sections("section2").Controls.Item("Image1").Picture = LoadPicture((App.Path & "\Image1.wmf"))
animal.Show
Clipboard.clear
End Sub

Private Sub Form_Load()
Dim X As Integer
Dim a As Integer

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
    Set recordset = New ADODB.recordset
    'recordset.Open "SELECT * FROM animalbite_chart", databaseconnection, 3, 3
    recordset.Open "Select * From animalbite_chart Where Year = '" & cboYear.Text & "'", databaseconnection, 3, 3
    
    recordset.Filter = ""
     
    If recordset.RecordCount = 0 Then
        MsgBox "No Data to Show on Chart!!!", vbCritical, "MSChart Demo": Exit Sub
    Else
        ReDim ArrayChart(1 To recordset.RecordCount, 1 To 5)
        
                 For X = 1 To recordset.RecordCount
                 ArrayChart(X, 1) = recordset!Month
                 ArrayChart(X, 2) = recordset!DogBite
                 ArrayChart(X, 3) = recordset!MOnkeyBite
                 ArrayChart(X, 4) = recordset!InsectBite
                 ArrayChart(X, 5) = recordset!SnakeBite
                ' ArrayChart(X, 6) = recordset!Year
                 
                 recordset.MoveNext
                 Next X

            MSChart1.ChartData = ArrayChart
            MSChart1.Refresh
    End If
End Sub


