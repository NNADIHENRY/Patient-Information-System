Attribute VB_Name = "modListView"
Option Explicit



Public Sub Lv_SetView(lv As MSComctlLib.ListView, s_Caption As String, Optional HideSelect As Boolean = True, Optional b_Grid As Boolean = True, Optional s_ColSort As String = vbNullString)

Dim a_Caption() As String, i As Byte
Dim Lbl As Label, Lbl2 As Label


Set Lbl = frmMain.Label1
With lv
    '.ListItems.Clear
    .View = lvwReport
    .ColumnHeaders.clear
    .Arrange = lvwAutoTop
    .LabelEdit = lvwManual
    .FullRowSelect = True
    '.Checkboxes = True
    .HideSelection = HideSelect
    .Font.Name = "Tahoma"
    .Font.Size = 10
    .ListItems.clear                                                  'Added by IanC
    '.ColumnHeaderIcons = f_Image.IMG1
    .GridLines = b_Grid
    .BackColor = vbWhite
    a_Caption = Split(s_Caption, "|")
    For i = 0 To UBound(a_Caption)
        Lbl.Caption = a_Caption(i)
        Lbl.AutoSize = True
        .ColumnHeaders.Add , , a_Caption(i), Lbl.Width + 500
    Next
    If s_ColSort = vbNullString Then Exit Sub
    a_Caption = Split(s_ColSort, "|")
    For i = 1 To UBound(a_Caption)
        .ColumnHeaders(i).Tag = IIf(a_Caption(i) <> "", a_Caption(i), "")
    Next
End With
End Sub

''Public Sub Lv_ResizeCol(ByVal lv As ListView)
''Dim ctr As Long, ctr1 As Long
''    ctr1 = 0
''    For ctr = ctr1 To lv.ColumnHeaders.Count - 1
''      SendMessage lv.hWnd, LVM_SETCOLUMNWIDTH, ctr, LVSCW_AUTOSIZE_USEHEADER
''    Next
''End Sub

Public Sub Lv_Background(lv As ListView, frmObj As Form, ByVal BackColorOne As OLE_COLOR, ByVal BackColorTwo As OLE_COLOR)
Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
    With lv
        If .View = lvwReport Then
            Set picAlt = frmObj.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = 225 '.ListItems(1).Height
            
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 1.9
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            frmObj.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
End Sub

Public Sub Lv_Sorting(ByVal lstView As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional ByVal TypeSort As String = "NUMBER")
On Error Resume Next
  
    With lstView
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        'LockWindowUpdate .hWnd
        
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.Index - 1
    
        Select Case UCase$(TypeSort)
        
        Case "DATE"
            strFormat = "YYYYMMDDHhNnSs"
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
        Case "NUMBER"
            strFormat = String(30, "0") & "." & String(30, "0")
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(Val(.Text)) Then
                                If CDbl(Val(.Text)) >= 0 Then
                                    .Text = Format(CDbl(Val(.Text)), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(Val(.Text)), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else   ' Assume sort by string
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
       ' LockWindowUpdate 0&
        .MousePointer = lngCursor
    End With
    
    Set lstView = Nothing
    Set ColumnHeader = Nothing
End Sub
Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function





