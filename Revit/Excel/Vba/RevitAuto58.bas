Attribute VB_Name = "RevitAuto58"
Option Explicit
Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hwnd As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Type point
    Id As Long
    column As Integer
    row As Integer
End Type

Dim ws1 As Worksheet 'Routes
Dim ws2 As Worksheet 'RouteDraw
Dim ws4 As Worksheet 'IdentityData
Dim ws5 As Worksheet 'EndPoint Entry

Const C0 As Integer = 12
Const c1 As Integer = 20
Dim Used(8127) As Boolean
Dim colPos As Integer
Const INIT_COL_POS = 4
Const COLSEP As Integer = 4
Const TRACKSEP As Integer = 6

Private Sub Delay(n As Integer)
    Sleep (CInt(Range("TimeOut").Value) * n)
End Sub

Public Sub ShowOnRevit(zoom As Boolean)

    Dim h2 As Long
    Dim n As Integer
    Dim List1 As String
    

    Dim Length As Long
    Const WindowName As String = "Autodesk Revit"
    Const ScreenName As String = "3D View"
    Const WinClass As String = "AfxFrameOrView140u"
    
    List1 = vbNullString
    Dim item As Range
    
        For Each item In Selection
            If item <> vbNullString Then
                If VarType(item.Value) = vbString Then
                    If Len(Trim(item.Value)) = 7 Then
                        List1 = List1 + item.Value + ";"
                    End If
                ElseIf VarType(item.Value) = vbDouble Then
                    If Len(CStr(item.Value)) = 7 Then
                        List1 = List1 + Str(item.Value) + ";"
                    End If
                End If
            End If
        Next
        
        List1 = Mid(List1, 1, Len(List1) - 1)
        
        If List1 = vbNullString Then Exit Sub
        
        'Debug.Print List1
        h2 = FindRevit
        
        'Debug.Print h2
        If h2 <> 0 Then
        
            SetForegroundWindow (h2)
            SendKeys ("md") 'Enter Modify Mode
            SendKeys ("is") 'Select by ID
            SendKeys (List1)
            SendKeys ("{TAB}")
            If zoom Then SendKeys ("S")
            Delay (1)
            SendKeys ("{TAB}{ENTER}")
        
        End If

End Sub
Private Sub ZoomOnExcel()

    Dim h2 As Long
    Dim n As Integer
    
    Dim PopUpText As String
    Dim item As Range
        
        h2 = FindRevit

        If h2 <> 0 Then
        
            SetForegroundWindow (h2)
            'SendKeys ("md") 'Enter Modify Mode
            SendKeys ("id") 'ID's of selection
            Delay (1)
            PopUpText = GetPopUpText
            SendKeys ("{ENTER}")
            
            If (PopUpText <> vbNullString) Then
                FindCell Worksheets("RouteDraw"), PopUpText
            End If
        End If

End Sub
Private Sub FindCell(ws As Worksheet, searchValue As String)
    
    Dim cell As Range
    Dim Value1 As Variant
    Dim Values() As String: Values = Split(searchValue, ",")
    For Each cell In ws.UsedRange
        If cell.Value <> vbNullString Then
            For Each Value1 In Values
                If Trim(cell.Value) = Value1 Then
                    cell.Activate
                    cell.Interior.Color = vbGreen
                End If
            Next
        End If
    Next cell

End Sub
Private Sub ShowOnExcel2()
    
    Dim h2 As Long
    Dim n As Integer
    
    Dim PopUpText As String
    Dim item As Range
        
        h2 = FindRevit

        If h2 <> 0 Then
        
            SetForegroundWindow (h2)
            'SendKeys ("md") 'Enter Modify Mode
            SendKeys ("id") 'ID's of selection
            Delay (1)
            PopUpText = GetPopUpText
            SendKeys ("{ENTER}")
            
            
            Dim Target As Range: Set Target = ActiveCell
                        
            Dim Value1 As Variant
            Dim index As Integer: index = 0
            Dim pos As Integer
            Dim DataPrev As Variant
            
            If PopUpText <> vbNullString Then

                For Each Value1 In Split(PopUpText, ",")
                    Target.Range("A1:N1").Value = Array("Id", "X1", "Y1", "Z1", "X2", "Y2", "Z2", "X3", "Y3", "Z3", "X4", "Y4", "Z4", "Dist")
                    
                    Target.Offset(index + 1, 0).Value = "'" + Str(Value1)
                    
                    On Error GoTo NoData
                    Dim Data As Variant: Data = GetData(Str(Value1))
                                            
                    If Not IsEmpty(DataPrev) Then
                        Target.Offset(index + 1, 13).Value = GetMinDistance(CInt(Data(1, C0 - 1)), CInt(DataPrev(1, C0 - 1)))
                    End If
                    
                    DataPrev = Data
                    pos = 1
                    For n = 20 To 31
                        Target.Offset(index + 1, pos).Value = Data(1, n)
                        pos = pos + 1
                    Next n
NoData:
                    index = index + 1
                Next
            End If
        End If

End Sub
Private Sub ShowOnExcel()
    
    Dim h2 As Long
    Dim n As Integer
    
    Dim PopUpText As String
    Dim item As Range
        
        h2 = FindRevit

        If h2 <> 0 Then
        
            SetForegroundWindow (h2)
            'SendKeys ("md") 'Enter Modify Mode
            SendKeys ("id")
            Delay (1)
            PopUpText = GetPopUpText
            SendKeys ("{ENTER}")
            
            
            Dim Target As Range: Set Target = ActiveCell
                        
            Dim Value1 As Variant
            Dim index As Integer: index = 0
            Dim pos As Integer
            Dim DataPrev As Variant
            
            If PopUpText <> vbNullString Then

                For Each Value1 In Split(PopUpText, ",")
                    Dim HeaderText As Variant
                    HeaderText = Array("Id", "Type", "Level", "Fitting", "Mark", "Service", "Width", "Height", "X1", "Y1", "Z1", "X2", "Y2", "Z2", "X3", "Y3", "Z3", "X4", "Y4", "Z4", "Dist")
                    
                    Range(Target.Cells(1, 1), Target.Cells(1, UBound(HeaderText) + 1)).Value = HeaderText
                    
                    'On Error GoTo NoData
                    Dim Data As Variant: Data = GetData(Str(Value1))
                    If Not VarType(Data) = vbString Then
                    
                        pos = 0
                        Target.Offset(index + 1, pos).Value = Value1: pos = pos + 1 'Id
                        Target.Offset(index + 1, pos).Value = Data(1, 5): pos = pos + 1 'Type
                        Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 44)): pos = pos + 1 'Level
                        Target.Offset(index + 1, pos).Value = (InStr(Data(1, 6), "Cable Tray") = 0): pos = pos + 1 'Fitting
                        Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 2)): pos = pos + 1 'Mark
                        Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 3)): pos = pos + 1 'Service
                        'Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 4)): pos = pos + 1 ' Comments
                        Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 9)): pos = pos + 1 'Width
                        Target.Offset(index + 1, pos).Value = "'" + CStr(Data(1, 10)): pos = pos + 1 'Height
                                                
                        Dim MinDistance As Double
                        If Not IsEmpty(DataPrev) Then
                            MinDistance = GetMinDistance(CInt(Data(1, C0 - 1)), CInt(DataPrev(1, C0 - 1)))
                        End If
                        
                        DataPrev = Data
                        For n = 20 To 31
                            Target.Offset(index + 1, pos).Value = Data(1, n)
                            pos = pos + 1
                        Next n
                        Target.Offset(index + 1, pos).Value = MinDistance
                    End If
NoData:

                    index = index + 1
                Next
            End If
        End If

End Sub
Private Function GetMinDistance(index1 As Integer, index2 As Integer) As Double

    Set ws1 = Worksheets("Routes")
    Dim Vals1 As Variant
    Dim Vals2 As Variant
    Vals1 = Range(ws1.Cells(index1 + 1, 20), ws1.Cells(index1 + 1, 31)).Value
    Vals2 = Range(ws1.Cells(index2 + 1, 20), ws1.Cells(index2 + 1, 31)).Value
    Dim Dist As Double: Dist = 99999
    Dim Dist1 As Double

    Dim n As Integer
    Dim m As Integer
    
    For n = 0 To 3
        For m = 0 To 3
            If Vals1(1, 1 + 3 * n) <> vbNullString And Vals2(1, 1 + 3 * m) <> vbNullString Then
                 Dist1 = Sqr(((Vals1(1, 1 + 3 * n) - Vals2(1, 1 + 3 * m)) ^ 2) + _
                            ((Vals1(1, 2 + 3 * n) - Vals2(1, 2 + 3 * m)) ^ 2) + _
                            ((Vals1(1, 3 + 3 * n) - Vals2(1, 3 + 3 * m)) ^ 2))
                'Debug.Print CStr(Dist1) + ":" + CStr(n) + ":" + CStr(m)
                If (Dist1 < Dist) Then Dist = Dist1
            End If
        Next m
    Next n

    GetMinDistance = Dist

End Function
Private Function GetPopUpText() As String
    Dim text As String * 255
    Dim h1 As Long
    Const name As String = "Element IDs of Selection"
    Const DialogClass As String = "#32770"
    h1 = FindWindowEx(0, 0, DialogClass, name)
    h1 = FindWindowEx(h1, 0, "Edit", vbNullString)
    Dim count As Long
    GetPopUpText = GetText(h1)
    
End Function

Private Function GetText(hwnd As Long) As String
    
    Const WM_GETTEXT = &HD
    Const WM_GETTEXTLENGTH = &HE
   
    Dim strText As String
    Dim lngLength As Long
    Dim lngRetVal As Long
    
    lngLength = SendMessage(hwnd, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
    strText = Space$(lngLength)
    lngRetVal = SendMessage(hwnd, WM_GETTEXT, ByVal lngLength, ByVal strText)
    strText = Left$(strText, lngRetVal)
    GetText = strText
    
End Function
Private Function GetData(Id As String)
    Dim ws1 As Worksheet
    Dim row1 As Integer

    Set ws1 = Worksheets("Routes")
    On Error Resume Next
    row1 = 0
    row1 = Application.WorksheetFunction.Match(CLng(Id), ws1.Columns(1), 0)
    On Error GoTo 0
    If row1 > 0 Then
        Dim Data As Variant
        Data = ws1.Range(ws1.Cells(row1, 1), ws1.Cells(row1, ws1.UsedRange.Columns.count))
        GetData = Data
    Else
        GetData = "NoData"
    End If
End Function
Private Function FindRevit() As Long
    Dim h1 As Long
    Const WindowName As String = "Autodesk Revit"
    Const ScreenName As String = "3D View"
    Const WinClass As String = "AfxFrameOrView140u"
    h1 = GetDesktopWindow()
    h1 = FindWindowByName(h1, WindowName)
    h1 = FindWindowByName(h1, ScreenName)
    h1 = FindWindowEx(h1, 0, WinClass, vbNullString)
    FindRevit = h1
End Function
Private Function FindWindowByName(h1 As Long, WindowName As String) As Long
    
    Dim h2 As Long
    Dim text As String * 255
    
    Do
        h2 = FindWindowEx(h1, h2, vbNullString, vbNullString)
        Call GetWindowText(h2, text, 255)
        If (Left(text, Len(WindowName)) = WindowName) Then
            Exit Do
        End If
        If h2 = 0 Then Exit Do
    Loop
    FindWindowByName = h2

End Function

Public Sub AddToContextMenu()
    Dim ContextMenu As CommandBar
    Dim newButton As CommandBarButton
    Dim Contr1 As Variant
    ' Reference the context menu for cells
    Set ContextMenu = Application.CommandBars("Cell")
    
    For Each Contr1 In ContextMenu.Controls
    'Debug.Print Contr1.Caption
        If Contr1.Caption = "Redraw Chart" Then
            Exit Sub
        End If
    Next Contr1

    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Selection Length" ' Text that appears in the menu
        .OnAction = "SelectionLength"    ' Macro to run when clicked
        .FaceId = 226        ' Icon for the menu item (optional)
        .BeginGroup = True
    End With
    
    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Mark Length Sum" ' Text that appears in the menu
        .OnAction = "MarkLengthSum"    ' Macro to run when clicked
        .FaceId = 308        ' Icon for the menu item (optional)
        .BeginGroup = False
    End With

    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Color Length" ' Text that appears in the menu
        .OnAction = "FindColor"    ' Macro to run when clicked
        .FaceId = 1691        ' Icon for the menu item (optional)
    End With

    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Redraw Chart" ' Text that appears in the menu
        .OnAction = "Dataloop"    ' Macro to run when clicked
        .FaceId = 37        ' Icon for the menu item (optional)
        .BeginGroup = True
    End With
    
    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Format Data" ' Text that appears in the menu
        .OnAction = "FormatData"    ' Macro to run when clicked
        .FaceId = 16            ' Icon for the menu item (optional)
    End With
    
    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Show on Revit" ' Text that appears in the menu
        .OnAction = "ShowOnRevit(false)"    ' Macro to run when clicked
        .FaceId = 417           ' Icon for the menu item (optional)
        .BeginGroup = True
    End With

    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Zoom on Revit" ' Text that appears in the menu
        .OnAction = "ShowOnRevit(true)"    ' Macro to run when clicked
        .FaceId = 645            ' Icon for the menu item (optional)
    End With

    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Show on Excel" ' Text that appears in the menu
        .OnAction = "ShowOnExcel"    ' Macro to run when clicked
        .FaceId = 263            ' Icon for the menu item (optional)
        .BeginGroup = True
    End With
    
    With ContextMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Find on Excel" ' Text that appears in the menu
        .OnAction = "ZoomOnExcel"    ' Macro to run when clicked
        .FaceId = 186            ' Icon for the menu item (optional)
    End With

    With CommandBars("Ply").Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = "Save Selected Sheets" ' Text that appears in the menu
        .OnAction = "SaveSelectedSheets"    ' Macro to run when clicked
        .FaceId = 1548        ' Icon for the menu item (optional)
        .BeginGroup = True
    End With

    AddCustomMenu

End Sub


Private Sub ResetMenus()
    Dim ContextMenu As CommandBar
    Set ContextMenu = Application.CommandBars("Cell")
        
        Dim Contr1 As Variant
    For Each Contr1 In ContextMenu.Controls
        'Debug.Print Contr1.Caption
        If Contr1.Caption = "Show on Revit" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Zoom on Revit" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Show on Excel" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Zoom on Excel" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Redraw Chart" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "SaveSelectedSheets" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Selection Length" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Color Length" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Format Data" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Find on Excel" Then
            Contr1.Delete
        ElseIf Contr1.Caption = "Mark Length Sum" Then
            Contr1.Delete
        End If
        
    Next

    Application.CommandBars("Revit").Delete

End Sub

'Private Sub ResetAllContextMenus()
'    Dim cmdBar As CommandBar
'    For Each cmdBar In Application.CommandBars
'        If cmdBar.BuiltIn Then
'            cmdBar.Reset
'        End If
'    Next cmdBar
'End Sub

Private Sub DataLoop()
    
    Set ws1 = Worksheets("Routes")
    Set ws2 = Worksheets("RouteDraw")
    Set ws5 = Worksheets("EndPoint Entry")
    ws2.Rows.Delete
    With ws2.Rows(1).Font
        .Bold = True
        .Size = 14
    End With
    ws2.Cells(1, 1) = "Routes"
    ws2.Cells.Interior.Pattern = xlNone
    
    Dim ref As Integer
    Dim n As Integer
    
    Dim rowA As Integer
    Dim colA As Integer
    Dim rowB As Integer
    Dim rowC As Integer
    
    Const TOP_ROW = 3
    
    colPos = INIT_COL_POS
   
    For n = 0 To 1023
        Used(n) = False
    Next n
   
    Dim count As Integer

    '******************************************************************************
    
    rowC = 2
    Do
        If ws5.Cells(rowC, 1) = vbNullString Then Exit Do
        Dim Id As Long: Id = ws5.Cells(rowC, 1)
        ws5.Cells(rowC, 3) = "NOT FOUND"
        Dim Ref1 As Variant
        On Error GoTo skip1
        Ref1 = Application.WorksheetFunction.VLookup(Id, ws1.UsedRange, C0 - 1, False)
        rowA = CInt(Ref1)
        ws5.Cells(rowC, 3) = "OK"
        
        If (rowA > 0) Then
        
            rowA = rowA + 1
            count = 0
            For n = 0 To 3
                If ws1.Cells(rowA, C0 + n) <> 0 Then count = count + 1
            Next n
            
            If count = 1 Then
                ref = ws1.Cells(rowA, C0 - 1)
                rowB = TOP_ROW
                ws2.Cells(rowB - 2, colPos - 2) = ws5.Cells(rowC, 2)
                If Not (Used(ref)) Then
                    Call SplitCol(ref, 9999, rowB, colPos)
                    colPos = colPos + TRACKSEP
                End If
            End If
        
        End If
        
        rowC = rowC + 1
            
    Loop
    
skip1:
    
    On Error GoTo 0
    '******************************************************************************
    
    rowA = 2
    Do
        If ws1.Cells(rowA, 1) = vbNullString Then Exit Do
            count = 0
            For n = 0 To 3
                If ws1.Cells(rowA, C0 + n) <> 0 Then count = count + 1
            Next n
            
            If count = 1 Then
                ref = ws1.Cells(rowA, C0 - 1)
                rowB = TOP_ROW
                If Not (Used(ref)) Then
                    Call SplitCol(ref, 9999, rowB, colPos)
                    colPos = colPos + TRACKSEP
                End If
            End If
        rowA = rowA + 1
    Loop

    GetIdentityData
    GetTotalTrays
    GetTotalFittings
    ShortCircuits

End Sub

Private Sub SplitCol(ByVal actRef As Integer, ByVal oldRef As Integer, ByVal rowNum As Integer, ByVal prevcolpos As Integer)
        
        
        DoEvents
        Dim rowNum1 As Integer
        Dim colpos1 As Integer
        Dim prevUsed As Boolean
        
        rowNum1 = rowNum
        colpos1 = colPos
        
        prevUsed = Used(actRef)
      
        Used(actRef) = True
        
        Dim colInc As Integer: colInc = 0
        Dim newRef As Integer
        Dim n As Integer
        
        If actRef = 524 Then
            actRef = actRef
            
        End If
        
        If (colpos1 <> prevcolpos) Then ws2.Cells(rowNum1 - 1, colpos1 + 1) = Range("Elb")

        Dim Branches(3) As Integer
        Dim BC As Integer: BC = 0
        
        'Find all routes
        For n = 0 To 3
            newRef = ws1.Cells(actRef + 1, C0 + n)
            If ((newRef <> 0) And (oldRef <> newRef)) Then
                Branches(BC) = newRef
                BC = BC + 1
            End If
        Next n
       
        'Process all branches
        For n = 0 To BC - 1
            If (colInc > 0) Then colPos = colPos + COLSEP
            If Not prevUsed Then Call SplitCol(Branches(n), actRef, rowNum + 1, colpos1)
            colInc = colInc + 1
        Next n
        
        With ws2
            'Cells(rowNum1, colpos1) = actRef
            .Cells(rowNum1, colpos1) = "'" + Str(ws1.Cells(actRef + 1, 1)) 'Here is ID
            .Cells(rowNum1, colpos1 - 1) = Format(ws1.Cells(actRef + 1, 8), "0####") 'Here is Length
            '.Cells(rowNum1, colpos1 - 2) = Format(ws1.Cells(actRef + 1, 3), "0####") 'Comments
            .Cells(rowNum1, colpos1 - 2) = Format(ws1.Cells(actRef + 1, 2), "0####") 'Mark 2
            'Cell used stop to avoid infinite loops
            If prevUsed Then .Cells(rowNum1, colpos1).Interior.Color = vbRed
            
            'Fill Horizontal Lines
            For n = prevcolpos + 1 To colpos1
                If .Cells(rowNum1 - 1, n) <> Range("Elb") Then
                    .Cells(rowNum1 - 1, n) = Range("Hor")
                Else
                    .Cells(rowNum1 - 1, n) = Range("Cross")
                End If
            Next n
            
            'Fill Vertical and Tee lines
            If colInc > 1 Then
                .Cells(rowNum1, colpos1 + 1) = Range("Tee")
            Else
                .Cells(rowNum1, colpos1 + 1) = Range("Ver")
            End If
        End With
        

End Sub

Private Sub FillReferences()

    Dim ws1 As Worksheet
    Dim ptr1 As Integer
    Dim n As Integer
    Dim rowNum As Integer
    
    Set ws1 = Worksheets("Routes")
    
    ptr1 = 2

    Do
        If ws1.Cells(ptr1, 1) = vbNullString Then Exit Do
            
            For n = 0 To 3
                            
                If ws1.Cells(ptr1, C0 + n) <> vbNullString Then
                    rowNum = ws1.Cells(ptr1, C0 + n) + 1
                    ws1.Cells(ptr1, C0 + 4 + n) = ws1.Cells(rowNum, 1)
                End If
    
            Next n
    
        ptr1 = ptr1 + 1
    Loop

End Sub

Private Function GetLength(Id As Long) As Double

    GetLength = Round(Application.WorksheetFunction.VLookup(Id, Worksheets("Routes").Columns(1), 8, False), 0)

End Function

Private Sub FormatData()

    If ActiveSheet.name = "Routes" Or ActiveSheet.name = "Document Data" Then
        Range("A1").Select
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        Cells.EntireColumn.AutoFit
    End If
End Sub

Private Sub GetIdentityData()
    
    Dim r1 As Integer
    Dim c1 As Integer
    Dim ptr4 As Integer
    Dim strVal As String
    ptr4 = 2
    Set ws2 = Worksheets("RouteDraw")
    Set ws4 = Worksheets("IdentityData")
    ws4.Cells.Clear
    Dim cell As Variant
    
    ws4.Cells(1, 1) = "Id"
    ws4.Cells(1, 2) = "Mark"
    
    For Each cell In ws2.UsedRange
        If IsNumeric(cell.Value) Then
            strVal = CStr(cell.Value)
            If Len(Trim(strVal)) = 7 Then
                ws4.Cells(ptr4, 1) = strVal
                ws4.Cells(ptr4, 2) = ws2.Cells(cell.row, cell.column - 2)
                ptr4 = ptr4 + 1
            End If
        End If
    Next

End Sub

Private Sub AddCustomMenu()
    
    Const MenuName As String = "REVIT"
    Dim cmdBar As CommandBar
    Dim cmdBar2 As CommandBar
    Dim subMenu As Object
    Dim newMenu As CommandBarPopup
    Dim newButton As CommandBarButton
    Dim bar As Object
    
    Dim message As String
    Dim n As Integer
    
    Dim MC As Variant
    MC = Array("CableTrayRoutes", "DisplayId", "DrawTrays", "GetCadData", "IdentityData")
    
    Dim Id As Variant
    Id = Array(1149, 1784, 2082, 1958, 1954)
    
    Dim ToolTips As Variant
    ToolTips = Array("Analize cable tray routes in Sheet Routes", _
                    "Attach Id text to Cable Tray Elements", _
                    "Draw new trays using Excel coordinates", _
                    "Get data from type and level data from Revit", _
                    "Fill Cable Tray Mark, Comments and Service Type fields")
    
    For Each bar In Application.CommandBars
        If bar.name = MenuName Then
            Application.CommandBars(MenuName).Delete
            Exit For
        End If
    Next
    
    Set cmdBar = Application.CommandBars.Add(MenuName, msoBarLeft, False, True)
    
    For n = 0 To UBound(MC)
        With cmdBar.Controls.Add(Type:=msoControlButton)
            .Caption = """" + MC(n) + """"
            .OnAction = "'ExecMacro """ + MC(n) + """'"
            .FaceId = Id(n)
            .Style = msoButtonIconAndCaption
            .TooltipText = ToolTips(n)
            .BeginGroup = True
            If MC(n) = "GetCadData" Then .BeginGroup = True
        End With
    Next n
    cmdBar.Visible = True

    With cmdBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Reset Menus"
        .OnAction = "ResetMenus"
        .FaceId = 358
        .Style = msoButtonIconAndCaption
        .TooltipText = "Reset Program Menus"
    End With

    With cmdBar.Controls.Add(Type:=msoControlButton)
        .Caption = "Get Mark Length"
        .OnAction = "GetMarkLength"
        .FaceId = 1150
        .Style = msoButtonIconAndCaption
        .TooltipText = "Get Length of Trays with same Mark"
    End With

End Sub
Private Sub DeleteMenu(name As String)
    Application.CommandBars(name).Delete
End Sub

Public Sub ExecMacro(Macro As String)
    
    Const TIMEOUTMAX = 50
    
    Dim TimeOut As Integer
    Dim h2 As Long
    Dim h3 As Long
    h2 = FindRevit
    'Debug.Print "Start"
    
    If h2 = 0 Then
        MsgBox ("Revit not Found")
        Exit Sub
    End If
    
    Delay (5)
    SetForegroundWindow (h2)

    Delay (5)
    'SendKeys ("{F10}gvp")
    SendKeys ("dp") 'Custom Dynamo Player shortcut

    TimeOut = TIMEOUTMAX
    Do
        Delay (1)
        h3 = FindWindowEx(0, 0, vbNullString, "Dynamo Player")
        If h3 <> 0 Or TimeOut <= 0 Then Exit Do
        TimeOut = TimeOut - 1
    Loop
    
    'Debug.Print "TimeOut:" + CStr(TimeOut) + " Hadle:" + CStr(h3)
    
    Application.DisplayAlerts = False
        
    If h3 > 0 Then
        'SetForegroundWindow (h3)
        Delay (TIMEOUTMAX - TimeOut)
        Delay (7)
        SendKeys ("{TAB 4}")
        Delay (3)
        SendKeys (Macro)
        Delay (3)
        SendKeys ("{TAB 6}{ENTER}")
        SendKeys ("%{F4}")
    End If

    Application.DisplayAlerts = True

End Sub

Public Sub SelectionLength()
    Dim area As Range
    Dim i As Integer
    Dim item As Variant
    Dim sum1 As Long: sum1 = 0
    
    If Not ActiveSheet.name = "RouteDraw" Then
         Call MsgBox("Mark Length", vbInformation, "For this Function Select Sheet")
         Exit Sub
    End If
    
    
    For i = 1 To Selection.Areas.count
        Set area = Selection.Areas(i)
        For Each item In area
            sum1 = sum1 + item.Offset(0, -1).Value
        Next
    Next i

    Call MsgBox(sum1, vbInformation, "Total Length")

End Sub

Private Sub CopyToClipboard(ByVal TextToCopy As String)
    Const CF_TEXT As Long = 1
    Const GMEM_MOVEABLE As Long = &H2

    Dim hGlobalMemory As LongPtr
    Dim lpMemory As LongPtr
    Dim hwnd As LongPtr
    
    ' Allocate global memory
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, Len(TextToCopy) + 1)
    lpMemory = GlobalLock(hGlobalMemory)
    
    ' Copy the text to the allocated memory
    RtlMoveMemory ByVal lpMemory, ByVal StrPtr(TextToCopy), Len(TextToCopy) + 1
    GlobalUnlock hGlobalMemory
    
    ' Open clipboard and set the data
    If OpenClipboard(hwnd) Then
        EmptyClipboard
        SetClipboardData CF_TEXT, hGlobalMemory
        CloseClipboard
    End If
End Sub

Public Sub FindColor()

Set ws2 = Worksheets("RouteDraw")

Dim sum1 As Long: sum1 = 0
Dim color1 As Long
Dim item As Range
color1 = ActiveCell.Interior.Color

If color1 = &HFFFFFF Then Exit Sub

For Each item In ActiveSheet.UsedRange
    If item.Interior.Color = color1 Then
        sum1 = sum1 + item.Offset(0, -1).Value
    End If
Next

ActiveCell.Value = sum1

End Sub
Private Sub GetTotalTrays()
Dim ws1 As Worksheet
Dim ws6 As Worksheet
Dim ptr1 As Integer
Dim ptr6 As Integer
Const FamilyName = 7

Set ws1 = Worksheets("Routes")
Set ws6 = Worksheets("TotalTrays")
ws6.Cells.Clear
Dim cn As Integer
Dim n As Variant
ptr6 = 1

Dim myCols As Variant
myCols = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

'Dim CList As String: CList = vbNullString
'For n = 0 To UBound(myCols)
'    If n > 0 Then CList = CList + ","
'    CList = CList + ws1.Cells(1, myCols(n))
'Next n
'
'Dim colNames() As String: colNames = Split(CList, ",")
'Dim myCols2 As String
'For n = 0 To UBound(colNames)
'    If n > 0 Then myCols2 = myCols2 + ","
'    myCols2 = myCols2 + CStr(ws1.Rows(1).Find(colNames(n)).Column)
'Next
'
'Debug.Print CList
'Debug.Print myCols2

For ptr1 = 1 To ws1.UsedRange.Rows.count
    
    If ptr1 = 1 Or ws1.Cells(ptr1, FamilyName) = "Cable Tray with Fittings" Then
        cn = 1
        For Each n In myCols
            ws6.Cells(ptr6, cn) = ws1.Cells(ptr1, n)
            cn = cn + 1
        Next n
        ptr6 = ptr6 + 1
    End If
Next ptr1

Dim myTable As ListObject

Set myTable = ws6.ListObjects.Add(xlSrcRange, ws6.UsedRange, , xlYes)
myTable.name = ws6.name + "Table"
SortTotalTrays

Dim PivotCache1 As PivotCache
Dim PivotTable1 As PivotTable
Dim PivotPos As Range: Set PivotPos = ws6.Cells(2, ws6.UsedRange.Columns.count + 2)

Set PivotCache1 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    myTable.name, Version:=8)

Set PivotTable1 = PivotCache1.CreatePivotTable(TableDestination:= _
    PivotPos, DefaultVersion:=8)

With PivotTable1
    .AddDataField(.PivotFields("Length"), "Sum of Length", xlSum).NumberFormat = "#,##0"
    .PivotFields("TypeName").Orientation = xlRowField
    .PivotFields("Width").Orientation = xlRowField
    .PivotFields("Height").Orientation = xlRowField
    .RowAxisLayout xlTabularRow

End With


End Sub
Private Sub GetTotalFittings()
Dim ws1 As Worksheet
Dim ws6 As Worksheet
Dim ptr1 As Integer
Dim ptr6 As Integer
Const FamilyName = 7

Set ws1 = Worksheets("Routes")
Set ws6 = Worksheets("TotalFittings")
ws6.Cells.Clear
Dim cn As Integer
Dim n As Variant
ptr6 = 1

Dim myCols As Variant
myCols = Array(1, 2, 3, 4, 5, 6, 7, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43)

For ptr1 = 1 To ws1.UsedRange.Rows.count
    
    If ptr1 = 1 Or Not ws1.Cells(ptr1, FamilyName) = "Cable Tray with Fittings" Then
        cn = 1
        For Each n In myCols
            ws6.Cells(ptr6, cn) = ws1.Cells(ptr1, n)
            cn = cn + 1
        Next n
        ptr6 = ptr6 + 1
    End If
Next ptr1

Dim myTable As ListObject

Set myTable = ws6.ListObjects.Add(xlSrcRange, ws6.UsedRange, , xlYes)
myTable.name = ws6.name + "Table"
SortTotalFittings

Dim PivotCache1 As PivotCache
Dim PivotTable1 As PivotTable
Dim PivotPos As Range: Set PivotPos = ws6.Cells(2, ws6.UsedRange.Columns.count + 2)
Dim pf As PivotField

Set PivotCache1 = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    myTable.name, Version:=8)

Set PivotTable1 = PivotCache1.CreatePivotTable(TableDestination:= _
    PivotPos, DefaultVersion:=8)

With PivotTable1
    .AddDataField .PivotFields("FamilyName"), "Count of FamilyName", xlCount
    With .PivotFields("FamilyName")
        .Orientation = xlRowField
        .LabelRange = "Type Of Tray"
    End With
    .PivotFields("Tray Width").Orientation = xlRowField
    .PivotFields("Tray Height").Orientation = xlRowField
    .PivotFields("Angle").Orientation = xlRowField
    .PivotFields("Bend Radius").Orientation = xlRowField
    .RowAxisLayout xlTabularRow

End With

With PivotTable1.PivotFields("FamilyName")
    
    Dim item As Variant
    
    For Each item In .PivotItems
      'Debug.Print item.Caption
      If Not InStr(1, item.Caption, "Bend") > 0 Then
        item.ShowDetail = False
      End If
    Next
    
End With

'PivotTable1.PivotFields("Bend Radius").LabelRange.Cells(5, 1) = "Juanito"

'For Each pf In PivotTable1.ColumnFields
'    'pf.ShowDetail = True
'Next

End Sub
Public Sub SaveSelectedSheets()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim SelWs As Sheets
    Set SelWs = ActiveWindow.SelectedSheets
    
    For Each ws1 In ActiveWorkbook.Worksheets
        For Each ws2 In SelWs
            If ws1.name = ws2.name Then
                'Debug.Print "Vis:" + ws1.name
                ws1.Visible = xlSheetVisible
                Exit For
            Else
                'ws1.Visible = xlSheetHidden
                'Debug.Print "Hid:" + ws1.name
                ws1.Visible = xlSheetVeryHidden
            End If
        Next ws2
    Next ws1
    
    SaveAsDialog (ActiveWorkbook.Path + "\" + ActiveWorkbook.name)
    
    For Each ws1 In ActiveWorkbook.Worksheets
        ws1.Visible = xlSheetVisible
    Next
    
End Sub

Private Sub SaveAsDialog(FilePath As String)
    Dim fd As FileDialog

    ' Create a FileDialog object as a Save As dialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Save As"
        .InitialFileName = FilePath ' Default file path and name
        .AllowMultiSelect = False ' Allow only one file to be selected

        ' Show the dialog box
        If .Show = -1 Then ' If the user selects a file
            FilePath = .SelectedItems(1) ' Get the selected file path
            ActiveWorkbook.SaveCopyAs FileName:=FilePath
        Else
            MsgBox "No file selected."
        End If
    End With
End Sub

Private Sub SortTotalTrays()
    
    With Worksheets("TotalTrays").ListObjects("TotalTraysTable").Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("TotalTraysTable[TypeName]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalTraysTable[Width]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalTraysTable[Height]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
    End With

End Sub
Private Sub SortTotalFittings()
    
    With Worksheets("TotalFittings").ListObjects("TotalFittingsTable").Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("TotalFittingsTable[FamilyName]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Width]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Height]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Angle]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Bend Radius]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Width 1]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Width 2]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Width 3]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Height 1]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Tray Height 2]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Length 1]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range("TotalFittingsTable[Length 3]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
    End With

End Sub
Public Sub GetMarkLength()
    Dim ws1 As Worksheet
    Dim ws6 As Worksheet
    Set ws1 = Worksheets("Routes")
    Set ws6 = Worksheets("Mark Length")
    ws6.Cells.Clear
    Dim ptr1 As Integer
    Dim Id As String
    Dim Mark As String
    Dim Length As Double
    Dim rowNum As Integer
    Dim markVals(2048) As String
    Dim lenVals(2048) As Double
    Dim IdVals(2048) As String
    Dim markCnt As Integer: markCnt = 0
    
    Dim n As Integer
    Dim m As Integer
    Dim found As Boolean
    
    For ptr1 = 2 To ws1.UsedRange.Rows.count
        found = False
        Id = CStr(ws1.Cells(ptr1, 1))
        Mark = ws1.Cells(ptr1, 2)
        Length = ws1.Cells(ptr1, 8)
            If Mark <> vbNullString Then
                                            
                For n = 0 To markCnt - 1
                    If markVals(n) = Mark Then
                        lenVals(n) = lenVals(n) + Length
                        IdVals(n) = IdVals(n) + ";" + Id
                        found = True
                        Exit For
                    End If
                Next n
                If Not found Then
                    markVals(markCnt) = Mark
                    lenVals(markCnt) = Length
                    IdVals(n) = Id
                    markCnt = markCnt + 1
                End If
                            
            End If
       
    Next ptr1
   
    For n = 1 To markCnt
        ws6.Cells(n, 1) = markVals(n - 1)
        ws6.Cells(n, 2) = lenVals(n - 1)
        Dim Ids() As String: Ids = Split(IdVals(n - 1), ";")
        For m = 0 To UBound(Ids)
            ws6.Cells(n, 3 + m) = Ids(m)
        Next m
    Next n
    
End Sub
Private Sub ColorColumn()
    
Dim n As Integer
Dim ws6 As Worksheet
Set ws6 = Worksheets("Mark Length")

For n = 1 To ws6.UsedRange.Rows.count

    ws6.Cells(n, 1).Interior.ColorIndex = n Mod 56 + 1

Next n

End Sub

Private Sub ColorRouteDraw()

Dim ws2 As Worksheet
Dim Data As Variant
Dim item As Variant
Set ws2 = Worksheets("RouteDraw")

For Each item In ws2.UsedRange

    If Left(item, 1) = "T" And Len(item) = 5 Then
        item.Interior.Color = CLng(Right(item, 4)) * 16 + 256
    End If
    
Next

End Sub

Private Sub ResetUsedRange()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.UsedRange.Calculate
End Sub

Private Sub MarkLengthSum()
    
    Dim ws6 As Worksheet
    Set ws6 = Worksheets("Mark Length")
    Dim area As Range
    Dim i As Integer
    Dim item As Variant
    Dim sum1 As Long: sum1 = 0
    Dim Id As Variant
    Dim n As Integer
    Dim List1 As String
    Dim h2 As Long
        
    If Not ActiveSheet.name = "Mark Length" Then
         Call MsgBox("Mark Length", vbInformation, "For this Function Select Sheet")
         Exit Sub
    End If
    
    For i = 1 To Selection.Areas.count
    Set area = Selection.Areas(i)
        For Each item In area
            sum1 = sum1 + item.Offset(0, 1).Value
            n = 0
            Do
                Id = item.Offset(0, 2 + n)
                If Id = vbNullString Then Exit Do
                List1 = List1 + CStr(Id) + ","
                n = n + 1
            Loop
        Next
    Next i
    
    Call MsgBox(sum1, vbInformation, "Total Length")
    
    Delay (1)
    
    If List1 = vbNullString Then Exit Sub
    List1 = Mid(List1, 1, Len(List1) - 1)
        
    'Debug.Print List1
    h2 = FindRevit
    
    'Debug.Print h2
    If h2 <> 0 Then
    
        SetForegroundWindow (h2)
        SendKeys ("md") 'Enter Modify Mode
        SendKeys ("is") 'Select by ID
        SendKeys (List1)
        SendKeys ("{TAB}")
        SendKeys ("S")
        Delay (2)
        SendKeys ("{TAB}{ENTER}")
    
    
    End If

End Sub

Private Sub ShortCircuits()

Dim ws2 As Worksheet 'RouteDraw
Set ws2 = Worksheets("RouteDraw")

Dim shape1 As Variant
For Each shape1 In ws2.Shapes
    shape1.Delete
Next

ws2.UsedRange.Calculate
Dim item As Variant
Dim points(1024) As point
Dim pointCount As Integer: pointCount = 0
Dim n As Integer

    For Each item In ws2.UsedRange
        If item.Interior.Color = vbRed Then
            points(pointCount).Id = item.Value
            points(pointCount).row = item.row
            points(pointCount).column = item.column
            pointCount = pointCount + 1
        End If
    Next item

    For Each item In ws2.UsedRange
        Dim Id As Long
        Id = getId(item.Value)
        
        If Id > 0 Then
            For n = 0 To pointCount - 1
                If (Id = points(n).Id) _
                And Not ((item.row = points(n).row) _
                And (item.column = points(n).column)) Then
                    Call JoinPoints(ws2, item.row, item.column, _
                    points(n).row, points(n).column)
                End If
            Next n
        End If
    Next item

End Sub
Private Sub JoinPoints(ByRef ws2 As Worksheet, r1 As Integer, c1 As Integer, r2 As Integer, c2 As Integer)

    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    Dim cn1 As Shape

    x1 = ws2.Cells(r1, c1).Left + ws2.Cells(r1, c1).Width / 2
    y1 = ws2.Cells(r1, c1).Top + ws2.Cells(r1, c1).Height / 2

    x2 = ws2.Cells(r2, c2).Left + ws2.Cells(r2, c2).Width / 2
    y2 = ws2.Cells(r2, c2).Top + ws2.Cells(r2, c2).Height / 2

    Set cn1 = ws2.Shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)
    cn1.Line.ForeColor.RGB = RGB(&HEE, &H82, &HE)  ' Apricot
    cn1.Line.Weight = 1
    cn1.Line.BeginArrowheadStyle = msoArrowheadDiamond
    cn1.Line.EndArrowheadStyle = msoArrowheadDiamond
    

End Sub

Function getId(Data As Variant) As Long

getId = 0

If VarType(Data) = vbString Then
    Data = Replace(Data, "'", "")
    If Len(Data) = 8 Then
        On Error Resume Next
        getId = CLng(Data)
    End If
Else
    getId = Data
End If

End Function







