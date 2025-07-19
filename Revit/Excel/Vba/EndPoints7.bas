Attribute VB_Name = "EndPoints7"
Option Explicit
Dim ws1 As Worksheet 'Routes
Dim ws2 As Worksheet 'RouteDraw
Dim ws3 As Worksheet 'EndPoints
Dim Data As Variant

Type Symbols
    Ver As String
    Hor As String
    Tee As String
    Elb As String
    Cross As String
    EP As String
End Type
    
Dim EpRow As Integer
Dim EpCol As Integer
Dim Sym As Symbols

Public Sub FolowEndpoints()

    Dim r1 As Integer
    Dim c1 As Integer

    getEndpoints

    DefineSymbols
    
    Set ws2 = Worksheets("RouteDraw")
    Set ws3 = Worksheets("EndPoints")
    ws3.Cells.Clear
    Data = ws2.UsedRange
    EpRow = 1

    For c1 = LBound(Data, 2) + 1 To UBound(Data, 2)
        For r1 = LBound(Data, 1) + 1 To UBound(Data, 1)
            If Data(r1, c1) = Sym.EP Then
                Call PrintEndPoint(r1, c1)
                Call IterateEndPoint(r1 + 1, c1, 0)
                Debug.Print CStr(r1) + ":" + CStr(c1)
            End If
        Next
    Next

'        Dim rg As String: rg = "M31"
'        Call IterateEndPoint(Range(rg).Row(), Range(rg).Column(), 0)
  
        Call StartIteration(Selection.row(), Selection.column())
End Sub
Private Sub StartIteration(r1 As Integer, c1 As Integer)
    
    EpCol = 1
    EpRow = EpRow + 1
    Call PrintEndPoint(r1, c1)
    
    Dim dir As Integer
    Dim pu As String
    pu = Data(r1 - 1, c1)
    
    Dim pd As String
    pd = IIf(r1 < UBound(Data, 1), Data(r1 + 1, c1), vbNullString)

    If pd <> vbNullString Then
        Call IterateEndPoint(r1 + 1, c1, 0)
    ElseIf pu <> vbNullString Then
        Call IterateEndPoint(r1 - 1, c1, 1)
    End If

End Sub

Private Sub PrintEndPoint(r1 As Integer, c1 As Integer)
    ws3.Cells(EpRow, EpCol) = "EP:" + CStr(Data(r1, c1 - 1))
    EpCol = EpCol + 1
End Sub

Private Sub IterateEndPoint(r1 As Integer, c1 As Integer, dr As Integer)
    
'    Call PrintEndPoint(r1, c1)
    If r1 > UBound(Data, 1) Then Exit Sub
    ws2.Cells(r1, c1).Interior.Color = vbCyan
    
    If Data(r1, c1) = Sym.EP Then
        Call PrintEndPoint(r1, c1 - 1)
        ws2.Cells(r1, c1).Interior.Color = vbYellow
        Exit Sub
    End If
    
    Dim pu As String
    pu = Data(r1 - 1, c1)
    
    Dim pd As String
    pd = IIf(r1 < UBound(Data, 1), Data(r1 + 1, c1), vbNullString)
    
    If dr = 0 Then
        If pd <> vbNullString Then
            If pd = Sym.Ver Then
               Call IterateEndPoint(r1 + 1, c1, dr)
            ElseIf pd = Sym.Tee Then
               Call IterateEndPoint(r1 + 1, c1, dr)
               Call FindElbowDown(r1 + 1, c1)
            ElseIf pd = Sym.EP Then
                Call PrintEndPoint(r1 + 1, c1)
            End If
        Else
            dr = 1 'Check
        End If
    Else
        If pu <> vbNullString Then
            If pu = Sym.Ver Then
                Call IterateEndPoint(r1 - 1, c1, dr)
            ElseIf pu = Sym.Tee Then
                Call IterateEndPoint(r1 - 1, c1, dr)
                Call FindElbowDown(r1 - 1, c1)
            ElseIf pu = Sym.Elb Then
                Call FindElbowUp(r1 - 1, c1)
            ElseIf pu = Sym.EP Then
                Call PrintEndPoint(r1 - 1, c1)
            End If
        End If
    End If
End Sub

Private Sub FindElbowDown(r1 As Integer, c1 As Integer)
    
    Dim col As Integer
    col = c1
    Do
        col = col + 1
        If Data(r1, col) = Sym.Elb Then
            Call IterateEndPoint(r1 + 1, col, 0)
            Exit Do
        ElseIf Data(r1, col) = Sym.Cross Then
            Call IterateEndPoint(r1 + 1, col, 0)
        ElseIf Data(r1, col) <> Sym.Hor Then
            Exit Do
        End If
        
    Loop
End Sub

Private Sub FindElbowUp(r1 As Integer, c1 As Integer)
    
    Dim col As Integer
    col = c1
    Do
        col = col - 1
        If Data(r1, col) = Sym.Tee Then
            Call IterateEndPoint(r1, col, 0)
            Call IterateEndPoint(r1, col, 1)
            Exit Do
        ElseIf Data(r1, col) = Sym.Cross Then
            Call IterateEndPoint(r1 + 1, col, 0)
        ElseIf Data(r1, col) <> Sym.Hor Then
            Exit Do
        End If
        
    Loop
End Sub

Private Sub getEndpoints()
    
    Set ws2 = Worksheets("RouteDraw")
    Set ws3 = Worksheets("EndPoints")
    
    DefineSymbols
        
    Const MAX_ROUTES = 256
    Dim routes(MAX_ROUTES) As Integer
    Dim routeCount As Integer: routeCount = 0
   
    Data = ws2.UsedRange
    Dim LB As Integer: LB = LBound(Data, 2)
    Dim UB As Integer: UB = UBound(Data, 2)
    
    Dim rowNum As Integer: rowNum = 2
    
    Dim n As Integer
    For n = LB To UB
        
        If Data(rowNum, n) <> vbNullString And IsNumeric(Data(rowNum, n)) Then
            ws2.Cells(rowNum, n + 1) = Sym.EP
            Call searchRoute(rowNum, n)
            routes(routeCount) = n
            routeCount = routeCount + 1
        End If
    Next
    
End Sub


Private Sub searchRoute(ByVal r1 As Integer, ByVal c1 As Integer)
    

    If r1 > UBound(Data, 1) Then
        ws2.Cells(r1 - 1, c1 + 1) = Sym.EP
        Exit Sub
    End If
    
    If Data(r1, c1) <> vbNullString Then
        Call searchRoute(r1 + 1, c1)
    
        If Data(r1, c1 + 1) = Sym.Tee Then
            Do
                c1 = c1 + 1
                If Data(r1, c1) = vbNullString Then Exit Do
                If Data(r1, c1) = Sym.Cross Then
                    Call searchRoute(r1 + 1, c1 - 1)
                ElseIf Data(r1, c1) = Sym.Elb Then
                    Call searchRoute(r1 + 1, c1 - 1)
                    Exit Do
                End If
            Loop
        End If
        
    Else
        ws2.Cells(r1 - 1, c1 + 1) = Sym.EP
    End If
    


End Sub
Private Sub DefineSymbols()
    
    Sym.Ver = Range("Ver")
    Sym.Hor = Range("Hor")
    Sym.Tee = Range("Tee")
    Sym.Elb = Range("Elb")
    Sym.Cross = Range("Cross")
    Sym.EP = Range("EP")

End Sub
