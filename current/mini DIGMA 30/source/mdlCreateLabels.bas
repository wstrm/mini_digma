Attribute VB_Name = "mdlCreateLabels"
'Public Sub CreateLabels()
'Dim vntData As Variant
'Dim intStartRow As Integer
'Dim intEndRow As Integer
'
'
'
''Skapa Kaplista
'intStartRow = GetStartRowLeft("KAPLISTA")
'intEndRow = GetEndRow(intStartRow)
'vntData = GetData("50649", 5, 1, 10, 7, 8, 9, intStartRow, intEndRow)
'
'
''Skapa Lager
'intStartRow = GetStartRowLeft("LAGER")
'intEndRow = GetEndRow(intStartRow)
'vntData = GetData("50649", 5, 1, 10, 7, 8, 9, intStartRow, intEndRow)
'
'
'
''Skapa Klipplista
'
'
'
'
''Skapa Plåtlager
'
'
'Call InsertData(vntData, wsKaplista)
'
'
''End Sub
'
''Hämtar startrad för vald grupp på utskrift vänster
'Private Function GetStartRowLeft(ByVal strGroup As String) As Integer
'    Dim i As Integer
'
'
'    intStartRow = 0
'    For i = 5 To 1000
'        If strGroup = Blad5.Cells(i, 5) Then
'            GetStartRowLeft = i + 4
'            Exit Function
'        End If
'    Next i
''
''
''End Function
'
''Hämtar startrad för vald grupp på utskrift Höger
'Private Function GetStartRowRight(ByVal strGroup As String) As Integer
'    Dim i As Integer
'    Dim intStartRow As Integer
'
'    intStartRow = 0
'    For i = 5 To 1000
'        If strGroup = Blad5.Cells(i, 18) Then
'            GetStartRowRight = i
'            Exit Function
'        End If
'    Next i
'
'    GetStartRow = intStartRow
'End Function
'
'Private Function GetEndRow(ByVal intStartRow As Integer) As Integer
'    Dim i As Integer
'    Dim intEndRow As Integer
'
'
'    For i = intStartRow To 1000
'        If Blad5.Cells(i, 5) = "" And Blad5.Cells(i, 5) = "" And Blad5.Cells(i, 10) = "" Then
'            GetEndRow = i - 1
'            Exit Function
'        End If
'    Next i
'
'
'End Function




Public Function GetLabelData(ByVal strorderNr, ByVal intBenamningCol As Integer, ByVal intRadCol As Integer, ByVal intAnmarkningCol As Integer, ByVal intLangdCol As Integer, _
ByVal intBreddCol As Integer, ByVal intTjockCol As Integer, ByVal intStartRow As Integer, ByVal intEndRow As Integer) As Variant

    Dim vntData() As Variant
    Dim i As Integer
    Dim intItem As Integer
    ReDim vntData(1 To 7, 1 To 1)

    intItem = 1
    For i = intStartRow To intEndRow
        ReDim Preserve vntData(1 To 7, 1 To intItem)
        vntData(1, intItem) = strorderNr
        vntData(2, intItem) = Blad5.Cells(i, intBenamningCol)
        vntData(3, intItem) = Blad5.Cells(i, intRadCol)
        vntData(4, intItem) = Blad5.Cells(i, intAnmarkningCol)
        vntData(5, intItem) = Blad5.Cells(i, intLangdCol)
        vntData(6, intItem) = Blad5.Cells(i, intBreddCol)
        vntData(7, intItem) = Blad5.Cells(i, intTjockCol)
        
        
        intItem = intItem + 1
    Next

    GetLabelData = vntData

End Function

'Private Function GetLager() As Variant
'
'
'End Function
'
'Private Function GetKlipplista() As Variant
'
'
'End Function
'
'Private Function GetPlatlager() As Variant
'
'
'End Function

Public Sub InsertLabelData(ByVal vntData As Variant, ws As Worksheet)
    Dim i As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    

    intCol = 2
    intRow = 2
    
    ws.Cells.ClearContents
    
    Call DeleteAllShape(ws)
    For i = 1 To UBound(vntData, 2)
    
        ws.Cells(intRow, intCol) = "Konstr. År:" & Chr(10) & Year(Date) & "-" & Month(Date)
        ws.Cells(intRow, intCol + 1) = vntData(1, i) & "-" & vntData(3, i)
        ws.Cells(intRow + 2, intCol) = vntData(2, i)
        ws.Cells(intRow + 3, intCol) = vntData(4, i)
        ws.Cells(intRow + 4, intCol) = "L:" & vntData(5, i) & " B:" & vntData(6, i) & " Tj:" & vntData(7, i)
        Blad1.Shapes("picLogga").Copy
        ws.Cells(intRow + 6, intCol).PasteSpecial
        ws.Cells(intRow + 6, intCol + 1) = "+46 (0)44 - 28 99 00"
        
        intCol = intCol + 4
        
        If intCol > 10 Then
            intCol = 2
            intRow = intRow + 8
        End If
        
    Next i
    
    'ws.Cells(1, 1).Select

End Sub

Private Sub DeleteAllShape(ws As Worksheet)

For Each sh In ws.Shapes
sh.Delete

Next


End Sub
