Attribute VB_Name = "MASTER"
Option Explicit
Dim OrdFil, OrderDir, OrdNum As String
Dim X, Y, Z, xa, xb, xc, xd, xe, xf, xg, xh, xi, xj As Integer
Dim TWX, TWZ, TUT As Object

    'DriveLetter = Left(miniDIGMAForm.OrderPath_Text.Text, 1)
    'ChDrive DriveLetter
    'ChDir twy.Range("Hemkatalog").Value
Public Sub Auto_Open()
    Application.ScreenUpdating = False
    Application.Goto Reference:=Range("A1"), scroll:=True
End Sub

Sub runMiniDIGMA()
Dim vntKapLista As Variant
Dim vntLager As Variant
Dim vntKlippLista As Variant
Dim vntPlatLager As Variant
Dim intEndRow As Integer
Dim intStartRow As Integer
Dim vntLabelData As Variant
Set TWX = ThisWorkbook.Sheets("x")
Set TUT = ThisWorkbook.Sheets("Utskrift")
OrdNum = miniDIGMAForm.OrderNummer_Text.Value

    With TWZ
        If miniDIGMAForm.OpenLoad_Check.Value = True Then GoTo kör    '1
        If miniDIGMAForm.PrintList_Check.Value = True Then GoTo kör    '3
        MsgBox "Du har inte valt någon åtgärd"
        End
kör:
        If miniDIGMAForm.OpenLoad_Check.Value = True Then
            TWX.Range("M:M").ClearContents
            Application.ScreenUpdating = False
            
            OpOrd
            Call SortByNumber
            If miniDIGMAForm.PrintList_Check.Value = True Then GoTo GåVidare
            End
        End If
GåVidare:
        If miniDIGMAForm.PrintList_Check.Value = True Then
           koptillutskrift
           
           'Sortera KapLista           intEndRow = 9
            intStartRow = 9
            intEndRow = 9
            Do While Blad5.Cells(intEndRow, 5) <> ""
                intEndRow = intEndRow + 1
            Loop
            If intEndRow > intStartRow Then

                
                'Hämta data till utskriftsblad
                vntKapLista = Blad5.Range(Cells(intStartRow, 1), Cells(intEndRow - 1, 10))
    
                vntKapLista = Transpose(vntKapLista, 1)
                Call MyQuickSort_Quad(vntKapLista, 1, UBound(vntKapLista, 2), 6, 9, 7, 8, 5, False)
                vntKapLista = Transpose(vntKapLista, 1)
                Blad5.Range(Cells(intStartRow, 1), Cells(intEndRow - 1, 10)) = vntKapLista
                
                Call ChangeNr(1)
                'Hämta data till Etikettblad
                vntLabelData = mdlCreateLabels.GetLabelData(OrdNum, 5, 1, 10, 7, 8, 9, intStartRow, intEndRow - 1)
                'Skriv till Etikettblad
                Call mdlCreateLabels.InsertLabelData(vntLabelData, wsKaplista)
                
            End If
           
           'Sortera Lager
            intEndRow = intEndRow + 6
            intStartRow = intEndRow

            Do While Blad5.Cells(intEndRow, 5) <> ""
                intEndRow = intEndRow + 1
            Loop

            If intEndRow > intStartRow Then
                vntKapLista = Blad5.Range(Cells(intStartRow, 1), Blad5.Cells(intEndRow - 1, 10))

                vntKapLista = Transpose(vntKapLista, 1)
                Call MyQuickSort_Quad(vntKapLista, 1, UBound(vntKapLista, 2), 6, 9, 7, 8, 5, False)
                vntKapLista = Transpose(vntKapLista, 1)
                Blad5.Range(Cells(intStartRow, 1), Cells(intEndRow - 1, 10)) = vntKapLista
                
                Call ChangeNr(1)
                'Hämta data till Etikettblad
                vntLabelData = mdlCreateLabels.GetLabelData(OrdNum, 5, 1, 10, 7, 8, 9, intStartRow, intEndRow - 1)
                'Skriv till Etikettblad
                Call mdlCreateLabels.InsertLabelData(vntLabelData, wsLager)
                
            End If

           'Sortera Klipplista
            intEndRow = 9
            intStartRow = 9

            Do While Blad5.Cells(intEndRow, 18) <> ""
                intEndRow = intEndRow + 1
            Loop

            If intEndRow > intStartRow Then
                vntKapLista = Blad5.Range(Cells(intStartRow, 14), Cells(intEndRow - 1, 26))

                vntKapLista = Transpose(vntKapLista, 1)
                Call MyQuickSort_Quad(vntKapLista, 1, UBound(vntKapLista, 2), 6, 10, 7, 8, 5, False)
                vntKapLista = Transpose(vntKapLista, 1)
                Blad5.Range(Cells(intStartRow, 14), Cells(intEndRow - 1, 26)) = vntKapLista
                
                Call ChangeNr(14)
                'Hämta data till Etikettblad
                vntLabelData = mdlCreateLabels.GetLabelData(OrdNum, 18, 14, 24, 25, 26, 23, intStartRow, intEndRow - 1)
                'Skriv till Etikettblad
                Call mdlCreateLabels.InsertLabelData(vntLabelData, wsKlipplista)

            End If
           
           'sortera plåtlager
            intEndRow = intEndRow + 6
            intStartRow = intEndRow

            Do While Blad5.Cells(intEndRow, 18) <> ""
                intEndRow = intEndRow + 1
            Loop
                
                If intEndRow > intStartRow Then
                    vntKapLista = Blad5.Range(Cells(intStartRow, 14), Cells(intEndRow - 1, 26))

                    vntKapLista = Transpose(vntKapLista, 1)
                    Call MyQuickSort_Quad(vntKapLista, 1, UBound(vntKapLista, 2), 6, 10, 7, 8, 5, False)
                    vntKapLista = Transpose(vntKapLista, 1)
                    Blad5.Range(Cells(intStartRow, 14), Cells(intEndRow - 1, 26)) = vntKapLista
                    
                Call ChangeNr(14)
                'Hämta data till Etikettblad
                vntLabelData = mdlCreateLabels.GetLabelData(OrdNum, 18, 14, 24, 25, 26, 23, intStartRow, intEndRow - 1)
                'Skriv till Etikettblad
                Call mdlCreateLabels.InsertLabelData(vntLabelData, wsPlatlager)
                End If
            
            'Loopar alla celler i vald kolumn och gör ord som är längre än vald längd kortare
            'för att få plats
            Call CreateShortWord1(5, 10)
            Call CreateShortWord2(18, 24)
            
            Call SplitRow(5, 1, 13, 10)
            Call SplitRow2(18, 14, 26, 24)

            Call ChangeNr(1)
            Call ChangeNr(14)
            
            Blad5.Range(Cells(9, 23), Cells(100, 23)).NumberFormat = "@"
            Blad5.Range(Cells(9, 9), Cells(100, 9)).NumberFormat = "@"
            miniDIGMAForm.Status_Label.Caption = "Klar"
        End If
    End With
    
End Sub

Public Sub CreateShortWord1(ByVal intCol1 As Integer, ByVal intCol12 As Integer)
    Dim j As Integer
    Dim i As Integer
    Dim tmpWord As String
    Dim astrWord() As String
    Dim intShortWord1 As Integer
    Dim intShortWord2 As Integer
    
    intShortWord1 = 22
    intShortWord2 = 29
    
    For i = 1 To 100
        astrWord = Split(Blad5.Cells(i, intCol1), " ")
        For j = 0 To UBound(astrWord)
            If Len(astrWord(j)) > intShortWord1 Then
                astrWord(j) = Left(astrWord(j), intShortWord1)
            End If
        tmpWord = tmpWord & " " & astrWord(j)
        Next j
        
        Blad5.Cells(i, intCol1) = Trim(tmpWord)
        tmpWord = ""
    Next i
    
   For i = 1 To 100
        astrWord = Split(Blad5.Cells(i, intCol12), " ")
        For j = 0 To UBound(astrWord)
            If Len(astrWord(j)) > intShortWord2 Then
                astrWord(j) = Left(astrWord(j), intShortWord2)
            End If
        tmpWord = tmpWord & " " & astrWord(j)
        Next j
        Blad5.Cells(i, intCol12) = Trim(tmpWord)
        tmpWord = ""
    Next i

End Sub

Public Sub CreateShortWord2(ByVal intCol1 As Integer, ByVal intCol12 As Integer)
    Dim j As Integer
    Dim i As Integer
    Dim tmpWord As String
    Dim astrWord() As String
    Dim intShortWord1 As Integer
    Dim intShortWord2 As Integer
    
    intShortWord1 = 12
    intShortWord2 = 17
    
    For i = 1 To 100
        astrWord = Split(Blad5.Cells(i, intCol1), " ")
        For j = 0 To UBound(astrWord)
            If Len(astrWord(j)) > intShortWord1 Then
                astrWord(j) = Left(astrWord(j), intShortWord1)
            End If
        tmpWord = tmpWord & " " & astrWord(j)
        Next j
        Blad5.Cells(i, intCol1) = Trim(tmpWord)
        tmpWord = ""
    Next i
    
   For i = 1 To 100
        astrWord = Split(Blad5.Cells(i, intCol12), " ")
        For j = 0 To UBound(astrWord)
            If Len(astrWord(j)) > intShortWord2 Then
                astrWord(j) = Left(astrWord(j), intShortWord2)
            End If
        tmpWord = tmpWord & " " & astrWord(j)
        Next j
        Blad5.Cells(i, intCol12) = Trim(tmpWord)
        tmpWord = ""
    Next i

End Sub

Sub OpOrd()
    TWX.Unprotect Password:="ki"
    TWX.Range("j:BB").ClearContents
    OrderDir = miniDIGMAForm.OrderPath_Text.Text
    OrdFil = miniDIGMAForm.OrderNummer_Text.Value
    Workbooks.Open Filename:=OrderDir & "\" & OrdFil & "\" & OrdFil & ".xls"
    Columns("A:n").Copy
    ThisWorkbook.Activate
    Sheets("x").Activate
    Range("alfa").Select
    Selection.PasteSpecial Paste:=xlValues: Application.CutCopyMode = False
    Workbooks(OrdFil & ".xls").Close SaveChanges:=False
    Columns("k:x").Copy Range("ae1")
End Sub

Sub SortByNumber()
    With ThisWorkbook
        With Sheets("x")
            .Unprotect Password:="ki"
                Range("ae2:Ar55555").Select
                Selection.Sort Key1:=Range("ar2"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
                Columns("AE:ar").Copy Range("k1")
                Range("K2:AB55555").Select
            .Protect Password:="ki"
        End With
    End With
End Sub

Sub CreateMenu()
'   This sub should be executed when the workbook is opened.
'   NOTE: There is no error handling in this subroutine

    Dim MenuSheet As Worksheet
    Dim MenuObject As CommandBarPopup

    Dim MenuItem As Object
    Dim SubMenuItem As CommandBarButton
    Dim Row As Integer
    Dim MenuLevel, NextLevel, PositionOrMacro, Caption, Divider, FaceId

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Location for menu data
    Set MenuSheet = ThisWorkbook.Sheets("M")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''

    '   Make sure the menus aren't duplicated
    Call DeleteMenu

    '   Initialize the row counter
    Row = 2

    '   Add the menus, menu items and submenu items using
    '   data stored on MenuSheet

    Do Until IsEmpty(MenuSheet.Cells(Row, 1))
        With MenuSheet
            MenuLevel = .Cells(Row, 1)
            Caption = .Cells(Row, 2)
            PositionOrMacro = .Cells(Row, 3)
            Divider = .Cells(Row, 4)
            'FaceId = .Cells(Row, 5)
            NextLevel = .Cells(Row + 1, 1)
        End With

        Select Case MenuLevel
        Case 1    ' A Menu
            '              Add the top-level menu to the Worksheet CommandBar
            Set MenuObject = Application.CommandBars(1). _
                             Controls.Add(Type:=msoControlPopup, _
                                          Before:=PositionOrMacro, _
                                          temporary:=True)
            MenuObject.Caption = Caption

        Case 2    ' A Menu Item
            If NextLevel = 3 Then
                Set MenuItem = MenuObject.Controls.Add(Type:=msoControlPopup)
            Else
                Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
                MenuItem.OnAction = PositionOrMacro
            End If
            MenuItem.Caption = Caption
            'If FaceId <> "" Then MenuItem.FaceId = FaceId
            If Divider Then MenuItem.BeginGroup = True

        Case 3    ' A SubMenu Item
            Set SubMenuItem = MenuItem.Controls.Add(Type:=msoControlButton)
            SubMenuItem.Caption = Caption
            SubMenuItem.OnAction = PositionOrMacro
            'If FaceId <> "" Then SubMenuItem.FaceId = FaceId
            If Divider Then SubMenuItem.BeginGroup = True
        End Select
        Row = Row + 1
    Loop
End Sub
Sub DeleteMenu()
    Dim MenuSheet As Worksheet
    Dim Row As Integer
    Dim Caption As String
    Set MenuSheet = ThisWorkbook.Sheets("M")
    Row = 2
    Do Until IsEmpty(MenuSheet.Cells(Row, 1))
        If MenuSheet.Cells(Row, 1) = 1 Then
            Caption = MenuSheet.Cells(Row, 2)
            uslo Caption
        End If
        Row = Row + 1
    Loop
End Sub
Sub uslo(Caption)
    For a = CommandBars(1).Controls().Count To 1 Step -1
        If CommandBars(1).Controls(a).Caption = Caption Then
            Application.CommandBars(1).Controls(Caption).Delete
        End If
    Next
End Sub
Sub koptillutskrift()
    TUT.Unprotect Password:="ki": Range(TUT.Cells(9, 1), TUT.Cells(333, 13)).Clear: Range(TUT.Cells(10, 14), TUT.Cells(333, 33)).Clear: TWX.Activate
    '-------------------------------------------------------------------------sid 1
    If Range("msk").Offset(0, 1).Value = 1 Then
        xa = 2    'msk börjar på rad 2
        xb = Range("msk").Offset(0, 3).Value    'msk omfattar xb rader
    End If
    '-------------------------------------------------------------------------sid 1
    If Range("msklag").Offset(0, 1).Value = 1 Then
        xc = Range("msklag").Offset(0, 2).Value    'msklag börjarp å rad xc
        xd = Range("msklag").Offset(0, 3).Value    'msklag omfattar xd rader
    End If
    '---------------------------------------------------------------------------------sid 2
    If Range("plåt").Offset(0, 1).Value = 1 Then
        xe = Range("plåt").Offset(0, 2).Value    'plåt börjar pårad xe
        xf = Range("plåt").Offset(0, 3).Value    'plåt omfattar xf rader
    End If
    '--------------------------------------------------------------------------------sid 2
    If Range("plåtlag").Offset(0, 1).Value = 1 Then
        xg = Range("plåtlag").Offset(0, 2).Value    'plåtlag börjarpårad xg
        xh = Range("plåtlag").Offset(0, 3).Value    'plåtlag omfattar xh rader
    End If
    '-----------------------------------------------------------------------
    xi = Range("ömått").Offset(0, 2).Value    'ömått = sista raden
    '--------------------------------------------------------------------------
    'RW tillägg-------------------------------
    'Sheets("Utskrift").Select
    'Range("A8:M8").Select
    'Selection.Copy
    'Range("N8").Select
    'ActiveSheet.Paste
    '---------------------------------------------
    '------------------------------------------------------------------------sid 1 börjar
    If Range("msk").Offset(0, 1).Value = 1 Then
        Range(TWX.Cells(xa, 11), TWX.Cells(xa + xb - 1, 23)).Copy TUT.Cells(9, 1)    'msk
    End If

    If Range("msklag").Offset(0, 1).Value = 1 Then
        Range(TUT.Cells(7, 1), TUT.Cells(8, 10)).Copy TUT.Cells(9 + xb + 1, 1)
        TUT.Cells(9 + xb + 1, 5).Value = "LAGER"
        Range(TWX.Cells(xc, 11), TWX.Cells(xc + xd - 1, 23)).Copy TUT.Cells(xc + xa + 8, 1)    'msklag
    Else
        Range(TUT.Cells(7, 1), TUT.Cells(8, 10)).Copy TUT.Cells(9 + xb + 1, 1)
        TUT.Cells(9 + xb + 1, 5).Value = "LAGER"
    End If
    '-------------------------------------------------------------------sid 1 slut

    '-------------------------------------------------------------------sid 2 börjar
    If Range("plåt").Offset(0, 1).Value = 1 Then
        Range(TWX.Cells(xe, 11), TWX.Cells(xe + xf - 1, 23)).Copy TUT.Cells(10, 14)    'plåt
    End If

    If Range("plåtlag").Offset(0, 1).Value = 1 Then
        Range(TUT.Cells(7, 14), TUT.Cells(9, 26)).Copy TUT.Cells(10 + xf + 1, 14)
        TUT.Cells(10 + xf + 1, 18).Value = "PLÅTLAGER"
        Range(TWX.Cells(xg, 11), TWX.Cells(xg + xh - 1, 23)).Copy TUT.Cells(10 + xf + 4, 14)    'plåtlag
    Else
        Range(TUT.Cells(7, 14), TUT.Cells(9, 26)).Copy TUT.Cells(10 + xf + 1, 14)
        TUT.Cells(10 + xf + 1, 18).Value = "PLÅTLAGER"
    End If
    '-------------------------------------------------------------------sid 2 slut
    'RW tillägg-------------------------------
    Sheets("Utskrift").Select
    Columns("X:Z").Select
    Selection.NumberFormat = "0" ' tar bort decimaler på klippmått
    'Application.CutCopyMode = False
    'Selection.Cut
    'Columns("T:T").Select
    'Selection.Insert Shift:=xlToRight
    'Columns("W:X").Select
    'Selection.Cut
    'Columns("AA:AA").Select
    ''Selection.Insert Shift:=xlToRight
    'Columns("T:T").ColumnWidth = 7
    'Columns("U:U").ColumnWidth = 7
    'Columns("V:V").ColumnWidth = 7
    'Columns("W:W").ColumnWidth = 3.5
    'Columns("X:X").ColumnWidth = 18
    'Columns("Y:Y").ColumnWidth = 7
    'Columns("Z:Z").ColumnWidth = 7

    '---------------------------------------------

    TUT.Activate
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Goto Reference:=Range("A1"), scroll:=True
End Sub
'msk     2       6     msk första rad är på rad 2 och omfattar 6 rader dvs slutar på rad 7
'msklag  8       20    msklag första rad är på rad 8 och omfattar 20 rader dvs slutar på rad 27
'plåt    28      14    plåt första rad är på rad 28 och omfattar 14 rader dvs slutar på rad 41
'plåtlag 42      1     plåtlag första rad är på rad 42 och omfattar 1 rader dvs slutar på rad 42
'ömått   43            ömått är sista raden

'Loop för att dela upp en sträng på fler rader
Public Sub SplitRow(ByVal intSearchCol As Integer, ByVal FirstCol As Integer, ByVal LastCol As Integer, ByVal intSearchCol2)
    Dim i As Integer
    Dim strLargeText As String
    Dim intWordLenght As Integer
    Dim strLargeText2 As String     'Ordets längd
    Dim intWordLenght2 As Integer   'Ordets längd
    Dim intLenght1 As Integer       'Max antal tecken i kolumn
    Dim intLenght2 As Integer       'Max antal tecken i kolumn
    Dim blnCol2Update As Boolean
    Dim blnCol1Update As Boolean
    
    
    intLenght1 = 23
    intLenght2 = 30
    
      
    For i = 9 To 100
        blnCol1Update = False
        blnCol2Update = False
        If Len(Blad5.Cells(i, intSearchCol)) > intLenght1 Or Len(Blad5.Cells(i, intSearchCol2)) > intLenght2 Then
            strLargeText = Blad5.Cells(i, intSearchCol)
            strLargeText2 = Blad5.Cells(i, intSearchCol2)
            'Skriver ut allt innan sista mellanslag
            If Len(strLargeText) > intLenght1 Then
                Blad5.Cells(i, intSearchCol) = Left(strLargeText, InStrRev(Mid(strLargeText, 1, intLenght1), " "))
                intWordLenght = Len(Left(strLargeText, InStrRev(Mid(strLargeText, 1, intLenght1), " ")))
                strLargeText = Mid(strLargeText, intWordLenght + 1)
                blnCol1Update = True
            End If
            If Len(strLargeText2) > intLenght2 Then
                Blad5.Cells(i, intSearchCol2) = Left(strLargeText2, InStrRev(Mid(strLargeText2, 1, intLenght2), " "))
                intWordLenght2 = Len(Left(strLargeText2, InStrRev(Mid(strLargeText2, 1, intLenght2), " ")))
                strLargeText2 = Mid(strLargeText2, intWordLenght2 + 1)
                blnCol2Update = True
            End If
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Insert Shift:=xlDown
            
            If blnCol1Update = True Then
                Blad5.Cells(i + 1, intSearchCol) = strLargeText
            End If
            
            If blnCol2Update = True Then
                Blad5.Cells(i + 1, intSearchCol2) = strLargeText2
            End If
            
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeBottom).Weight = xlHairline

        End If
    Next i
End Sub

'Loop för att dela upp en sträng på fler rader i listan till vänster
Public Sub SplitRow2(ByVal intSearchCol As Integer, ByVal FirstCol As Integer, ByVal LastCol As Integer, ByVal intSearchCol2)
    Dim i As Integer
    Dim strLargeText As String
    Dim intWordLenght As Integer
    Dim strLargeText2 As String     'Ordets längd
    Dim intWordLenght2 As Integer   'Ordets längd
    Dim intLenght1 As Integer       'Max antal tecken i kolumn
    Dim intLenght2 As Integer       'Max antal tecken i kolumn
    Dim blnCol2Update As Boolean
    Dim blnCol1Update As Boolean
    
    
    intLenght1 = 13
    intLenght2 = 18
    
      
    For i = 9 To 100
        blnCol1Update = False
        blnCol2Update = False
        If Len(Blad5.Cells(i, intSearchCol)) > intLenght1 Or Len(Blad5.Cells(i, intSearchCol2)) > intLenght2 Then
            strLargeText = Blad5.Cells(i, intSearchCol)
            strLargeText2 = Blad5.Cells(i, intSearchCol2)
            'Skriver ut allt innan sista mellanslag
            If Len(strLargeText) > intLenght1 Then
                Blad5.Cells(i, intSearchCol) = Left(strLargeText, InStrRev(Mid(strLargeText, 1, intLenght1), " "))
                intWordLenght = Len(Left(strLargeText, InStrRev(Mid(strLargeText, 1, intLenght1), " ")))
                strLargeText = Mid(strLargeText, intWordLenght + 1)
                blnCol1Update = True
            End If
            If Len(strLargeText2) > intLenght2 Then
                Blad5.Cells(i, intSearchCol2) = Left(strLargeText2, InStrRev(Mid(strLargeText2, 1, intLenght2), " "))
                intWordLenght2 = Len(Left(strLargeText2, InStrRev(Mid(strLargeText2, 1, intLenght2), " ")))
                strLargeText2 = Mid(strLargeText2, intWordLenght2 + 1)
                blnCol2Update = True
            End If
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Insert Shift:=xlDown
            
            If blnCol1Update = True Then
                Blad5.Cells(i + 1, intSearchCol) = strLargeText
            End If
            
            If blnCol2Update = True Then
                Blad5.Cells(i + 1, intSearchCol2) = strLargeText2
            End If
            
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeTop).LineStyle = xlContinuous
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Blad5.Range(Cells(i + 1, FirstCol), Cells(i + 1, LastCol)).Borders(xlEdgeBottom).Weight = xlHairline

        End If
    Next i
End Sub

Public Sub ChangeNr(ByVal intCol As Integer)
Dim i As Integer
Dim intNr As Integer
Dim vntText As Variant

intNr = 1
For i = 9 To 100
    vntText = Blad5.Cells(i, intCol)
    If IsNumeric(vntText) = True And vntText <> "" Then
        Blad5.Cells(i, intCol) = intNr
        intNr = intNr + 1
    End If


Next i


End Sub

' Command which calls quicksort routine below and sorts 3rd, 1st, 2nd, and 4th dimensions
'MyQuickSort_Single aryTest(), 1, UBound(aryTest(), 2), 3, 1, 2, 4, True


' The called quicksort subroutine (array passed By Reference -- so array itself is modified
Public Sub MyQuickSort_Quad(ByRef SortArray As Variant, _
            ByVal First As Long, ByVal Last As Long, _
            ByVal PrimeSort As String, ByVal SecSort As Single, _
            ByVal TriSort As Single, ByVal QuadSort As Single, _
            ByVal QuintSort As Long, _
            Optional ByVal Ascending As Boolean = True)
Dim Low As Long
Dim High As Long
Dim Temp As Variant
Dim List_Separator1 As Variant
Dim List_Separator2 As Variant
Dim List_Separator3 As Variant
Dim List_Separator4 As Variant
Dim List_Separator5 As Variant
Dim TempArray() As Variant
Dim i
ReDim TempArray(UBound(SortArray, 1))

Low = First
High = Last
List_Separator1 = SortArray(PrimeSort, (First + Last) / 2)
List_Separator2 = SortArray(SecSort, (First + Last) / 2)
List_Separator3 = SortArray(TriSort, (First + Last) / 2)
List_Separator4 = SortArray(QuadSort, (First + Last) / 2)
List_Separator5 = SortArray(QuintSort, (First + Last) / 2)
Do
    If Ascending = True Then
        Do While (SortArray(PrimeSort, Low) < List_Separator1) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) < List_Separator2)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) < List_Separator3)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) = List_Separator3) And (SortArray(QuadSort, Low) < List_Separator4)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) = List_Separator3) And (SortArray(QuadSort, Low) = List_Separator4) And _
            (SortArray(QuintSort, Low) > List_Separator5))
            Low = Low + 1
        Loop
        Do While (SortArray(PrimeSort, High) > List_Separator1) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) > List_Separator2)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) > List_Separator3)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) = List_Separator3) And (SortArray(QuadSort, High) > List_Separator4)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) = List_Separator3) And (SortArray(QuadSort, High) = List_Separator4) And _
            (SortArray(QuintSort, High) < List_Separator5))
            High = High - 1
        Loop
    Else
        Do While (SortArray(PrimeSort, Low) > List_Separator1) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) > List_Separator2)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) > List_Separator3)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) = List_Separator3) And (SortArray(QuadSort, Low) > List_Separator4)) Or _
            ((SortArray(PrimeSort, Low) = List_Separator1) And (SortArray(SecSort, Low) = List_Separator2) And _
            (SortArray(TriSort, Low) = List_Separator3) And (SortArray(QuadSort, Low) = List_Separator4) And _
            (SortArray(QuintSort, Low) < List_Separator5))
            Low = Low + 1
        Loop
        Do While (SortArray(PrimeSort, High) < List_Separator1) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) < List_Separator2)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) < List_Separator3)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) = List_Separator3) And (SortArray(QuadSort, High) < List_Separator4)) Or _
            ((SortArray(PrimeSort, High) = List_Separator1) And (SortArray(SecSort, High) = List_Separator2) And _
            (SortArray(TriSort, High) = List_Separator3) And (SortArray(QuadSort, High) = List_Separator4) And _
            (SortArray(QuintSort, High) > List_Separator5))
            High = High - 1
        Loop
    End If
    If (Low <= High) Then
        For i = LBound(SortArray, 1) To UBound(SortArray, 1)        ' Lower bounds indicates lowest dimension
            TempArray(i) = SortArray(i, Low)
        Next
        For i = LBound(SortArray, 1) To UBound(SortArray, 1)
            SortArray(i, Low) = SortArray(i, High)
        Next
        For i = LBound(SortArray, 1) To UBound(SortArray, 1)
            SortArray(i, High) = TempArray(i)
        Next
        Low = Low + 1
        High = High - 1
    End If
Loop While (Low <= High)
If (First < High) Then MyQuickSort_Quad SortArray, First, High, PrimeSort, SecSort, TriSort, QuadSort, QuintSort, Ascending
If (Low < Last) Then MyQuickSort_Quad SortArray, Low, Last, PrimeSort, SecSort, TriSort, QuadSort, QuintSort, Ascending
End Sub
Public Function Transpose(ByVal vavntArrayIn As Variant, Optional ByVal vlngFirst As Long = 0) As Variant
    Dim avntTemp() As Variant
    Dim i As Long
    Dim j As Long
    
    'Vänd på arrayen
    If (vlngFirst > 0) Then
        ReDim avntTemp(vlngFirst To UBound(vavntArrayIn, 2), vlngFirst To UBound(vavntArrayIn, 1))
    Else
        ReDim avntTemp(UBound(vavntArrayIn, 2), UBound(vavntArrayIn, 1))
    End If
    For i = LBound(vavntArrayIn, 2) To UBound(vavntArrayIn, 2)
        For j = LBound(vavntArrayIn, 1) To UBound(vavntArrayIn, 1)
            avntTemp(i, j) = vavntArrayIn(j, i)
        Next j
    Next i
    Transpose = avntTemp
End Function
