Attribute VB_Name = "Träkorg"
Dim filFOR As String
Dim filBAK As String
Dim sheetFOR As Worksheet
Dim sheetBAK As Worksheet
Dim sheetMERGE As Worksheet
Dim sheetARTIKEL As Worksheet
Dim sheetTEMP As Worksheet

Sub ÖppnaFOR()
Attribute ÖppnaFOR.VB_Description = "Öppna .FOR-fil för att sedan visas i Excel"
Attribute ÖppnaFOR.VB_ProcData.VB_Invoke_Func = " \n14"

' ---ÖppnaFOR---
' William Wennerström - 2013/06/19
' "Öppnar .FOR-fil för att sedan visas i Excel"

    On Error GoTo Errmsg

 'Göm fönster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.AutoFilterMode = False
    
'Radera gamla blad, ifall de finns kvar av någon anledning
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("TEMP").Delete
    Sheets(".FOR").Delete
    Sheets("KAPNOTA").Delete
    Sheets("KONVERTERA").Delete
    Application.DisplayAlerts = True
    
'Lägg till nya blad
    Sheets.Add.Name = "TEMP"
    Sheets.Add.Name = "KAPNOTA"
    Sheets.Add.Name = ".FOR"
    Sheets.Add.Name = "KONVERTERA"
    
 'Bestäm variabler
    Set sheetFOR = Sheets(".FOR")
    Set sheetBAK = Sheets("KAPNOTA")
    Set sheetMERGE = Sheets("KONVERTERA")
    Set sheetARTIKEL = Sheets("ARTIKELREG")
    Set sheetTEMP = Sheets("TEMP")
    filFOR = "L:\AM\PRO\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value & ".FOR"
    filBAK = "L:\AM\PRO\KAPNOTA\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    
'Göm blad
    sheetARTIKEL.Visible = False
    
'Formatera blad
    'sheetFOR.Select
    'Columns("J:J").Select
    'Selection.NumberFormat = "@"
    'sheetMERGE.Select
    'Columns("C:C").Select
    'Selection.NumberFormat = "@"
    'sheetTEMP.Select
    'Columns("A:A").Select
    'Selection.NumberFormat = "@"
    
'Leta efter .FOR-fil i L:\AM\PRO för att sedan öppna den och förbereder den för Excel.
    sheetFOR.Select
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filFOR, _
        Destination:=Range("'.FOR'!$A$1"))
        .Name = "FORFIL"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(9, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(3, 7, 4, 5, 5, 23, 28, 5)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    On Error GoTo Errmsg2
    
    'Kör MergeBAK_FOR
    Call MergeBAK_FOR
    
Exit Sub

Errmsg:
   MsgBox ("Misslyckades att öppna '" & filFOR & "', se till att den finns eller inte är öppen."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
Errmsg2:
    MsgBox ("Misslyckades att köra ÖppnaBAK"), vbOKOnly, ".BAK-fil"
    Application.ScreenUpdating = True
End Sub
Sub MergeBAK_FOR()
Attribute MergeBAK_FOR.VB_Description = "Makrot inspelat 2013-06-19 av mattias"
Attribute MergeBAK_FOR.VB_ProcData.VB_Invoke_Func = " \n14"

' ---MergeBAK_FOR---
' William Wennerström - 2013/06/20 | Uppdaterad 2013-06-25 & 2013-06-26
' "Sammansätter BAK med FOR-filen."

    'Bestäm variabler
    Dim filBakNoDir As String
    filFOR = "L:\AM\PRO\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value & ".FOR"
    filBAK = "L:\AM\PRO\KAPNOTA\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    filBakNoDir = Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    
    
    On Error GoTo Errmsg
    
    sheetBAK.Activate
With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filBAK, _
        Destination:=Range("$A$1"))
        .Name = Sheets("PROGRAMÖVERSIKT").TextBox1.Value & ".BAK"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(9, 1, 9, 1, 9)
        .TextFileFixedColumnWidths = Array(14, 15, 11, 18)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    On Error GoTo Errmsg2
    
    'Merge FOR & BAK
    sheetFOR.Activate
    
    
    'Lägg in 27st rader överst.
    ActiveSheet.[1:27].Insert Shift:=xlDown
    
    'Kopiera från BAK till FOR
    sheetBAK.Activate
    Range("A2,A3,A4,A6,A9,A10:A13").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B1").Select
    ActiveSheet.Paste
    sheetBAK.Activate
    Range("B2,B3,B4,B6,B9,B10").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B12").Select
    ActiveSheet.Paste
   
    'Stäng BAK och öppna igen med ny FieldInfo
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(15 _
        , 1), Array(40, 1))
        
    'Kopiera från BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B5,B7,B8").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B18").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    
    'Stäng BAK och öppna igen med ny FieldInfo
    ActiveWindow.Close False
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(36 _
        , 1), Array(58, 1))
        
    'Kopiera in från BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B13,B12,B11").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B21").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    
    'Stäng BAK och öppna igen med ny FieldInfo
    ActiveWindow.Close False
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(15 _
        , 1), Array(59, 1))
        
    'Kopiera in från BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B14,B15,B16").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B24").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    ActiveWindow.Close False
    
    'Lägg in information till "PROGRAMÖVERSIKT"
    Sheets("Meny").[B12] = Sheets(".FOR").[B13]
    Sheets("Meny").[B13] = Sheets(".FOR").[B15]
    
    'Välj B30:E?? för att sedan konvertera nummer & text
    sheetFOR.Activate
    With ActiveSheet
    Range(.Range("A30"), ("D" & .UsedRange.Rows.Count)).Select
    End With
    
    'Konvertera punkt till komma
    Const sCOMMA = ","
    Const sDOT = "."
    With Selection
            .Replace sDOT, sCOMMA, xlPart
    End With
    
    'Behåll endast siffror och komma
    On Error Resume Next
    Dim RE As Object
    Dim rng As Range
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Global = True
        'Tillåt "," och alla nummer mellan 0-9
       .Pattern = "[^,0-9]"
       For Each rng In Selection
            rng.Value = .Replace(rng.Value, "")
        Next rng
    End With
    
    'Läs som nummer istället för text
    'Blir ett error här, använder därav Resume Next
    Dim xCell As Range
    On Error Resume Next
    For Each xCell In Selection
    xCell.Value = CDec(xCell.Value)
    Next xCell
    
    'Radera alla ogiltiga artiklar (alla artiklar utan ett rad nummer.)
    sheetFOR.Activate
    On Error Resume Next
    Range(("G31"), ("G" & ActiveSheet.UsedRange.Rows.Count)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    'Slut
    'On Error GoTo Errmsg2
    Call RWartikel_Sök
    Exit Sub
    

Errmsg:
   MsgBox ("Misslyckades att öppna '" & filBAK & "', se till att den finns eller inte är öppen."), vbOKOnly, ".BAK-fil"
    Application.ScreenUpdating = True
    Exit Sub
Errmsg2:
   MsgBox ("Misslyckades att köra sammansättning av '" & filBAK & "' och '" & filFOR & "'."), vbOKOnly, "Sammansättning"
    Application.ScreenUpdating = True
End Sub

Sub RWartikel_Sök()
    
' ---RWartikel_Sök---
' William Wennerström - 2013/06/25
' "Öppnar RWartikel.xls, hämtar sedan resurs till .FOR-fil."
    
    On Error GoTo Errmsg
    
    'Deklarera variabler
    Dim strLastRow As String
    Dim rngC As Range
    Dim strToFind As String, FirstAddress As String, FirstRow As String, xCellRow As String, artNr As String
    Dim rngtest As String
    Dim selRng, xCell, rngFind As Range
    Set sheetFOR = Sheets(".FOR")
    Set sheetBAK = Sheets("KAPNOTA")
    Set sheetMERGE = Sheets("KONVERTERA")
    Set sheetARTIKEL = Sheets("ARTIKELREG")
    Set sheetTEMP = Sheets("TEMP")
    
    'Välj alla artiklar
    
    sheetFOR.Activate
    With ActiveSheet
    Range(.Range("E30"), ("E" & .UsedRange.Rows.Count)).Select
    Set selRng = Selection
    End With
    
    'Bestäm artikel att leta efter
    For Each xCell In selRng
    
    'Bestäm vilken artikel som ska hittas
    strToFind = xCell.Value
    xCellRow = xCell.Row
    
    'Sök i RWartikel och kopiera resurs för varje artikel
    sheetARTIKEL.Activate
    With ActiveSheet.Range("C1:C700")
        Set rngC = .Find(What:=strToFind, LookAt:=xlWhole)
            If Not rngC Is Nothing Then
                FirstAddress = rngC.Address
                'FirstRow = rngC.Row
                Do
                    'MsgBox ("Range: " & rngC.Address & " Row: " & FirstRow & " Art. nr: " & xCell.Value)
                    Range(rngC.Address).Select
                    ActiveCell.Offset(0, 12).Select
                    ActiveCell.Resize(, 2).Select
                    Selection.Copy
                    Set rngC = .FindNext(rngC)
                Loop While Not rngC Is Nothing And rngC.Address <> FirstAddress
            End If
    End With
    
    'Klistra in varje resurs i temp fil
    sheetTEMP.Activate
    Range("C1").Value = Range("C1").Value + 1
    Range("A" & (Range("C1").Value), ("B" & (Range("C1").Value))).PasteSpecial
    
    Next xCell
    
    On Error GoTo Errmsg2
    
    'Välj och ta bort all format på celler
    With ActiveSheet
    Range(.Range("A1"), ("B" & .UsedRange.Rows.Count)).Select
        With Selection
            .Interior.ColorIndex = xlNone
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            With .Font
                .FontStyle = "Regular"
                .Size = "10"
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
            End With
        End With
    'Kopiera alla resurser från temp fil
    Selection.Copy
    End With
    
    'Klistra in i .FOR-fil
    sheetFOR.Activate
    With ActiveSheet
    Range("I30").Select
    .Paste
    
    'Separera Material från anm.2
    With sheetFOR
    .[F:F].TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="§", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    End With
    
    On Error GoTo Errmsg3
    
    End With
    
    Call Konvertera_DIGMA
    
    Exit Sub
    
Errmsg:
   MsgBox ("Misslyckades att hitta resurser i 'RWartikel', se till att den finns."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
    
Errmsg2:
   MsgBox ("Misslyckades att återställa format på celler i 'temp'."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
   
Errmsg3:
   MsgBox ("Misslyckades att radera temporära blad, dem kanske redan blivit raderade?"), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
End Sub

Sub Konvertera_DIGMA()
    On Error GoTo Errmsg
    
    'Bestäm variabler
    Dim selRange, selRangeMERGE As String
    Set sheetFOR = Sheets(".FOR")
    Set sheetBAK = Sheets("KAPNOTA")
    Set sheetMERGE = Sheets("KONVERTERA")
    Set sheetARTIKEL = Sheets("ARTIKELREG")
    Set sheetTEMP = Sheets("TEMP")
    
    'Lägg in beskrivningar
    With sheetMERGE
        .[A1].FormulaR1C1 = "ItemNo"
        .[B1].FormulaR1C1 = "Ordernummer"
        .[C1].FormulaR1C1 = "Art.nr."
        .[D1].FormulaR1C1 = "Antal"
        .[E1].FormulaR1C1 = "Benämning"
        .[F1].FormulaR1C1 = "Material"
        .[G1].FormulaR1C1 = "Längd"
        .[H1].FormulaR1C1 = "Bredd"
        .[I1].FormulaR1C1 = "Tj."
        .[J1].FormulaR1C1 = "Anm.2"
        .[K1].FormulaR1C1 = "klippm.höjd"
        .[L1].FormulaR1C1 = "Klippm.br1."
        .[M1].FormulaR1C1 = "Klippm.br2."
        .[N1].FormulaR1C1 = "resurs"
        .[O1].FormulaR1C1 = "Konstruktör"
    End With
    
    'Definiera range för .FOR
    sheetFOR.Activate
    With ActiveSheet
    Range(.Range("A31"), ("A" & .UsedRange.Rows.Count)).Select
    selRange = Selection.Row + Selection.Rows.Count - 1
    End With
    
    'Lägg in varje kolumn i sheetFOR till sheetMERGE
    sheetFOR.Range("J31:J" & selRange).Copy Destination:=sheetMERGE.Range("C2")
    sheetFOR.Range("A31:A" & selRange).Copy Destination:=sheetMERGE.Range("D2")
    sheetFOR.Range("H31:H" & selRange).Copy Destination:=sheetMERGE.Range("E2")
    sheetFOR.Range("K31:K" & selRange).Copy Destination:=sheetMERGE.Range("F2")
    sheetFOR.Range("B31:B" & selRange).Copy Destination:=sheetMERGE.Range("G2")
    sheetFOR.Range("C31:C" & selRange).Copy Destination:=sheetMERGE.Range("H2")
    sheetFOR.Range("D31:D" & selRange).Copy Destination:=sheetMERGE.Range("I2")
    sheetFOR.Range("L31:L" & selRange).Copy Destination:=sheetMERGE.Range("J2")
    sheetFOR.Range("I31:I" & selRange).Copy Destination:=sheetMERGE.Range("N2")
    
    'Separera mått från "x" till olika celler
    With sheetFOR
    .[B18].TextToColumns Destination:=Range("B18"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="x", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    End With
    
    'Definiera range för KONVERTERA
    sheetMERGE.Activate
    With ActiveSheet
    Range(.Range("A1"), ("A" & .UsedRange.Rows.Count)).Select
    selRangeMERGE = Selection.Row + Selection.Rows.Count - 1
    End With
    
    'Lägg in fler rader, benämningar & resurs
    With sheetMERGE
        
        '-------------BENÄMNINGAR--------------'
        .Range("E" & selRangeMERGE + 2).FormulaR1C1 = "LAMINAT INVÄNDIGT"
        .Range("E" & selRangeMERGE + 3).FormulaR1C1 = " > MÅTT"
        .Range("E" & selRangeMERGE + 4).FormulaR1C1 = "LAMINAT UTVÄNDIGT"
        .Range("E" & selRangeMERGE + 5).FormulaR1C1 = "LAMINAT TAK"
        .Range("E" & selRangeMERGE + 6).FormulaR1C1 = "GOLVBELÄGGNING"
        .Range("E" & selRangeMERGE + 7).FormulaR1C1 = "MÅTT"
        .Range("E" & selRangeMERGE + 8).FormulaR1C1 = "MÅTT"
        .Range("E" & selRangeMERGE + 9).FormulaR1C1 = "MÅTT"
        
        '----------------RESURS----------------'
        .Range("N" & selRangeMERGE + 1).FormulaR1C1 = "ömått"
        .Range("N" & selRangeMERGE + 2).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 3).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 4).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 5).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 6).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 7).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 8).FormulaR1C1 = "msk"
        .Range("N" & selRangeMERGE + 9).FormulaR1C1 = "msk"
        
        '----------------STJÄRNA---------------'
        .Range("A" & selRangeMERGE + 1).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 2).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 3).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 4).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 5).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 6).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 7).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 8).FormulaR1C1 = "*"
        .Range("A" & selRangeMERGE + 9).FormulaR1C1 = "*"
        
        '---------------RADNUMMER--------------'
        sheetMERGE.Range("A1:A" & selRangeMERGE).Value = "0"
    End With
    
    'Lägg in information
    sheetFOR.Range("B19").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 2)
    sheetFOR.Range("B20").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 3)
    sheetFOR.Range("B5").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 4)
    sheetFOR.Range("B6").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 5)
    sheetMERGE.Range("J" & selRangeMERGE + 6).Value = sheetFOR.[B25].Value & " " & sheetFOR.[B26].Value
    sheetFOR.Range("B21").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 7)
    sheetFOR.Range("B22").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 8)
    sheetFOR.Range("B23").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 9)
    
    'Lägg in beredare & order nr.
    sheetFOR.Range("B16").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 1)
    sheetMERGE.Range("C" & selRangeMERGE + 1).FormulaR1C1 = Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    
    'Lägg in mått
    sheetFOR.Range("B18").Copy Destination:=sheetMERGE.Range("G" & selRangeMERGE + 1)
    sheetFOR.Range("C18").Copy Destination:=sheetMERGE.Range("H" & selRangeMERGE + 1)
    sheetFOR.Range("D18").Copy Destination:=sheetMERGE.Range("I" & selRangeMERGE + 1)
    
    'Radera ogiltiga mått etc.
    sheetMERGE.Activate
    On Error Resume Next
    Range("J" & selRangeMERGE + 1, "J" & selRangeMERGE + 9).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Application.DisplayAlerts = False
    Sheets("TEMP").Delete
    Sheets(".FOR").Delete
    Sheets("KAPNOTA").Delete

    Application.Calculation = xlCalculationAutomatic
    
'Spara som ny order
    
    sheetMERGE.Activate
    Dim saveDir As String
    Dim saveFile As String
    saveDir = "L:\Order\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    saveFile = "L:\Order\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value & "\" & Sheets("PROGRAMÖVERSIKT").TextBox1.Value & ".xls"
    If FileFolderExists(saveDir) Then
    Else
        MkDir saveDir
    End If
    ActiveSheet.Copy
    With ActiveSheet.UsedRange
        .Copy
        .PasteSpecial xlValues
        .PasteSpecial xlFormats
    End With
    ActiveWorkbook.SaveAs Filename:=saveFile, _
    FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
    ReadOnlyRecommended:=False, CreateBackup:=False
    sheetMERGE.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
Errmsg:
   MsgBox ("Misslyckades att konvertera till DIGMA format."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
    
End Sub
