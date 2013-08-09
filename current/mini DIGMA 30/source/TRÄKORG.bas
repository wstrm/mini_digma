Attribute VB_Name = "Tr�korg"
Dim filFOR As String
Dim filBAK As String
Dim filBakNoDir As String
Dim sheetFOR As Worksheet
Dim sheetBAK As Worksheet
Dim sheetMERGE As Worksheet
Dim sheetARTIKEL As Worksheet
Dim sheetTEMP As Worksheet
Dim sheetRESURSER As Worksheet

Sub �ppnaFOR()
Attribute �ppnaFOR.VB_Description = "�ppna .FOR-fil f�r att sedan visas i Excel"
Attribute �ppnaFOR.VB_ProcData.VB_Invoke_Func = " \n14"

' ---�ppnaFOR---
' William Wennerstr�m - 2013/06/19
' "�ppnar .FOR-fil f�r att sedan visas i Excel"

    On Error GoTo Errmsg

 'G�m f�nster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveSheet.AutoFilterMode = False
    
'Radera gamla blad, ifall de finns kvar av n�gon anledning
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("TEMP").Delete
    Sheets(".FOR").Delete
    Sheets("KAPNOTA").Delete
    Sheets("KONVERTERA").Delete
    Application.DisplayAlerts = True
    
'L�gg till nya blad
    Sheets.Add.Name = "TEMP"
    Sheets.Add.Name = "KAPNOTA"
    Sheets.Add.Name = ".FOR"
    Sheets.Add.Name = "KONVERTERA"
    
 'Best�m variabler
    Set sheetFOR = Sheets(".FOR")
    Set sheetBAK = Sheets("KAPNOTA")
    Set sheetMERGE = Sheets("KONVERTERA")
    Set sheetARTIKEL = Sheets("ARTIKELREG")
    Set sheetTEMP = Sheets("TEMP")
    filFOR = miniDIGMAForm.FORfile_Path.Text & "\" & miniDIGMAForm.OrderNummer_Text.Value & ".FOR"
    filBAK = miniDIGMAForm.Kapnot_Path.Text & "\" & miniDIGMAForm.OrderNummer_Text.Value
    
'Visa blad
    sheetARTIKEL.Visible = True
    
'Leta efter .FOR-fil i L:\AM\PRO f�r att sedan �ppna den och f�rbereder den f�r Excel.
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
    
    'On Error GoTo Errmsg2
    
    'K�r MergeBAK_FOR
    Call MergeBAK_FOR
    
Exit Sub

Errmsg:
   MsgBox ("Misslyckades att �ppna '" & filFOR & "', se till att den finns eller inte �r �ppen."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
Errmsg2:
    MsgBox ("Misslyckades att k�ra �ppnaBAK"), vbOKOnly, ".BAK-fil"
    Application.ScreenUpdating = True
End Sub
Sub MergeBAK_FOR()
Attribute MergeBAK_FOR.VB_Description = "Makrot inspelat 2013-06-19 av mattias"
Attribute MergeBAK_FOR.VB_ProcData.VB_Invoke_Func = " \n14"

' ---MergeBAK_FOR---
' William Wennerstr�m - 2013/06/20 | Uppdaterad 2013-06-25 & 2013-06-26
' "Sammans�tter BAK med FOR-filen."

    'Best�m variabler
    filBakNoDir = miniDIGMAForm.OrderNummer_Text.Value
    
    On Error GoTo Errmsg
    
    sheetBAK.Activate
With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filBAK, _
        Destination:=Range("$A$1"))
        .Name = filBakNoDir & ".BAK"
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
    
    'L�gg in 27st rader �verst.
    sheetFOR.[1:27].Insert Shift:=xlDown
    
    'L�s in korgtyp och �ndra bild i bladet TR�KORG
    Call GetIngType(sheetBAK.Range("A11"), sheetBAK.Range("A12"), sheetBAK.Range("A13"))
        
    'Kopiera fr�n BAK till FOR
    sheetBAK.Range("A2,A3,A4,A6,A9,A10:A13").Copy Destination:=sheetFOR.Range("B1")
    sheetBAK.Range("B2,B3,B4,B6,B9,B10").Copy Destination:=sheetFOR.Range("B12")
   
    'St�ng BAK och �ppna igen med ny FieldInfo
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(15 _
        , 1), Array(40, 1))
        
    'Kopiera fr�n BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B5,B7,B8").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B18").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    
    'St�ng BAK och �ppna igen med ny FieldInfo
    ActiveWindow.Close False
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(36 _
        , 1), Array(58, 1))
        
    'Kopiera in fr�n BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B13,B12,B11").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B21").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    
    'St�ng BAK och �ppna igen med ny FieldInfo
    ActiveWindow.Close False
    Workbooks.OpenText Filename:=filBAK, Origin:=xlMSDOS, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(15 _
        , 1), Array(59, 1))
        
    'Kopiera in fr�n BAK till FOR
    Windows(filBakNoDir).Activate
    Range("B14,B15,B16").Select
    Selection.Copy
    sheetFOR.Activate
    Range("B24").Select
    ActiveSheet.Paste
    Windows(filBakNoDir).Activate
    ActiveWindow.Close False
    
    'L�gg in information till "RESURSER"
    Sheets("RESURSER").[A5] = Sheets(".FOR").[B13]
    Sheets("RESURSER").[A6] = Sheets(".FOR").[B15]
    
    'V�lj B30:E?? f�r att sedan konvertera nummer & text
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
    
    'Beh�ll endast siffror och komma
    On Error Resume Next
    Dim RE As Object
    Dim rng As Range
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Global = True
        'Till�t "," och alla nummer mellan 0-9
       .Pattern = "[^,0-9]"
       For Each rng In Selection
            rng.Value = .Replace(rng.Value, "")
        Next rng
    End With
    
    'L�s som nummer ist�llet f�r text
    'Blir ett error h�r, anv�nder d�rav Resume Next
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
    Call RWartikel_S�k
    Exit Sub
    

Errmsg:
   MsgBox ("Misslyckades att �ppna '" & filBAK & "', se till att den finns eller inte �r �ppen."), vbOKOnly, ".BAK-fil"
    Application.ScreenUpdating = True
    Exit Sub
Errmsg2:
   MsgBox ("Misslyckades att k�ra sammans�ttning av '" & filBAK & "' och '" & filFOR & "'."), vbOKOnly, "Sammans�ttning"
    Application.ScreenUpdating = True
End Sub

Sub RWartikel_S�k()
    
' ---RWartikel_S�k---
' William Wennerstr�m - 2013/06/25
' "�ppnar RWartikel.xls, h�mtar sedan resurs till .FOR-fil."
    
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
    
    'V�lj alla artiklar
    
    sheetFOR.Activate
    With ActiveSheet
    Range(.Range("E30"), ("E" & .UsedRange.Rows.Count)).Select
    Set selRng = Selection
    End With
    
    'Best�m artikel att leta efter
    For Each xCell In selRng
    
    'Best�m vilken artikel som ska hittas
    strToFind = xCell.Value
    xCellRow = xCell.Row
    
    'S�k i RWartikel och kopiera resurs f�r varje artikel
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
    
    'V�lj och ta bort all format p� celler
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
    'Kopiera alla resurser fr�n temp fil
    Selection.Copy
    End With
    
    'Klistra in i .FOR-fil
    sheetFOR.Activate
    With ActiveSheet
    Range("I30").Select
    .Paste
    
    'Separera Material fr�n anm.2
    With sheetFOR
    .[F:F].TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="�", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    End With
    
    On Error GoTo Errmsg3
    
    End With
    
    Call Konvertera_DIGMA
    
    Exit Sub
    
Errmsg:
   miniDIGMAForm.Status_Label.Caption = "Misslyckades att �terst�lla format p� celler i 'temp'."
   MsgBox ("Misslyckades att hitta resurser i 'RWartikel', se till att den finns."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
    
Errmsg2:
   miniDIGMAForm.Status_Label.Caption = "Misslyckades att �terst�lla format p� celler i 'temp'."
   MsgBox ("Misslyckades att �terst�lla format p� celler i 'temp'."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   Exit Sub
   
Errmsg3:
   miniDIGMAForm.Status_Label.Caption = "Misslyckades att radera tempor�ra blad, dem kanske redan blivit raderade?"
   MsgBox ("Misslyckades att radera tempor�ra blad, dem kanske redan blivit raderade?"), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
End Sub

Sub Konvertera_DIGMA()
    On Error GoTo Errmsg
    
    'Best�m variabler
    Dim selRange, selRangeMERGE As String
    Dim sheetRESURSER As Worksheet
    Set sheetRESURSER = Sheets("RESURSER")
    Set sheetFOR = Sheets(".FOR")
    Set sheetBAK = Sheets("KAPNOTA")
    Set sheetMERGE = Sheets("KONVERTERA")
    Set sheetARTIKEL = Sheets("ARTIKELREG")
    Set sheetTEMP = Sheets("TEMP")
    
    'L�gg in beskrivningar
    With sheetMERGE
        .[A1].FormulaR1C1 = "ItemNo"
        .[B1].FormulaR1C1 = "Ordernummer"
        .[C1].FormulaR1C1 = "Art.nr."
        .[D1].FormulaR1C1 = "Antal"
        .[E1].FormulaR1C1 = "Ben�mning"
        .[F1].FormulaR1C1 = "Material"
        .[G1].FormulaR1C1 = "L�ngd"
        .[H1].FormulaR1C1 = "Bredd"
        .[I1].FormulaR1C1 = "Tj."
        .[J1].FormulaR1C1 = "Anm.2"
        .[K1].FormulaR1C1 = "klippm.h�jd"
        .[L1].FormulaR1C1 = "Klippm.br1."
        .[M1].FormulaR1C1 = "Klippm.br2."
        .[N1].FormulaR1C1 = "resurs"
        .[O1].FormulaR1C1 = "Konstrukt�r"
    End With
    
    'Definiera range f�r .FOR
    sheetFOR.Activate
    With ActiveSheet
    Range(.Range("A31"), ("A" & .UsedRange.Rows.Count)).Select
    selRange = Selection.Row + Selection.Rows.Count - 1
    End With
    
    'L�gg in varje kolumn i sheetFOR till sheetMERGE
    sheetFOR.Range("J31:J" & selRange).Copy Destination:=sheetMERGE.Range("C2")
    sheetFOR.Range("A31:A" & selRange).Copy Destination:=sheetMERGE.Range("D2")
    sheetFOR.Range("H31:H" & selRange).Copy Destination:=sheetMERGE.Range("E2")
    sheetFOR.Range("K31:K" & selRange).Copy Destination:=sheetMERGE.Range("F2")
    sheetFOR.Range("B31:B" & selRange).Copy Destination:=sheetMERGE.Range("G2")
    sheetFOR.Range("C31:C" & selRange).Copy Destination:=sheetMERGE.Range("H2")
    sheetFOR.Range("D31:D" & selRange).Copy Destination:=sheetMERGE.Range("I2")
    sheetFOR.Range("L31:L" & selRange).Copy Destination:=sheetMERGE.Range("J2")
    sheetFOR.Range("I31:I" & selRange).Copy Destination:=sheetMERGE.Range("N2")
    
    'Separera m�tt fr�n "x" till olika celler
    With sheetFOR
    .[B18].TextToColumns Destination:=Range("B18"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="x", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    End With
    
    'Kalla p� FindDimDraw, vilket hittar alla dimensioner och konverterar dem f�r att sedan kunna anv�ndas
    'n�r CalcDimDraw r�knar ut de olika m�tten f�r tr�ramsritningen
    Call FindDimDraw
    
    'Definiera range f�r KONVERTERA
    sheetMERGE.Activate
    With ActiveSheet
    Range(.Range("A1"), ("A" & .UsedRange.Rows.Count)).Select
    selRangeMERGE = Selection.Row + Selection.Rows.Count - 1
    End With
    
    'L�gg in fler rader, ben�mningar & resurs
    With sheetMERGE
        
        '-------------BEN�MNINGAR--------------'
        '.Range("E" & selRangeMERGE + 2).FormulaR1C1 = "LAMINAT INV�NDIGT"
        '.Range("E" & selRangeMERGE + 3).FormulaR1C1 = " > M�TT"
        '.Range("E" & selRangeMERGE + 4).FormulaR1C1 = "LAMINAT UTV�NDIGT"
        '.Range("E" & selRangeMERGE + 5).FormulaR1C1 = "LAMINAT TAK"
        '.Range("E" & selRangeMERGE + 6).FormulaR1C1 = "GOLVBEL�GGNING"
        '.Range("E" & selRangeMERGE + 7).FormulaR1C1 = "M�TT"
        '.Range("E" & selRangeMERGE + 8).FormulaR1C1 = "M�TT"
        '.Range("E" & selRangeMERGE + 9).FormulaR1C1 = "M�TT"
        
        '----------------RESURS----------------'
        .Range("N" & selRangeMERGE + 1).FormulaR1C1 = "�m�tt"
        
        '----------------STJ�RNA---------------'
        .Range("A" & selRangeMERGE + 1).FormulaR1C1 = "*"
        
        '---------------RADNUMMER--------------'
        sheetMERGE.Range("A1:A" & selRangeMERGE).Value = "0"
    End With
    
    'L�gg in information
    sheetFOR.Range("B19").Copy Destination:=sheetRESURSER.Range("A7") 'Lam. inv
    sheetFOR.Range("B20").Copy Destination:=sheetRESURSER.Range("A8") '> M�tt.
    sheetFOR.Range("B5").Copy Destination:=sheetRESURSER.Range("A9") 'Lam. inv.
    sheetFOR.Range("B6").Copy Destination:=sheetRESURSER.Range("A10") 'Lam. tak
    sheetRESURSER.Range("A11").Value = sheetFOR.[B25].Value & " " & sheetFOR.[B26].Value 'Golvbel.
    'sheetFOR.Range("B21").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 7) 'M�tt 1
    'sheetFOR.Range("B22").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 8) 'M�tt 2
    'sheetFOR.Range("B23").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 9) 'M�tt 3
    
    'L�gg in beredare & order nr.
    sheetFOR.Range("B16").Copy Destination:=sheetMERGE.Range("J" & selRangeMERGE + 1)
    sheetMERGE.Range("C" & selRangeMERGE + 1).FormulaR1C1 = filBakNoDir
    
    'L�gg in m�tt
    sheetFOR.Range("B18").Copy Destination:=sheetMERGE.Range("G" & selRangeMERGE + 1)
    sheetFOR.Range("C18").Copy Destination:=sheetMERGE.Range("H" & selRangeMERGE + 1)
    sheetFOR.Range("D18").Copy Destination:=sheetMERGE.Range("I" & selRangeMERGE + 1)
    
    'Radera ogiltiga m�tt etc.
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
    saveDir = miniDIGMAForm.OrderPath_Text.Text & "\" & filBakNoDir
    saveFile = miniDIGMAForm.OrderPath_Text.Text & "\" & filBakNoDir & "\" & filBakNoDir & ".xls"
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
    Workbooks(filBakNoDir & ".xls").Close SaveChanges:=False
    sheetMERGE.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
Errmsg:
   MsgBox ("Misslyckades att konvertera till DIGMA format."), vbOKOnly, ".FOR-fil"
   Application.ScreenUpdating = True
   miniDIGMAForm.Status_Label.Caption = "Misslyckades att konvertera till DIGMA format."
   Exit Sub
    
End Sub

Public Sub FindDimDraw()

    Set sheetFOR = Sheets(".FOR")
    
    On Error Resume Next
    Application.DisplayAlerts = False
    With sheetFOR
    .[B21].TextToColumns Destination:=Range("E21"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1)), _
        TrailingMinusNumbers:=True
    .[B22].TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1)), _
        TrailingMinusNumbers:=True
    .[B23].TextToColumns Destination:=Range("E23"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1)), _
        TrailingMinusNumbers:=True
    '------------------------------------------------------------------------------'
    .[E21].TextToColumns Destination:=Range("B21"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="x", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    .[E22].TextToColumns Destination:=Range("B22"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="x", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    .[E23].TextToColumns Destination:=Range("B23"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="x", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    End With
    Application.DisplayAlerts = True
    
    On Error GoTo 0
    
    Call CalcDimDraw(sheetFOR.[B21].Value, sheetFOR.[C21].Value, sheetFOR.[D21].Value, sheetFOR.[F21].Value, _
        sheetFOR.[B22].Value, sheetFOR.[C22].Value, sheetFOR.[D22].Value, sheetFOR.[F22].Value, _
        sheetFOR.[B23].Value, sheetFOR.[C23].Value, sheetFOR.[D23].Value, sheetFOR.[F23].Value, _
        sheetFOR.[B18].Value, sheetFOR.[C18].Value, sheetFOR.[D18].Value)

End Sub
                            'Ing. Kod 1 - 3 m�tt: (orginal format: BxCxD,F)                        | Inv. m�tt: (orginal format: BxCxD)'
                            'B21   C21   D21   F21   B22   C22   D22   F22   B23   C23   D23   F23 | B18   C18   D18'
Public Function CalcDimDraw(B11A, B11B, B11C, B11D, B12A, B12B, B12C, B12D, B13A, B13B, B13C, B13D, B18A, B18B, B18C As String) As String

    'MsgBox "Resultat: " & B11A & B11B & B11C & B11D & B12A & B12B & B12C & B12D & B13A & B13B & B13C & B13D & B18A & B18B & B18C

    Set sheetRESURSER = Sheets("RESURSER")

    'Ing.Kod 1:'
    'NR1:'
    sheetRESURSER.[A19].Value = B11D

    'NR2:'
    sheetRESURSER.[A20].Value = B11A

    'NR3:'
    sheetRESURSER.[A21].Value = B18A - B11D - B11A

    'NR4:'
    sheetRESURSER.[A22].Value = B11B

    'Ing.Kod 2:'
    'NR1:
    sheetRESURSER.[A23].Value = B12D

    'NR2:'
    sheetRESURSER.[A24].Value = B12A

    'NR3:'
    sheetRESURSER.[A25].Value = B18A - B12D - B12A

    'NR4:'
    sheetRESURSER.[A26].Value = B12B

    'Ing.Kod 3:'
    'NR1:
    sheetRESURSER.[A27].Value = B13D

    'NR2:'
    sheetRESURSER.[A28].Value = B13A

    'NR3:'
    sheetRESURSER.[A29].Value = B18B - B13D - B13A

    'NR4:'
    sheetRESURSER.[A30].Value = B13B

End Function

'FUNKTIONER:
Public Function GetIngType(intIngKOD1, intIngKOD2, intIngKOD3 As String) As Boolean
    If Not intIngKOD1 = "" And Not intIngKOD2 = "" Then
        If Right(intIngKOD3, 1) = "H" Then
            'MsgBox "FRAM, UPP & H�GER"
            With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = True
                .Shapes("top-left-bottom").Visible = False
                .Shapes("left-bottom").Visible = False
                .Shapes("right-bottom").Visible = False
                .Shapes("top-bottom").Visible = False
                .Shapes("bottom").Visible = False
                .Shapes("HIDER-LEFT").Visible = True
                .Shapes("HIDER-RIGHT").Visible = False
                .Shapes("HIDER-TOP").Visible = False
            End With
        ElseIf Right(intIngKOD3, 1) = "V" Then
            'MsgBox "FRAM, UPP & V�NSTER"
            With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = False
                .Shapes("top-left-bottom").Visible = True
                .Shapes("left-bottom").Visible = False
                .Shapes("right-bottom").Visible = False
                .Shapes("top-bottom").Visible = False
                .Shapes("bottom").Visible = False
                .Shapes("HIDER-LEFT").Visible = False
                .Shapes("HIDER-RIGHT").Visible = True
                .Shapes("HIDER-TOP").Visible = False
            End With
        Else
            'MsgBox "FRAM & UPP"
            With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = False
                .Shapes("top-left-bottom").Visible = False
                .Shapes("left-bottom").Visible = False
                .Shapes("right-bottom").Visible = False
                .Shapes("top-bottom").Visible = True
                .Shapes("bottom").Visible = False
                .Shapes("HIDER-LEFT").Visible = True
                .Shapes("HIDER-RIGHT").Visible = True
                .Shapes("HIDER-TOP").Visible = False
            End With
        End If
    ElseIf Right(intIngKOD3, 1) = "H" Then
        'MsgBox "FRAM & H�GER"
        With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = False
                .Shapes("top-left-bottom").Visible = False
                .Shapes("left-bottom").Visible = False
                .Shapes("right-bottom").Visible = True
                .Shapes("top-bottom").Visible = False
                .Shapes("bottom").Visible = False
                .Shapes("HIDER-LEFT").Visible = True
                .Shapes("HIDER-RIGHT").Visible = False
                .Shapes("HIDER-TOP").Visible = True
        End With
    ElseIf Right(intIngKOD3, 1) = "V" Then
        'MsgBox "FRAM & V�NSTER"
        With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = False
                .Shapes("top-left-bottom").Visible = False
                .Shapes("left-bottom").Visible = True
                .Shapes("right-bottom").Visible = False
                .Shapes("top-bottom").Visible = False
                .Shapes("bottom").Visible = False
                .Shapes("HIDER-LEFT").Visible = False
                .Shapes("HIDER-RIGHT").Visible = True
                .Shapes("HIDER-TOP").Visible = True
        End With
    Else
        'MsgBox "INGEN KOD2 & 3 AKA. FRAM"
        With Sheets("TR�KORG")
                .Shapes("top-right-bottom").Visible = False
                .Shapes("top-left-bottom").Visible = False
                .Shapes("left-bottom").Visible = False
                .Shapes("right-bottom").Visible = False
                .Shapes("top-bottom").Visible = False
                .Shapes("bottom").Visible = True
                .Shapes("HIDER-LEFT").Visible = True
                .Shapes("HIDER-RIGHT").Visible = True
                .Shapes("HIDER-TOP").Visible = True
        End With
    End If
End Function
