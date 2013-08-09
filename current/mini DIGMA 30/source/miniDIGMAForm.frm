VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} miniDIGMAForm 
   Caption         =   "Kontrollpanel - mini DIGMA 30"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10545
   OleObjectBlob   =   "miniDIGMAForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "miniDIGMAForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrderNrText As String
Dim KundText As String
Dim HissTyp As String
Dim Anmärkning As String
Dim KonfOrder, KonfFor, KonfKap As String
Dim Mått1, Mått2, Mått3, Mått4, Mått5, Mått6 As String
Dim OldMeny As Worksheet

Private Sub OrderNummer_Text_Change()
    Sheets("RESURSER").[A18] = OrderNummer_Text.Value
End Sub

Sub Run_miniDIGMA_Click()
    'Variabler
    Status_Label.Caption = "Arbetar..."
    Set OldMeny = ThisWorkbook.Sheets("Meny")
    OrderNrText = OrderNummer_Text.Value
    KundText = Kund_Text.Text
    HissTyp = Typ_Text.Text
    Anmärkning = Note_Text.Text
    Mått1 = Dimension1_Text.Value
    Mått2 = Dimension2_Text.Value
    Mått3 = Dimension3_Text.Value
    Mått4 = Dimension4_Text.Value
    Mått5 = Dimension5_Text.Value
    Mått6 = Dimension6_Text.Value
    KonfOrder = OrderPath_Text.Text & "\"
    KonfFor = FORfile_Path.Text & "\"
    KonfKap = Kapnot_Path.Text & "\"
    
    'MsgBox ("DEBUG: " & vbCrLf & "-----------------" & vbCrLf & OrderNrText & vbCrLf & KundText & vbCrLf & HissTyp & vbCrLf & Anmärkning & vbCrLf & Mått1 & " - " & Mått2 & " - " & Mått3 & vbCrLf & Mått4 & " - " & Mått5 & " - " & Mått6 & vbCrLf & KonfOrder & vbCrLf & KonfFor & vbCrLf & KonfKap & vbCrLf & FolderAndFile)
    
    'Se till att Ordernummer inte är tomt
    If OrderNrText = "" Then Status_Label.Caption = "Var vänlig skriv in ett order nummer": MsgBox "Var vänlig skriv in ett order nummer.", vbInformation, "mini DIGMA 30": Exit Sub
        
    'Leta efter ordernummer för att se att mappen finns, annars leta efter den bland träkorgar
    If FileFolderExists(KonfKap & "\" & OrderNrText) Then
        Call TRÄKORG
    ElseIf FileFolderExists(KonfOrder & "\" & OrderNrText & "\" & OrderNrText & ".xls") Then
        Call DIGMA
    Else
        Status_Label.Caption = "Order nummer '" & OrderNrText & "' finns inte"
        MsgBox "Order nummer '" & OrderNrText & "' finns inte.", vbCritical, "mini DIGMA 30"
    End If
End Sub

Sub DIGMA()

'Lägg in all information
    'Sheets("Meny").[B3] = OrderNrText
    Sheets("RESURSER").[A5] = KundText
    Sheets("RESURSER").[A6] = HissTyp
    Sheets("RESURSER").[A18] = OrderNrText
    Sheets("TRÄKORG").Visible = False
    'Sheets("Meny").[B16] = Anmärkning
    'Sheets("Meny").CheckBoxes("Kryssruta 4").Value = OpenLoad_Check.Value
    'Sheets("Meny").CheckBoxes("Print").Value = PrintList_Check.Value
    'Sheets("Meny").OptionButtons("Alternativknapp 5").Value = OneDoor_Radio.Value
    'Sheets("Meny").OptionButtons("Alternativknapp 6").Value = TwoDoor_Radio.Value
    'Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Special_Radio.Value
    
    If OneDoor_Radio.Value = True Then
    With Sheets("RESURSER")
        .[A15] = Dimension4_Text.Value
        .[A16] = Dimension5_Text.Value
        .[A17] = Dimension6_Text.Value
        .[A12].FormulaR1C1 = ""
        .[A13].FormulaR1C1 = ""
        .[A14].FormulaR1C1 = ""
    End With
    With Sheets("UTSKRIFT")
        .Shapes("bild_special").Visible = False
        .Shapes("bild_ingenritning").Visible = False
        .Shapes("bild_2_ing_sid1").Visible = False
        .Shapes("bild_1_ing_sid1").Visible = True
        .Shapes("bild_special2").Visible = False
        .Shapes("bild_ingenritning2").Visible = False
        .Shapes("bild_2_ing_sid2").Visible = False
        .Shapes("bild_1_ing_sid2").Visible = True
        .Shapes("bild_plåtkorg").Visible = True
        .Shapes("bild_plåtkorg2").Visible = True
    End With
        Call OneDoor_Radio_Click
    ElseIf TwoDoor_Radio.Value = True Then
    With Sheets("RESURSER")
        .[A15] = Dimension4_Text.Value
        .[A16] = Dimension5_Text.Value
        .[A17] = Dimension6_Text.Value
        .[A12] = Dimension1_Text.Value
        .[A13] = Dimension2_Text.Value
        .[A14] = Dimension3_Text.Value
    End With
    With Sheets("UTSKRIFT")
        .Shapes("bild_special").Visible = False
        .Shapes("bild_ingenritning").Visible = False
        .Shapes("bild_2_ing_sid1").Visible = True
        .Shapes("bild_1_ing_sid1").Visible = False
        .Shapes("bild_special2").Visible = False
        .Shapes("bild_ingenritning2").Visible = False
        .Shapes("bild_2_ing_sid2").Visible = True
        .Shapes("bild_1_ing_sid2").Visible = False
        .Shapes("bild_plåtkorg").Visible = True
        .Shapes("bild_plåtkorg2").Visible = True
    End With
        Call TwoDoor_Radio_Click
    ElseIf Special_Radio.Value = True Then
    With Sheets("UTSKRIFT")
        .Shapes("bild_special").Visible = True
        .Shapes("bild_special2").Visible = True
        .Shapes("bild_ingenritning").Visible = False
        .Shapes("bild_2_ing_sid1").Visible = False
        .Shapes("bild_1_ing_sid1").Visible = False
        .Shapes("bild_ingenritning2").Visible = False
        .Shapes("bild_2_ing_sid2").Visible = False
        .Shapes("bild_1_ing_sid2").Visible = False
        .Shapes("bild_plåtkorg").Visible = True
        .Shapes("bild_plåtkorg2").Visible = True
    End With
        Call Special_Radio_Click
    End If
        Call runMiniDIGMA
End Sub

Sub TRÄKORG()

    'ThisWorkbook.Sheets("Meny").[B3] = OrderNrText
    Sheets("RESURSER").[A18] = OrderNrText
    'ThisWorkbook.Sheets("Meny").[B16] = Anmärkning
    'ThisWorkbook.Sheets("Meny").CheckBoxes("Kryssruta 4").Value = OpenLoad_Check.Value
    'ThisWorkbook.Sheets("Meny").CheckBoxes("Print").Value = PrintList_Check.Value
    'ThisWorkbook.Sheets("Meny").OptionButtons("Alternativknapp 5").Value = OneDoor_Radio.Value
    'ThisWorkbook.Sheets("Meny").OptionButtons("Alternativknapp 6").Value = TwoDoor_Radio.Value
    'ThisWorkbook.Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Special_Radio.Value
    
    Sheets("TRÄKORG").Visible = True
    
    With Sheets("UTSKRIFT")
        .Shapes("bild_special").Visible = False
        .Shapes("bild_special2").Visible = False
        .Shapes("bild_ingenritning").Visible = True
        .Shapes("bild_2_ing_sid1").Visible = False
        .Shapes("bild_1_ing_sid1").Visible = False
        .Shapes("bild_ingenritning2").Visible = True
        .Shapes("bild_2_ing_sid2").Visible = False
        .Shapes("bild_1_ing_sid2").Visible = False
        .Shapes("bild_plåtkorg").Visible = False
        .Shapes("bild_plåtkorg2").Visible = False
    End With
    
    'TA BORT FÖR ATT ERSÄTTAS MED DET NYA RITNINGS SYSTEMET!
    'If Sheets("PROGRAMÖVERSIKT").OptionButton1.Value = True Then
    '    Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
    '    Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
    '    Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
    'ElseIf Sheets("PROGRAMÖVERSIKT").OptionButton2.Value = True Then
    '    Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
    '    Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
    '    Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
    '    Sheets("Meny").[F21] = Sheets("PROGRAMÖVERSIKT").TextBox5.Value
    '    Sheets("Meny").[G21] = Sheets("PROGRAMÖVERSIKT").TextBox6.Value
    '    Sheets("Meny").[H21] = Sheets("PROGRAMÖVERSIKT").TextBox7.Value
    Call ÖppnaFOR
    Call runMiniDIGMA
    Exit Sub
    Status_Label.Caption = "Klar"
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'-------------KONFIGURATION------------'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
Sub ÖppnaKatalogDialog_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        OrderPath_Text.Text = .SelectedItems(1)
    End With
End Sub
Sub ÖppnaKatalogDialog2_Click()
        With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        FORfile_Path.Text = .SelectedItems(1)
    End With
End Sub
Sub ÖppnaKatalogDialog3_Click()
            With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Kapnot_Path.Text = .SelectedItems(1)
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
    If CloseMode = 0 Then
    
        Application.DisplayAlerts = True
        MsgBoxSave = MsgBox("Vill du spara ändringarna?", vbYesNoCancel + vbDefaultButton1 + vbQuestion, "mini DIGMA 30")

        If MsgBoxSave = vbYes Then
            ThisWorkbook.Save
            Application.Quit
        ElseIf MsgBoxSave = vbNo Then
            Application.Quit
        Else
            Cancel = True
        End If
        Application.DisplayAlerts = False
    End If
     
End Sub

Private Sub FORfile_Path_Change()
    With Sheets("RESURSER")
        .[A2].FormulaR1C1 = FORfile_Path.Text
    End With
End Sub

Private Sub Kapnot_Path_Change()
    With Sheets("RESURSER")
        .[A3].FormulaR1C1 = Kapnot_Path.Text
    End With
End Sub

Private Sub OrderPath_Text_Change()
    With Sheets("RESURSER")
        .[A1].FormulaR1C1 = OrderPath_Text.Text
    End With
End Sub

Private Sub Note_Text_Change()
    ThisWorkbook.Sheets("RESURSER").[A4].FormulaR1C1 = Note_Text.Text
End Sub

Private Sub Kund_Text_Change()
    ThisWorkbook.Sheets("RESURSER").[A5].FormulaR1C1 = Kund_Text.Text
End Sub

Private Sub Typ_Text_Change()
    ThisWorkbook.Sheets("RESURSER").[A6].FormulaR1C1 = Typ_Text.Text
End Sub

Private Sub UserForm_Initialize()
    OrderPath_Text.Text = ThisWorkbook.Sheets("RESURSER").[A1].FormulaR1C1
    FORfile_Path.Text = ThisWorkbook.Sheets("RESURSER").[A2].FormulaR1C1
    Kapnot_Path.Text = ThisWorkbook.Sheets("RESURSER").[A3].FormulaR1C1
    ThisWorkbook.Activate
    MultiPage1.Value = 0
End Sub

Private Sub Special_Radio_Click()
    OneDoor_IMG.Visible = False
    TwoDoor_IMG.Visible = False
    Dimension1_Text.Visible = False
    Dimension2_Text.Visible = False
    Dimension3_Text.Visible = False
    Dimension4_Text.Visible = False
    Dimension5_Text.Visible = False
    Dimension6_Text.Visible = False
    
End Sub

Private Sub TwoDoor_Radio_Click()
    OneDoor_IMG.Visible = False
    TwoDoor_IMG.Visible = True
    Dimension1_Text.Visible = True
    Dimension2_Text.Visible = True
    Dimension3_Text.Visible = True
    Dimension4_Text.Visible = True
    Dimension5_Text.Visible = True
    Dimension6_Text.Visible = True
    
End Sub

Private Sub OneDoor_Radio_Click()
    OneDoor_IMG.Visible = True
    TwoDoor_IMG.Visible = False
    Dimension1_Text.Visible = False
    Dimension2_Text.Visible = False
    Dimension3_Text.Visible = False
    Dimension4_Text.Visible = True
    Dimension5_Text.Visible = True
    Dimension6_Text.Visible = True
   
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'--------------FUNKTIONER--------------'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo Error
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
Error:
    On Error GoTo 0
End Function
