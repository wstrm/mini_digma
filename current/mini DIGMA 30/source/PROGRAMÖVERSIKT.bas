Attribute VB_Name = "PROGRAMÖVERSIKT"
Dim OrderNrText As String
Dim KundText As String
Dim HissTyp As String
Dim Anmärkning As String
Dim KonfOrder, KonfFor, KonfKap, FolderAndFile As String
Dim Mått1, Mått2, Mått3, Mått4, Mått5, Mått6 As String
Dim OldMeny As Worksheet
Dim Programöversikt As Worksheet

Sub Kör_minidigma28()

'Variabler
    Set OldMeny = Sheets("Meny")
    Set Programöversikt = Sheets("PROGRAMÖVERSIKT")
    OrderNrText = Sheets("PROGRAMÖVERSIKT").TextBox1.Value
    KundText = Sheets("PROGRAMÖVERSIKT").TextBox2.Text
    HissTyp = Sheets("PROGRAMÖVERSIKT").TextBox3.Text
    Anmärkning = Sheets("PROGRAMÖVERSIKT").TextBox4.Text
    Mått1 = Sheets("PROGRAMÖVERSIKT").TextBox5.Value
    Mått2 = Sheets("PROGRAMÖVERSIKT").TextBox6.Value
    Mått3 = Sheets("PROGRAMÖVERSIKT").TextBox7.Value
    Mått4 = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
    Mått5 = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
    Mått6 = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
    KonfOrder = Sheets("PROGRAMÖVERSIKT").TextBox11.Text & "\"
    KonfFor = Sheets("PROGRAMÖVERSIKT").TextBox12.Text & "\"
    KonfKap = Sheets("PROGRAMÖVERSIKT").TextBox13.Text & "\"
    FolderAndFile = KonfOrder & OrderNrText & "\" & OrderNrText & ".xls"

'Leta efter ordernummer för att se att mappen finns, annars leta efter den bland träkorgar
If FileFolderExists(FolderAndFile) Then
        Call DIGMA
    Else
        Call Träkorg
End If

End Sub

Sub DIGMA()

'Lägg in all information
    Sheets("Meny").[B3] = OrderNrText
    Sheets("Meny").[B12] = KundText
    Sheets("Meny").[B13] = HissTyp
    Sheets("Meny").[B16] = Anmärkning
    Sheets("Meny").CheckBoxes("Kryssruta 4").Value = Sheets("PROGRAMÖVERSIKT").CheckBox1.Value
    Sheets("Meny").CheckBoxes("Print").Value = Sheets("PROGRAMÖVERSIKT").CheckBox2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 5").Value = Sheets("PROGRAMÖVERSIKT").OptionButton1.Value
    Sheets("Meny").OptionButtons("Alternativknapp 6").Value = Sheets("PROGRAMÖVERSIKT").OptionButton2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Sheets("PROGRAMÖVERSIKT").OptionButton3.Value
    
    If Sheets("PROGRAMÖVERSIKT").OptionButton1.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
    ElseIf Sheets("PROGRAMÖVERSIKT").OptionButton2.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
        Sheets("Meny").[F21] = Sheets("PROGRAMÖVERSIKT").TextBox5.Value
        Sheets("Meny").[G21] = Sheets("PROGRAMÖVERSIKT").TextBox6.Value
        Sheets("Meny").[H21] = Sheets("PROGRAMÖVERSIKT").TextBox7.Value
    End If
        Call ok
End Sub

Sub Träkorg()

If FileFolderExists(KonfFor & OrderNrText & ".FOR") Then
    Sheets("Meny").[B3] = OrderNrText
    Sheets("Meny").[B16] = Anmärkning
    Sheets("Meny").CheckBoxes("Kryssruta 4").Value = Sheets("PROGRAMÖVERSIKT").CheckBox1.Value
    Sheets("Meny").CheckBoxes("Print").Value = Sheets("PROGRAMÖVERSIKT").CheckBox2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 5").Value = Sheets("PROGRAMÖVERSIKT").OptionButton1.Value
    Sheets("Meny").OptionButtons("Alternativknapp 6").Value = Sheets("PROGRAMÖVERSIKT").OptionButton2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Sheets("PROGRAMÖVERSIKT").OptionButton3.Value
    
    If Sheets("PROGRAMÖVERSIKT").OptionButton1.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
    ElseIf Sheets("PROGRAMÖVERSIKT").OptionButton2.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAMÖVERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAMÖVERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAMÖVERSIKT").TextBox8.Value
        Sheets("Meny").[F21] = Sheets("PROGRAMÖVERSIKT").TextBox5.Value
        Sheets("Meny").[G21] = Sheets("PROGRAMÖVERSIKT").TextBox6.Value
        Sheets("Meny").[H21] = Sheets("PROGRAMÖVERSIKT").TextBox7.Value
    End If
    Call ÖppnaFOR
    Call ok
    Exit Sub
    
Errmsg:
    MsgBox ("Misslyckades att köra ÖppnaFOR"), vbOKOnly, ".FOR-fil"
    Else
        MsgBox ("Ordernumret finns inte.")
End If

End Sub


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'-------------KONFIGURATION------------'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
Sub ÖppnaKatalogDialog()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAMÖVERSIKT").TextBox11.Text = .SelectedItems(1)
    End With
End Sub
Sub ÖppnaKatalogDialog2()
        With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAMÖVERSIKT").TextBox12.Text = .SelectedItems(1)
    End With
End Sub
Sub ÖppnaKatalogDialog3()
            With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAMÖVERSIKT").TextBox13.Text = .SelectedItems(1)
    End With
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
