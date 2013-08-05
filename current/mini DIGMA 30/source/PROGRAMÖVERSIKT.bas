Attribute VB_Name = "PROGRAM�VERSIKT"
Dim OrderNrText As String
Dim KundText As String
Dim HissTyp As String
Dim Anm�rkning As String
Dim KonfOrder, KonfFor, KonfKap, FolderAndFile As String
Dim M�tt1, M�tt2, M�tt3, M�tt4, M�tt5, M�tt6 As String
Dim OldMeny As Worksheet
Dim Program�versikt As Worksheet

Sub K�r_minidigma28()

'Variabler
    Set OldMeny = Sheets("Meny")
    Set Program�versikt = Sheets("PROGRAM�VERSIKT")
    OrderNrText = Sheets("PROGRAM�VERSIKT").TextBox1.Value
    KundText = Sheets("PROGRAM�VERSIKT").TextBox2.Text
    HissTyp = Sheets("PROGRAM�VERSIKT").TextBox3.Text
    Anm�rkning = Sheets("PROGRAM�VERSIKT").TextBox4.Text
    M�tt1 = Sheets("PROGRAM�VERSIKT").TextBox5.Value
    M�tt2 = Sheets("PROGRAM�VERSIKT").TextBox6.Value
    M�tt3 = Sheets("PROGRAM�VERSIKT").TextBox7.Value
    M�tt4 = Sheets("PROGRAM�VERSIKT").TextBox8.Value
    M�tt5 = Sheets("PROGRAM�VERSIKT").TextBox9.Value
    M�tt6 = Sheets("PROGRAM�VERSIKT").TextBox10.Value
    KonfOrder = Sheets("PROGRAM�VERSIKT").TextBox11.Text & "\"
    KonfFor = Sheets("PROGRAM�VERSIKT").TextBox12.Text & "\"
    KonfKap = Sheets("PROGRAM�VERSIKT").TextBox13.Text & "\"
    FolderAndFile = KonfOrder & OrderNrText & "\" & OrderNrText & ".xls"

'Leta efter ordernummer f�r att se att mappen finns, annars leta efter den bland tr�korgar
If FileFolderExists(FolderAndFile) Then
        Call DIGMA
    Else
        Call Tr�korg
End If

End Sub

Sub DIGMA()

'L�gg in all information
    Sheets("Meny").[B3] = OrderNrText
    Sheets("Meny").[B12] = KundText
    Sheets("Meny").[B13] = HissTyp
    Sheets("Meny").[B16] = Anm�rkning
    Sheets("Meny").CheckBoxes("Kryssruta 4").Value = Sheets("PROGRAM�VERSIKT").CheckBox1.Value
    Sheets("Meny").CheckBoxes("Print").Value = Sheets("PROGRAM�VERSIKT").CheckBox2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 5").Value = Sheets("PROGRAM�VERSIKT").OptionButton1.Value
    Sheets("Meny").OptionButtons("Alternativknapp 6").Value = Sheets("PROGRAM�VERSIKT").OptionButton2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Sheets("PROGRAM�VERSIKT").OptionButton3.Value
    
    If Sheets("PROGRAM�VERSIKT").OptionButton1.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAM�VERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAM�VERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAM�VERSIKT").TextBox8.Value
    ElseIf Sheets("PROGRAM�VERSIKT").OptionButton2.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAM�VERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAM�VERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAM�VERSIKT").TextBox8.Value
        Sheets("Meny").[F21] = Sheets("PROGRAM�VERSIKT").TextBox5.Value
        Sheets("Meny").[G21] = Sheets("PROGRAM�VERSIKT").TextBox6.Value
        Sheets("Meny").[H21] = Sheets("PROGRAM�VERSIKT").TextBox7.Value
    End If
        Call ok
End Sub

Sub Tr�korg()

If FileFolderExists(KonfFor & OrderNrText & ".FOR") Then
    Sheets("Meny").[B3] = OrderNrText
    Sheets("Meny").[B16] = Anm�rkning
    Sheets("Meny").CheckBoxes("Kryssruta 4").Value = Sheets("PROGRAM�VERSIKT").CheckBox1.Value
    Sheets("Meny").CheckBoxes("Print").Value = Sheets("PROGRAM�VERSIKT").CheckBox2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 5").Value = Sheets("PROGRAM�VERSIKT").OptionButton1.Value
    Sheets("Meny").OptionButtons("Alternativknapp 6").Value = Sheets("PROGRAM�VERSIKT").OptionButton2.Value
    Sheets("Meny").OptionButtons("Alternativknapp 7").Value = Sheets("PROGRAM�VERSIKT").OptionButton3.Value
    
    If Sheets("PROGRAM�VERSIKT").OptionButton1.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAM�VERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAM�VERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAM�VERSIKT").TextBox8.Value
    ElseIf Sheets("PROGRAM�VERSIKT").OptionButton2.Value = True Then
        Sheets("Meny").[F33] = Sheets("PROGRAM�VERSIKT").TextBox10.Value
        Sheets("Meny").[G33] = Sheets("PROGRAM�VERSIKT").TextBox9.Value
        Sheets("Meny").[H33] = Sheets("PROGRAM�VERSIKT").TextBox8.Value
        Sheets("Meny").[F21] = Sheets("PROGRAM�VERSIKT").TextBox5.Value
        Sheets("Meny").[G21] = Sheets("PROGRAM�VERSIKT").TextBox6.Value
        Sheets("Meny").[H21] = Sheets("PROGRAM�VERSIKT").TextBox7.Value
    End If
    Call �ppnaFOR
    Call ok
    Exit Sub
    
Errmsg:
    MsgBox ("Misslyckades att k�ra �ppnaFOR"), vbOKOnly, ".FOR-fil"
    Else
        MsgBox ("Ordernumret finns inte.")
End If

End Sub


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'-------------KONFIGURATION------------'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
Sub �ppnaKatalogDialog()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAM�VERSIKT").TextBox11.Text = .SelectedItems(1)
    End With
End Sub
Sub �ppnaKatalogDialog2()
        With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAM�VERSIKT").TextBox12.Text = .SelectedItems(1)
    End With
End Sub
Sub �ppnaKatalogDialog3()
            With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then MsgBox "Ingen katalog vald.": Exit Sub
        Sheets("PROGRAM�VERSIKT").TextBox13.Text = .SelectedItems(1)
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
