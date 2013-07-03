VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub OptionButton1_Click()
    Application.ScreenUpdating = False
    ActiveSheet.Shapes.Range(Array("Ritning2")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Ritning1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox6")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox7")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox8")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox9")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox10")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox5")).Visible = msoFalse
    Sheets("Utskrift").Activate
    ActiveSheet.Shapes.Range(Array("bild_1_ing_sid1", "bild_1_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("bild_2_ing_sid1", "bild_2_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("bild_spec_sid1", "bild_spec_sid2")).Select
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.Font.ColorIndex = 2
    Selection.ShapeRange.Fill.Visible = msoFalse
    Sheets("PROGRAMÖVERSIKT").Activate
    Application.ScreenUpdating = True
End Sub

Private Sub OptionButton2_Click()
    Application.ScreenUpdating = False
    ActiveSheet.Shapes.Range(Array("Ritning2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Ritning1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox6")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox7")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox8")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox9")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox10")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("TextBox5")).Visible = msoTrue
    Sheets("Utskrift").Activate
    ActiveSheet.Shapes.Range(Array("bild_2_ing_sid1", "bild_2_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("bild_1_ing_sid1", "bild_1_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("bild_spec_sid1", "bild_spec_sid2")).Select
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.Font.ColorIndex = 2
    Selection.ShapeRange.Fill.Visible = msoFalse
    Sheets("PROGRAMÖVERSIKT").Activate
    Application.ScreenUpdating = True
End Sub

Private Sub OptionButton3_Click()
    Application.ScreenUpdating = False
    ActiveSheet.Shapes.Range(Array("Ritning2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Ritning1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox6")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox7")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox8")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox9")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox10")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("TextBox5")).Visible = msoFalse
    Sheets("Utskrift").Activate
    ActiveSheet.Shapes.Range(Array("bild_2_ing_sid1", "bild_2_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("bild_1_ing_sid1", "bild_1_ing_sid2")).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("bild_spec_sid1", "bild_spec_sid2")).Select
    Selection.ShapeRange.ZOrder msoBringToFront
    Selection.Font.ColorIndex = 0
    Selection.ShapeRange.Fill.Visible = msoTrue
    Selection.ShapeRange.Fill.Solid
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    Sheets("PROGRAMÖVERSIKT").Activate
    Application.ScreenUpdating = True
End Sub

Private Sub TextBox1_Change()
    Sheets("Meny").[B3] = Sheets("PROGRAMÖVERSIKT").TextBox1.Text
End Sub

Private Sub TextBox4_Change()
    Sheets("Meny").[B16] = Sheets("PROGRAMÖVERSIKT").TextBox4.Text
End Sub
