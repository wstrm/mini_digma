Attribute VB_Name = "LEGACY"
Sub Inmatning_1_ing()
'
' Inmatning_1_ing Makro
' Makrot gjort 2010-12-03 av AN
'

   Application.ScreenUpdating = False
   
   ' Ser till att rätt bild visas på utskriftssida
   
   Sheets("Utskrift").Select
   
   ActiveSheet.Shapes.Range(Array("bild_1_ing_sid1", "bild_1_ing_sid2")).Select
   Selection.ShapeRange.Line.Visible = msoTrue
   ActiveSheet.Shapes.Range(Array("bild_2_ing_sid1", "bild_2_ing_sid2")).Select
   Selection.ShapeRange.Line.Visible = msoFalse
   ActiveSheet.Shapes.Range(Array("bild_spec_sid1", "bild_spec_sid2")).Select
   Selection.ShapeRange.ZOrder msoSendToBack
   Selection.Font.ColorIndex = 2
   Selection.ShapeRange.Fill.Visible = msoFalse
   ActiveCell.Select
   
   Sheets("Meny").Select
    
    ' Uppdaterar inmatningsformuläret
    
    Range("F33:H33").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    Range("G24").Select
    ActiveCell.FormulaR1C1 = "Rygg"
    Range("G30").Select
    ActiveCell.FormulaR1C1 = "Fram"
    
    Range("F21:H21").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
       
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    Range("F24:H30").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Range("G31").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Range("G23").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
        
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("F33").Select
        
End Sub
Sub Inmatning_2_ing()
'
' Inmatning_2_ing Makro
' Makrot gjort 2010-12-03 av AN


   Application.ScreenUpdating = False
   
   ' Ser till att rätt bild visas på utskriftssida
   
   Sheets("Utskrift").Select
   
   ActiveSheet.Shapes.Range(Array("bild_2_ing_sid1", "bild_2_ing_sid2")).Select
   Selection.ShapeRange.Line.Visible = msoTrue
   ActiveSheet.Shapes.Range(Array("bild_1_ing_sid1", "bild_1_ing_sid2")).Select
   Selection.ShapeRange.Line.Visible = msoFalse
   ActiveSheet.Shapes.Range(Array("bild_spec_sid1", "bild_spec_sid2")).Select
   Selection.ShapeRange.ZOrder msoSendToBack
   Selection.Font.ColorIndex = 2
   Selection.ShapeRange.Fill.Visible = msoFalse
   ActiveCell.Select
   
   Sheets("Meny").Select
    
    ' Uppdaterar inmatningsformuläret
    
    Range("F33:H33").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    Range("G24").Select
    ActiveCell.FormulaR1C1 = "Rygg"
    Range("G30").Select
    ActiveCell.FormulaR1C1 = "Fram"
    
      
    Range("F24:H30").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        
     Range("G31").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        
    Range("F21:H21").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    
    Range("G23").Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSolid
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
   
    Range("F33").Select

End Sub

Sub Inmatning_special()
'
' Inmatning_special Makro
' Makrot gjort 2010-12-03 av AN


   Application.ScreenUpdating = False
   
   ' Ser till att rätt bild visas på utskriftssida
   
   Sheets("Utskrift").Select
   
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
   ActiveCell.Select
   
   Sheets("Meny").Select
    
    ' Uppdaterar inmatningsformuläret
    
    Range("F21:H33").Select
    Selection.ClearContents
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    Range("D21").Select

  
End Sub

