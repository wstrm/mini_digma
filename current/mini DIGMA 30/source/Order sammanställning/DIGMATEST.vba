VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DIGMATEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub K�r_Click()

    'K�r makro "�ppnaFOR"
    'On Error GoTo Errmsg
    Call �ppnaFOR
    Exit Sub
    
Errmsg:
    MsgBox ("Misslyckades att k�ra �ppnaFOR"), vbOKOnly, ".FOR-fil"
End Sub
