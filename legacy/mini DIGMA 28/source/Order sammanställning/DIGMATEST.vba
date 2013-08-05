VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DIGMATEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Kör_Click()

    'Kör makro "ÖppnaFOR"
    'On Error GoTo Errmsg
    Call ÖppnaFOR
    Exit Sub
    
Errmsg:
    MsgBox ("Misslyckades att köra ÖppnaFOR"), vbOKOnly, ".FOR-fil"
End Sub
