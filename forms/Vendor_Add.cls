Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CloseformBtn_Click()
On Error GoTo Err_CloseFormBtn_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.close

Exit_CloseFormBtn_Click:
    Exit Sub

Err_CloseFormBtn_Click:
    MsgBox Err.Description
    Resume Exit_CloseFormBtn_Click

End Sub
